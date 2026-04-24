"""
Excel 上传与分析 API：读取表格、返回预览与基础统计，支持自动入库和历史记录管理。
"""

from __future__ import annotations

import io
import json
import logging
import os
import re
from datetime import datetime, date as date_type
from pathlib import Path
from typing import Any

import pandas as pd
from dotenv import load_dotenv
from psycopg import connect
from psycopg import sql

from fastapi import FastAPI, File, HTTPException, UploadFile, Query
from fastapi.responses import HTMLResponse, JSONResponse
from fastapi.staticfiles import StaticFiles

# 导入本地模块
from models import (
    DEFAULT_WEIGHTS, DEFAULT_COLUMN_ALIASES, WEIGHT_LABELS,
    SaveCustomModeRequest, DeleteCustomModeRequest,
    SaveColumnMappingRequest, DeleteColumnMappingRequest,
)
from utils import (
    MAX_PREVIEW_ROWS, MAX_SAVE_ROWS,
    UPLOAD_SESSIONS_TABLE, UPLOAD_DATA_TABLE, COLUMN_MAPPING_TABLE,
    normalize_col_name, find_col, to_float, parse_json_value,
    slugify_mode_name, infer_sql_type, normalize_cell_for_insert,
    dataframe_to_preview, safe_cell_value,
    parse_sheet_name_to_date, parse_all_sheet_dates, get_date_display_order,
)

BASE_DIR = Path(__file__).resolve().parent
load_dotenv(BASE_DIR / ".env")

# 配置日志系统
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    handlers=[
        logging.StreamHandler(),
    ]
)
logger = logging.getLogger("excel_analyzer")

# 文件上传限制配置
MAX_FILE_SIZE = 50 * 1024 * 1024  # 50MB
MAX_FILE_SIZE_MB = 50
ALLOWED_EXTENSIONS = {".xlsx", ".xls"}

app = FastAPI(title="Excel 分析")


def get_pg_dsn() -> str:
    """获取PostgreSQL连接串，如果未配置则抛出异常"""
    dsn = os.getenv("PG_DSN", "").strip()
    if not dsn:
        raise HTTPException(status_code=500, detail="缺少 PG_DSN 环境变量，请先配置 PostgreSQL 连接串")
    return dsn


def _ensure_upload_tables_exist(conn) -> None:
    """确保上传历史表、数据表和列映射配置表存在（委托给 db.schema 模块）。"""
    from db.schema import ensure_upload_tables_exist
    ensure_upload_tables_exist(conn)


def _save_upload_to_db(
    conn,
    filename: str,
    rows: list[dict[str, Any]],
    columns: list[str],
    has_workload_analysis: bool,
    column_mapping_id: int | None = None,
    sheet_name: str | None = None,
    selected_columns: str | None = None,
    display_names: dict[str, str] | None = None,
    column_types: dict[str, str] | None = None,
    chart_types: dict[str, str] | None = None,
    config_name: str | None = None,
) -> int:
    """保存上传数据到数据库，返回 session_id。使用批量插入优化性能。"""
    today = datetime.now().date()
    row_count = len(rows)
    col_count = len(columns)

    with conn.cursor() as cur:
        # 插入会话记录（任何导入的数据都标记为有分析）
        cur.execute(
            sql.SQL("""
            INSERT INTO {} (upload_date, filename, row_count, col_count, columns_json, has_workload_analysis, has_analysis, column_mapping_id, sheet_name, selected_columns, display_names, column_types, chart_types, config_name)
            VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
            RETURNING id
            """).format(sql.Identifier(UPLOAD_SESSIONS_TABLE)),
            (today, filename, row_count, col_count, json.dumps(columns), has_workload_analysis, True, column_mapping_id, sheet_name, selected_columns, json.dumps(display_names or {}), json.dumps(column_types or {}), json.dumps(chart_types or {}), config_name)
        )
        session_id = cur.fetchone()[0]

        # 使用批量插入优化（每批1000行）
        batch_size = 1000
        table_name = sql.Identifier(UPLOAD_DATA_TABLE)
        
        # 准备批量数据并执行插入
        for batch_start in range(0, len(rows), batch_size):
            batch_end = min(batch_start + batch_size, len(rows))
            for idx in range(batch_start, batch_end):
                cur.execute(
                    sql.SQL("INSERT INTO {} (session_id, row_index, row_data) VALUES (%s, %s, %s)").format(table_name),
                    (session_id, idx, json.dumps(rows[idx]))
                )
        
        conn.commit()

    return session_id


def _build_person_risk(person: dict[str, Any]) -> dict[str, Any]:
    risk_score = person["escalation_help"] * 1.2 + person["pending_ticket"] * 0.7 + person["oncall_open"] * 0.65
    encourage_score = (
        person["issue_ticket_output"] * 1.0
        + person["requirement_ticket_output"] * 0.9
        + person["wiki_output"] * 1.2
        + person["analysis_report_output"] * 1.1
    )
    level = "low"
    if risk_score >= 18:
        level = "high"
    elif risk_score >= 10:
        level = "medium"

    suggestions: list[str] = []
    if person["escalation_help"] >= 4:
        suggestions.append("透传求助偏高，建议前移排障与值守经验复盘。")
    if person["kernel_issue"] >= person["governance_issue"] + person["consult_issue"]:
        suggestions.append("内核问题占比高，建议安排专项根因分析与变更演练。")
    if encourage_score <= 4:
        suggestions.append("沉淀产出偏低，建议补充 wiki 与问题分析报告。")
    if not suggestions:
        suggestions.append("当前风险可控，建议持续保持知识沉淀与问题闭环。")

    return {
        "risk_score": round(risk_score, 2),
        "risk_level": level,
        "encourage_score": round(encourage_score, 2),
        "suggestions": suggestions,
    }


def _build_risk_predictions(people: list[dict[str, Any]], is_multi_date: bool, total_count: int, full_attendance: int) -> dict[str, Any]:
    """
    构建智能风险预测报告，基于多日数据分析个人和团队风险点。
    
    Args:
        people: 人员数据列表
        is_multi_date: 是否为多日汇总数据
        total_count: 总人数
        full_attendance: 全勤人数
    
    Returns:
        {
            "high_pressure_people": 高压人员列表及原因,
            "escalation_concerns": 透传求助异常人员,
            "low_productivity_people": 产出偏低人员,
            "attendance_issues": 出勤异常人员,
            "kernel_focus_people": 内核问题集中人员,
            "summary": 风险预测摘要
        }
    """
    predictions = {
        "high_pressure_people": [],
        "escalation_concerns": [],
        "low_productivity_people": [],
        "attendance_issues": [],
        "kernel_focus_people": [],
        "summary": [],
    }
    
    # 计算团队平均值（用于对比判断）
    avg_workload = sum(p.get("workload_score", 0) for p in people) / len(people) if people else 0
    avg_escalation = sum(p.get("escalation_help", 0) for p in people) / len(people) if people else 0
    avg_output = sum(p.get("issue_ticket_output", 0) + p.get("requirement_ticket_output", 0) + p.get("wiki_output", 0) + p.get("analysis_report_output", 0) for p in people) / len(people) if people else 0
    
    for p in people:
        name = p.get("name", "")
        workload = p.get("workload_score", 0)
        escalation = p.get("escalation_help", 0)
        output_total = p.get("issue_ticket_output", 0) + p.get("requirement_ticket_output", 0) + p.get("wiki_output", 0) + p.get("analysis_report_output", 0)
        kernel = p.get("kernel_issue", 0)
        date_count = p.get("date_count", 1)
        
        # 1. 高压人员：工作量分超过平均值1.5倍
        if workload > avg_workload * 1.5 and workload > 30:
            reason = f"工作量分{round(workload, 1)}（团队平均{round(avg_workload, 1)}），压力过大"
            predictions["high_pressure_people"].append({"name": name, "workload": workload, "reason": reason})
        
        # 2. 透传求助异常：透传高于平均值2倍，且产出较低
        if escalation > avg_escalation * 2 and escalation > 1 and output_total < avg_output * 0.5:
            reason = f"透传求助{round(escalation, 1)}次（团队平均{round(avg_escalation, 1)}），但产出仅{round(output_total, 1)}"
            predictions["escalation_concerns"].append({
                "name": name, 
                "escalation": escalation, 
                "output": output_total,
                "reason": reason,
                "suggestion": "透传较多但独立处理问题数量较少，建议加强技能培训和问题闭环能力"
            })
        
        # 3. 产出偏低：总产出低于平均值0.3倍
        if output_total < avg_output * 0.3 and output_total < 3:
            predictions["low_productivity_people"].append({"name": name, "output": output_total})
        
        # 4. 出勤异常：仅适用于多日汇总数据
        if is_multi_date and date_count < 3:  # 出现少于3天认为出勤异常
            predictions["attendance_issues"].append({
                "name": name,
                "date_count": date_count,
                "suggestion": f"出现天数{int(date_count)}天，请假/缺席次数较多，需关注工作安排"
            })
        
        # 5. 内核问题集中：内核问题占比超过50%
        daily_issue = p.get("daily_issue_total", 0)
        if kernel > 0 and daily_issue > 0 and kernel / daily_issue > 0.5:
            predictions["kernel_focus_people"].append({
                "name": name,
                "kernel_ratio": round(kernel / daily_issue * 100, 1),
                "suggestion": "内核问题处理集中，建议专项根因分析"
            })
    
    # 构建摘要
    if predictions["high_pressure_people"]:
        predictions["summary"].append(f"⚠️ {len(predictions['high_pressure_people'])}人工作量压力过大，建议合理分配任务")
    if predictions["escalation_concerns"]:
        predictions["summary"].append(f"⚠️ {len(predictions['escalation_concerns'])}人透传求助异常多但产出较低，需关注技能成长")
    if predictions["attendance_issues"]:
        predictions["summary"].append(f"📊 {len(predictions['attendance_issues'])}人出勤天数较少，请关注工作连续性")
    if predictions["kernel_focus_people"]:
        predictions["summary"].append(f"🔧 {len(predictions['kernel_focus_people'])}人内核问题占比高，建议专项攻关")
    if predictions["low_productivity_people"]:
        predictions["summary"].append(f"📝 {len(predictions['low_productivity_people'])}人产出偏低，建议加强激励")
    
    if not predictions["summary"]:
        predictions["summary"].append("✅ 团队整体风险可控，各项指标分布合理")
    
    return predictions


def _build_team_suggestions(people: list[dict[str, Any]], totals: dict[str, Any], is_multi_date: bool) -> list[str]:
    """
    构建团队层面的建议，基于整体数据分析。
    
    Args:
        people: 人员数据列表
        totals: 总量数据
        is_multi_date: 是否为多日汇总数据
    
    Returns:
        建议列表
    """
    suggestions = []
    
    # 分析团队整体情况
    high_risk_count = sum(1 for p in people if p.get("risk_level") == "high")
    total_people = len(people)
    
    # 透传求助比例
    if totals.get("escalation_help", 0) > 10:
        avg_escalation = totals["escalation_help"] / total_people if total_people > 0 else 0
        suggestions.append(f"团队透传求助总量较高（{totals['escalation_help']}次，人均{round(avg_escalation, 1)}次），建议组织排障经验分享会。")
    
    # 内核问题比例
    kernel_ratio = totals.get("kernel_issue", 0) / (totals.get("daily_issue_total", 1) or 1)
    if kernel_ratio > 0.3:
        suggestions.append(f"内核问题占比{round(kernel_ratio * 100, 1)}%，建议专项跟进内核问题根因分析。")
    
    # 知识沉淀
    wiki_total = totals.get("wiki_output", 0) + totals.get("analysis_report_output", 0)
    if wiki_total < 5:
        suggestions.append("团队知识沉淀产出偏低，建议设立wiki产出激励机制。")
    
    # 高风险人员比例
    if high_risk_count > 0 and total_people > 0:
        high_ratio = high_risk_count / total_people
        if high_ratio > 0.2:
            suggestions.append(f"高风险人员占比{round(high_ratio * 100, 1)}%，建议重点关注并合理调整工作分配。")
    
    # 多日数据特有建议
    if is_multi_date:
        # 分析数据波动趋势
        suggestions.append("已按多日数据汇总分析，工作量分基于日均数据计算，更公平反映实际负荷。")
    
    if not suggestions:
        suggestions.append("团队各项指标健康，建议持续保持当前工作节奏和知识沉淀习惯。")
    
    return suggestions


def _get_column_aliases_from_config(mapping_config: dict[str, Any] | None) -> dict[str, list[str]]:
    """从配置中提取列别名映射，如果无配置则使用默认值。"""
    if mapping_config is None:
        return DEFAULT_COLUMN_ALIASES.copy()
    
    aliases = {}
    for key in DEFAULT_COLUMN_ALIASES.keys():
        config_key = f"{key}_aliases"
        if config_key in mapping_config:
            aliases[key] = parse_json_value(mapping_config[config_key]) or []
        else:
            aliases[key] = DEFAULT_COLUMN_ALIASES[key]
    return aliases


def _build_workload_analysis(
    df: pd.DataFrame, 
    column_aliases: dict[str, list[str]] | None = None,
    is_multi_date_summary: bool = False
) -> dict[str, Any] | None:
    """构建工作量分析，支持自定义列别名和多日汇总模式。
    
    Args:
        df: 数据DataFrame
        column_aliases: 列别名映射
        is_multi_date_summary: 是否为多日汇总数据（使用平均值列计算工作量分）
    """
    cols = [str(c) for c in df.columns.tolist()]
    
    # 使用传入的别名或默认别名
    if column_aliases is None:
        column_aliases = DEFAULT_COLUMN_ALIASES.copy()
    
    name_col = find_col(cols, column_aliases.get("name", ["姓名"]))
    if not name_col:
        return None
    
    # 检查是否存在"出现天数"列，判断是否为多日汇总数据
    date_count_col = find_col(cols, ["出现天数"])
    if date_count_col and not is_multi_date_summary:
        is_multi_date_summary = True
    
    # 对于多日汇总数据，优先使用"平均值"列来计算工作量分
    # 这样可以公平评估每个人的日均工作量，避免天数多的用户工作量分偏高
    col_map = {}
    metric_keys = [
        "oncall_open", "pending_ticket", "new_issue_yesterday",
        "governance_issue", "kernel_issue", "consult_issue", "escalation_help",
        "issue_ticket_output", "requirement_ticket_output", "wiki_output", "analysis_report_output"
    ]
    
    for key in metric_keys:
        # 对于多日汇总数据，优先查找平均值列
        avg_col_name = None
        if is_multi_date_summary:
            base_aliases = column_aliases.get(key, [])
            # 构建平均值列的别名（如 "oncall接单未闭环的数量_平均"）
            avg_aliases = [f"{alias}_平均" for alias in base_aliases] + [f"{alias} 平均" for alias in base_aliases]
            avg_aliases.extend([f"{key}_平均", f"{key} 平均"])
            avg_col_name = find_col(cols, avg_aliases)
        
        if avg_col_name:
            # 使用平均值列
            col_map[key] = avg_col_name
        else:
            # 使用原始列或总和列
            col_map[key] = find_col(cols, column_aliases.get(key, []))

    # 检查是否找到了足够的列（至少需要姓名列和一些数值列）
    found_cols = sum(1 for v in col_map.values() if v)
    if found_cols < 3:
        return None

    weights = DEFAULT_WEIGHTS.copy()

    people: list[dict[str, Any]] = []
    for row in df.fillna("").to_dict(orient="records"):
        person = str(row.get(name_col, "")).strip()
        if not person:
            continue
        
        # 获取指标值
        metrics = {k: to_float(row.get(v, 0)) if v else 0.0 for k, v in col_map.items()}
        daily_issue_total = metrics["governance_issue"] + metrics["kernel_issue"] + metrics["consult_issue"]
        score = sum(metrics[k] * w for k, w in weights.items())
        
        # 对于多日汇总数据，工作量分基于平均值，更合理
        item = {
            "name": person,
            **{k: round(v, 2) for k, v in metrics.items()},
            "daily_issue_total": round(daily_issue_total, 2),
            "workload_score": round(score, 2),
        }
        
        # 如果有出现天数列，添加到结果中
        if date_count_col:
            item["date_count"] = to_float(row.get(date_count_col, 1))
        
        item.update(_build_person_risk(item))
        people.append(item)

    if not people:
        return None

    by_score = sorted(people, key=lambda x: x["workload_score"], reverse=True)
    by_escalation = sorted(people, key=lambda x: x["escalation_help"], reverse=True)
    by_risk = sorted(people, key=lambda x: x["risk_score"], reverse=True)
    
    # 计算总量（注意：对于多日汇总数据，这些是基于平均值的总量）
    totals = {k: round(sum(p[k] for p in people), 2) for k in weights}
    totals["daily_issue_total"] = round(sum(p["daily_issue_total"] for p in people), 2)
    totals["risk_score"] = round(sum(p["risk_score"] for p in people), 2)
    totals["encourage_score"] = round(sum(p["encourage_score"] for p in people), 2)
    totals["gaussdb_focus_index"] = round(
        totals["new_issue_yesterday"] * 1.3
        + totals["kernel_issue"] * 1.45
        + totals["governance_issue"] * 1.05
        + totals["consult_issue"] * 0.9
        - totals["escalation_help"] * 0.7
        + totals["wiki_output"] * 1.25
        + totals["analysis_report_output"] * 1.15,
        2,
    )
    
    # 计算人数统计
    total_people_count = len(people)
    full_attendance_count = sum(1 for p in people if p.get("date_count", 1) >= 2) if date_count_col else total_people_count
    
    risk_level_counts = {"high": 0, "medium": 0, "low": 0}
    for p in people:
        risk_level_counts[p["risk_level"]] += 1

    return {
        "weights": weights,
        "col_map": col_map,
        "people": by_score,
        "transparent_ranking": [{"name": p["name"], "escalation_help": p["escalation_help"]} for p in by_escalation],
        "totals": totals,
        "top_score_names": [p["name"] for p in by_score[:3]],
        "top_transparent_names": [p["name"] for p in by_escalation[:3]],
        "high_risk_names": [p["name"] for p in by_risk if p["risk_level"] == "high"][:5],
        "risk_level_counts": risk_level_counts,
        "is_multi_date_summary": is_multi_date_summary,  # 标记是否为多日汇总数据
        "total_people_count": total_people_count,  # 总人数
        "full_attendance_count": full_attendance_count,  # 全勤人数
        # 新增：智能风险预测和团队建议
        "risk_predictions": _build_risk_predictions(people, is_multi_date_summary, total_people_count, full_attendance_count),
        "team_suggestions": _build_team_suggestions(people, totals, is_multi_date_summary),
        "gaussdb_online_focus": {
            "description": "聚焦现网问题承压、内核问题处理与知识沉淀能力，透传求助作为负向项。",
            "index": totals["gaussdb_focus_index"],
            "highlights": [
                f"内核问题总量: {totals['kernel_issue']}",
                f"昨日新增问题总量: {totals['new_issue_yesterday']}",
                f"wiki+分析报告产出: {round(totals['wiki_output'] + totals['analysis_report_output'], 2)}",
                f"透传求助总量: {totals['escalation_help']}",
            ],
        },
    }


def _dataframe_to_payload(
    df: pd.DataFrame, 
    column_aliases: dict[str, list[str]] | None = None
) -> dict[str, Any]:
    df = df.fillna("")
    preview = df.head(MAX_PREVIEW_ROWS)
    rows = json.loads(preview.to_json(orient="records", force_ascii=False))
    numeric = df.select_dtypes(include=["number"])
    describe = None
    if not numeric.empty:
        desc = numeric.describe().round(4)
        describe = json.loads(desc.to_json(orient="index", force_ascii=False))

    dtypes = {str(c): str(t) for c, t in df.dtypes.items()}
    missing = df.isna().sum().astype(int).to_dict()
    missing = {str(k): int(v) for k, v in missing.items()}

    return {
        "shape": {"rows": int(df.shape[0]), "cols": int(df.shape[1])},
        "columns": [str(c) for c in df.columns.tolist()],
        "dtypes": dtypes,
        "missing_counts": missing,
        "preview_rows": rows,
        "preview_truncated": len(df) > MAX_PREVIEW_ROWS,
        "numeric_describe": describe,
        "workload_analysis": _build_workload_analysis(df, column_aliases),
        "column_aliases": column_aliases or DEFAULT_COLUMN_ALIASES,
    }


@app.post("/api/upload/preview")
async def preview_excel(file: UploadFile = File(...)) -> JSONResponse:
    """预览Excel文件，返回所有sheet的信息。"""
    # 文件名安全检查
    name = (file.filename or "").lower()
    if not name.endswith((".xlsx", ".xls")):
        raise HTTPException(
            status_code=400,
            detail="请上传 .xlsx 或 .xls 文件",
        )
    
    # 文件名安全处理（防止路径遍历）
    safe_filename = Path(file.filename or "unknown.xlsx").name
    
    # 检查文件大小
    raw = await file.read()
    if not raw:
        raise HTTPException(status_code=400, detail="文件为空")
    if len(raw) > MAX_FILE_SIZE:
        raise HTTPException(
            status_code=400,
            detail=f"文件大小超过限制（最大 {MAX_FILE_SIZE_MB}MB）",
        )

    try:
        bio = io.BytesIO(raw)
        engine = "openpyxl" if name.endswith(".xlsx") else "xlrd"
        # 读取所有sheet
        xls = pd.ExcelFile(bio, engine=engine)
        sheet_names = xls.sheet_names
        
        sheets_info = []
        # 解析所有sheet名称为日期
        date_info_list = parse_all_sheet_dates(sheet_names)
        
        for idx, sheet_name in enumerate(sheet_names):
            df = pd.read_excel(xls, sheet_name=sheet_name)
            
            # 应用 safe_cell_value 处理所有单元格，兼容空值和公式错误
            for col in df.columns:
                df[col] = df[col].apply(safe_cell_value)
            
            # 替换 NaN 为空字符串
            df = df.fillna("")
            
            # 获取列信息
            columns = [str(c) for c in df.columns.tolist()]
            dtypes = {str(c): str(t) for c, t in df.dtypes.items()}
            
            # 判断每列是否为数值型
            column_types = {}
            for col in columns:
                dtype = dtypes.get(col, "")
                if "int" in dtype or "float" in dtype or "number" in dtype:
                    column_types[col] = "numeric"
                elif "datetime" in dtype or "date" in dtype:
                    column_types[col] = "datetime"
                else:
                    column_types[col] = "text"
            
            # 获取前几行预览
            preview_rows = json.loads(df.head(5).to_json(orient="records", force_ascii=False))
            
            # 获取日期识别信息
            date_info = date_info_list[idx] if idx < len(date_info_list) else {}
            
            sheets_info.append({
                "sheet_name": sheet_name,
                "row_count": int(df.shape[0]),
                "col_count": int(df.shape[1]),
                "columns": columns,
                "column_types": column_types,
                "preview_rows": preview_rows,
                "has_workload_analysis": _build_workload_analysis(df) is not None,
                # 新增日期识别信息
                "date_info": {
                    "is_date": date_info.get("is_date", False),
                    "parsed_date": date_info.get("parsed_date"),
                    "display_name": date_info.get("display_name", sheet_name),
                    "format_type": date_info.get("format_type"),
                },
            })
        
        # 如果有多个sheet且列结构一致，生成汇总数据
        merged_info = None
        if len(sheets_info) > 1:
            # 检查列结构是否一致
            first_columns = sheets_info[0]["columns"]
            columns_match = all(s["columns"] == first_columns for s in sheets_info)
            
            if columns_match:
                # 合并所有sheet数据（按姓名聚合，不依赖行位置）
                # 先找出姓名列和数值列
                name_col = None
                for col in first_columns:
                    if "姓名" in col or "名字" in col or col == "name":
                        name_col = col
                        break
                
                # 收集所有数据，按sheet和姓名组织
                all_data_by_person: dict[str, dict[str, Any]] = {}  # {person_name: {metric: [values]}}
                person_dates: dict[str, list[str]] = {}  # {person_name: [dates]}
                
                for sheet_name in sheet_names:
                    df_sheet = pd.read_excel(xls, sheet_name=sheet_name)
                    
                    if name_col and name_col in df_sheet.columns:
                        # 找出数值列（排除姓名列）
                        potential_numeric_cols = [col for col in df_sheet.columns if col != name_col]
                        
                        for idx, row in df_sheet.iterrows():
                            person = str(row.get(name_col, "")).strip()
                            if not person:
                                continue
                            
                            # 初始化人员数据结构
                            if person not in all_data_by_person:
                                all_data_by_person[person] = {}
                                person_dates[person] = []
                            
                            person_dates[person].append(sheet_name)
                            
                            # 收集每个数值列的值（空值按0处理）
                            for col in potential_numeric_cols:
                                value = row.get(col)
                                # 处理空值：NaN、None、空字符串都按0处理
                                if value is None or (isinstance(value, float) and pd.isna(value)):
                                    numeric_value = 0.0
                                elif isinstance(value, str) and value.strip() == "":
                                    numeric_value = 0.0
                                else:
                                    try:
                                        numeric_value = float(value)
                                    except (TypeError, ValueError):
                                        numeric_value = 0.0
                                
                                if col not in all_data_by_person[person]:
                                    all_data_by_person[person][col] = []
                                all_data_by_person[person][col].append(numeric_value)
                
                if name_col and all_data_by_person:
                    # 构建汇总DataFrame
                    summary_rows = []
                    for person in sorted(all_data_by_person.keys()):
                        row_data = {name_col: person, "出现天数": len(person_dates[person]), "出现日期": person_dates[person]}
                        
                        for col, values in all_data_by_person[person].items():
                            # 总和
                            row_data[col] = sum(values)
                            # 平均值（保留2位小数）
                            row_data[f"{col}_平均"] = round(sum(values) / len(values), 2)
                        
                        summary_rows.append(row_data)
                    
                    summary_df = pd.DataFrame(summary_rows)
                    
                    # 计算每日明细数据（原始数据，空值按0处理）
                    daily_details = {}
                    for sheet_name in sheet_names:
                        df_sheet = pd.read_excel(xls, sheet_name=sheet_name)
                        for col in df_sheet.columns:
                            df_sheet[col] = df_sheet[col].apply(safe_cell_value)
                        df_sheet = df_sheet.fillna(0)  # 数值列空值填0
                        df_sheet["_date_source"] = sheet_name
                        daily_details[sheet_name] = json.loads(df_sheet.head(20).to_json(orient="records", force_ascii=False))
                    
                    summary_preview = json.loads(summary_df.head(10).to_json(orient="records", force_ascii=False))
                    
                    # 数值列列表（排除姓名、出现天数、出现日期等）
                    numeric_cols = [col for col in summary_df.columns 
                                   if col != name_col and col != "出现天数" and col != "出现日期"
                                   and not col.endswith("_平均")]
                    
                    # 计算汇总后的工作量分析数据（使用平均值计算工作量分）
                    workload_preview = _build_workload_analysis(summary_df) if len(numeric_cols) >= 10 else None
                    
                    merged_info = {
                        "can_merge": True,
                        "has_workload_analysis": workload_preview is not None,
                        "total_rows": sum(len(pd.read_excel(xls, sheet_name=sn)) for sn in sheet_names),
                        "unique_people": len(summary_df),
                        "date_count": len(sheet_names),
                        "dates": sheet_names,
                        "summary_preview": summary_preview,
                        "summary_columns": list(summary_df.columns),
                        "daily_details": daily_details,
                        "numeric_columns": numeric_cols,
                        "date_coverage": {
                            "full_coverage": len(summary_df[summary_df["出现天数"] == len(sheet_names)]),
                            "partial_coverage": len(summary_df[summary_df["出现天数"] < len(sheet_names)]),
                        },
                        "workload_preview": workload_preview,
                        "message": f"检测到 {len(sheet_names)} 个日期数据，可自动汇总。共 {sum(len(pd.read_excel(xls, sheet_name=sn)) for sn in sheet_names)} 条记录，{len(summary_df)} 人（全勤 {len(summary_df[summary_df['出现天数'] == len(sheet_names)])} 人）。"
                    }
                else:
                    merged_info = {
                        "can_merge": True,
                        "total_rows": sum(len(pd.read_excel(xls, sheet_name=sn)) for sn in sheet_names),
                        "date_count": len(sheet_names),
                        "dates": sheet_names,
                        "message": f"检测到 {len(sheet_names)} 个sheet，可合并展示。"
                    }
            else:
                merged_info = {
                    "can_merge": False,
                    "message": "各sheet列结构不一致，无法自动汇总。"
                }
        
        logger.info(f"预览文件成功: {safe_filename}, {len(sheets_info)} 个sheet")
    except Exception as e:
        logger.error(f"解析 Excel 失败: {e}")
        raise HTTPException(
            status_code=422,
            detail=f"无法解析 Excel：{e!s}",
        ) from e

    return JSONResponse(content={
        "ok": True,
        "filename": safe_filename,
        "sheets": sheets_info,
        "sheet_count": len(sheets_info),
        "merged_info": merged_info,  # 汇总信息
        # 新增日期识别汇总
        "date_summary": {
            "total_sheets": len(sheet_names),
            "recognized_dates": sum(1 for d in date_info_list if d.get("is_date")),
            "unrecognized_sheets": [d.get("original_name") for d in date_info_list if not d.get("is_date")],
            "date_range": {
                "start": min([d.get("parsed_date") for d in date_info_list if d.get("parsed_date")]) if any(d.get("parsed_date") for d in date_info_list) else None,
                "end": max([d.get("parsed_date") for d in date_info_list if d.get("parsed_date")]) if any(d.get("parsed_date") for d in date_info_list) else None,
            },
            "sorted_dates": [d.get("parsed_date") for d in get_date_display_order(date_info_list) if d.get("parsed_date")],
        },
    })


@app.post("/api/upload")
async def upload_excel(
    file: UploadFile = File(...),
    column_mapping_id: int | None = Query(None, description="列映射配置ID，可选"),
    sheet_name: str | None = Query(None, description="指定sheet名称，可选"),
    selected_columns: str | None = Query(None, description="选择的列，逗号分隔"),
    display_names: str | None = Query(None, description="显示名称映射，格式 col:displayName"),
    column_types: str | None = Query(None, description="列数据类型，格式 col:type"),
    chart_types: str | None = Query(None, description="图表类型，格式 col:chartType"),
    config_name: str | None = Query(None, description="配置名称，用于后续加载"),
    merge_sheets: bool = Query(False, description="是否合并多个sheet的数据进行汇总"),
) -> JSONResponse:
    # 文件名安全检查
    name = (file.filename or "").lower()
    if not name.endswith((".xlsx", ".xls")):
        raise HTTPException(
            status_code=400,
            detail="请上传 .xlsx 或 .xls 文件",
        )
    
    # 文件名安全处理
    safe_filename = Path(file.filename or "unknown.xlsx").name
    
    # 检查文件大小
    raw = await file.read()
    if not raw:
        raise HTTPException(status_code=400, detail="文件为空")
    if len(raw) > MAX_FILE_SIZE:
        raise HTTPException(
            status_code=400,
            detail=f"文件大小超过限制（最大 {MAX_FILE_SIZE_MB}MB）",
        )

    try:
        bio = io.BytesIO(raw)
        engine = "openpyxl" if name.endswith(".xlsx") else "xlrd"
        
        # 支持多sheet汇总模式
        if merge_sheets and not sheet_name:
            xls = pd.ExcelFile(bio, engine=engine)
            sheet_names = xls.sheet_names
            
            if len(sheet_names) > 1:
                # 查找姓名列
                first_df = pd.read_excel(xls, sheet_name=sheet_names[0])
                name_col = None
                for col in first_df.columns:
                    if "姓名" in col or "名字" in col or col == "name":
                        name_col = col
                        break
                
                if name_col:
                    # 按姓名聚合，不依赖行位置，空值按0处理
                    all_data_by_person: dict[str, dict[str, Any]] = {}
                    person_dates: dict[str, list[str]] = {}
                    
                    for sn in sheet_names:
                        df_sheet = pd.read_excel(xls, sheet_name=sn)
                        
                        if name_col in df_sheet.columns:
                            potential_numeric_cols = [col for col in df_sheet.columns if col != name_col]
                            
                            for idx, row in df_sheet.iterrows():
                                person = str(row.get(name_col, "")).strip()
                                if not person:
                                    continue
                                
                                if person not in all_data_by_person:
                                    all_data_by_person[person] = {}
                                    person_dates[person] = []
                                
                                person_dates[person].append(sn)
                                
                                for col in potential_numeric_cols:
                                    value = row.get(col)
                                    # 空值按0处理
                                    if value is None or (isinstance(value, float) and pd.isna(value)):
                                        numeric_value = 0.0
                                    elif isinstance(value, str) and value.strip() == "":
                                        numeric_value = 0.0
                                    else:
                                        try:
                                            numeric_value = float(value)
                                        except (TypeError, ValueError):
                                            numeric_value = 0.0
                                    
                                    if col not in all_data_by_person[person]:
                                        all_data_by_person[person][col] = []
                                    all_data_by_person[person][col].append(numeric_value)
                    
                    # 构建汇总DataFrame
                    summary_rows = []
                    for person in sorted(all_data_by_person.keys()):
                        row_data = {name_col: person, "出现天数": len(person_dates[person]), "出现日期": person_dates[person]}
                        
                        for col, values in all_data_by_person[person].items():
                            row_data[col] = sum(values)
                            row_data[f"{col}_平均"] = round(sum(values) / len(values), 2)
                        
                        summary_rows.append(row_data)
                    
                    merged_df = pd.DataFrame(summary_rows)
                    logger.info(f"多sheet汇总完成: {len(sheet_names)}个sheet, 汇总后{len(merged_df)}人, 全勤{len([p for p in person_dates if len(person_dates[p]) == len(sheet_names)])}人")
                else:
                    # 无姓名列，简单合并
                    all_dfs = []
                    for sn in sheet_names:
                        df_sheet = pd.read_excel(xls, sheet_name=sn)
                        for col in df_sheet.columns:
                            df_sheet[col] = df_sheet[col].apply(safe_cell_value)
                        all_dfs.append(df_sheet)
                    merged_df = pd.concat(all_dfs, ignore_index=True)
                    logger.info(f"多sheet合并完成: {len(sheet_names)}个sheet, {len(merged_df)}行")
                
                df = merged_df.fillna(0)  # 数值列空值填0
            else:
                df = pd.read_excel(bio, engine=engine)
                for col in df.columns:
                    df[col] = df[col].apply(safe_cell_value)
                df = df.fillna(0)
        elif sheet_name:
            df = pd.read_excel(bio, engine=engine, sheet_name=sheet_name)
        else:
            df = pd.read_excel(bio, engine=engine)
    except Exception as e:
        raise HTTPException(
            status_code=422,
            detail=f"无法解析 Excel：{e!s}",
        ) from e
    
    # 支持选择指定列
    if selected_columns:
        cols_list = [c.strip() for c in selected_columns.split(",") if c.strip()]
        # 过滤出存在的列
        existing_cols = [c for c in cols_list if c in df.columns]
        if existing_cols:
            df = df[existing_cols]

    # 获取列映射配置
    column_aliases = None
    mapping_name = None
    dsn = os.getenv("PG_DSN", "").strip()
    
    if dsn and column_mapping_id:
        try:
            with connect(dsn) as conn:
                _ensure_upload_tables_exist(conn)
                with conn.cursor() as cur:
                    cur.execute(
                        sql.SQL("SELECT * FROM {} WHERE id = %s").format(
                            sql.Identifier(COLUMN_MAPPING_TABLE)
                        ),
                        (column_mapping_id,)
                    )
                    mapping_row = cur.fetchone()
                    if mapping_row:
                        mapping_config = {
                            "id": mapping_row[0],
                            "mapping_name": mapping_row[1],
                            "name_aliases": mapping_row[2],
                            "oncall_open_aliases": mapping_row[3],
                            "pending_ticket_aliases": mapping_row[4],
                            "new_issue_yesterday_aliases": mapping_row[5],
                            "governance_issue_aliases": mapping_row[6],
                            "kernel_issue_aliases": mapping_row[7],
                            "consult_issue_aliases": mapping_row[8],
                            "escalation_help_aliases": mapping_row[9],
                            "issue_ticket_output_aliases": mapping_row[10],
                            "requirement_ticket_output_aliases": mapping_row[11],
                            "wiki_output_aliases": mapping_row[12],
                            "analysis_report_output_aliases": mapping_row[13],
                        }
                        column_aliases = _get_column_aliases_from_config(mapping_config)
                        mapping_name = mapping_row[1]
        except Exception as e:
            logger.warning(f"列映射配置加载失败，使用默认配置: {e}")
    
    if column_aliases is None:
        column_aliases = DEFAULT_COLUMN_ALIASES.copy()
        mapping_name = "默认配置"

    payload = _dataframe_to_payload(df, column_aliases)
    
    # 应用 safe_cell_value 处理所有单元格，兼容空值和公式错误
    for col in df.columns:
        df[col] = df[col].apply(safe_cell_value)
    
    all_rows = json.loads(df.fillna("").to_json(orient="records", force_ascii=False))
    if len(all_rows) > MAX_SAVE_ROWS:
        all_rows = all_rows[:MAX_SAVE_ROWS]
    payload["all_rows"] = all_rows
    payload["all_rows_truncated"] = len(df) > MAX_SAVE_ROWS
    payload["column_mapping_id"] = column_mapping_id
    payload["column_mapping_name"] = mapping_name
    # 添加导入配置信息
    payload["sheet_name"] = sheet_name
    payload["selected_columns"] = selected_columns
    payload["display_names"] = display_names
    payload["column_types"] = column_types
    payload["chart_types"] = chart_types
    
    # 解析图表类型配置为字典格式，方便前端使用
    chart_type_config = {}
    if chart_types:
        for item in chart_types.split(","):
            if ":" in item:
                col, chart_type = item.split(":", 1)
                chart_type_config[col.strip()] = chart_type.strip()
    payload["chart_type_config"] = chart_type_config
    
    # 解析显示名称配置为字典格式
    display_name_config = {}
    if display_names:
        for item in display_names.split(","):
            if ":" in item:
                col, display_name = item.split(":", 1)
                display_name_config[col.strip()] = display_name.strip()
    payload["display_name_config"] = display_name_config

    # 自动保存到 PostgreSQL（如果配置了 PG_DSN）
    if dsn:
        columns = payload["columns"]
        has_workload = payload["workload_analysis"] is not None
        try:
            with connect(dsn) as conn:
                _ensure_upload_tables_exist(conn)
                
                # 解析column_types为字典
                column_type_config = {}
                if column_types:
                    for item in column_types.split(","):
                        if ":" in item:
                            col, col_type = item.split(":", 1)
                            column_type_config[col.strip()] = col_type.strip()
                
                session_id = _save_upload_to_db(
                    conn, safe_filename, all_rows, columns, has_workload,
                    column_mapping_id, sheet_name, selected_columns,
                    display_name_config, column_type_config, chart_type_config, config_name
                )
                payload["saved_to_db"] = True
                payload["session_id"] = session_id
        except Exception as e:
            # 入库失败不影响返回数据，记录警告日志
            logger.warning(f"入库失败: {e}")
            payload["saved_to_db"] = False
            payload["db_error"] = str(e)
    else:
        payload["saved_to_db"] = False
        payload["db_error"] = "未配置 PG_DSN"

    return JSONResponse(content=payload)


@app.get("/api/upload/history")
async def get_upload_history(
    start_date: date_type | None = Query(None, description="起始日期，格式 YYYY-MM-DD"),
    end_date: date_type | None = Query(None, description="结束日期，格式 YYYY-MM-DD"),
) -> JSONResponse:
    """获取历史上传记录，按日期分组，支持日期范围筛选。"""
    dsn = get_pg_dsn()
    try:
        with connect(dsn) as conn:
            with conn.cursor() as cur:
                # 构建带日期筛选的查询
                base_query = sql.SQL("""
                    SELECT 
                        id, upload_date, upload_time, filename, row_count, col_count,
                        has_workload_analysis, has_analysis, notes, sheet_name, config_name
                    FROM {}
                    WHERE 1=1
                """).format(sql.Identifier(UPLOAD_SESSIONS_TABLE))
                
                conditions = []
                params: list[Any] = []
                
                if start_date:
                    conditions.append(sql.SQL("upload_date >= %s"))
                    params.append(start_date)
                if end_date:
                    conditions.append(sql.SQL("upload_date <= %s"))
                    params.append(end_date)
                
                if conditions:
                    base_query = sql.SQL("{} AND {}").format(
                        base_query,
                        sql.SQL(" AND ").join(conditions)
                    )
                
                order_query = sql.SQL("{} ORDER BY upload_date DESC, upload_time DESC LIMIT 100").format(base_query)
                
                cur.execute(order_query, params)
                rows = cur.fetchall()
                # 按日期分组
                by_date: dict[str, list[dict[str, Any]]] = {}
                for r in rows:
                    date_str = r[1].isoformat() if r[1] else ""
                    if date_str not in by_date:
                        by_date[date_str] = []
                    by_date[date_str].append({
                        "session_id": r[0],
                        "upload_time": r[2].isoformat() if r[2] else None,
                        "filename": r[3],
                        "row_count": r[4],
                        "col_count": r[5],
                        "has_workload_analysis": r[6],
                        "has_analysis": r[7] if r[7] is not None else True,
                        "notes": r[8],
                        "sheet_name": r[9],
                        "config_name": r[10],
                    })
                # 转换为列表格式
                items = [{"date": date, "sessions": sessions} for date, sessions in by_date.items()]
                
                # 获取所有可用日期列表（用于日期选择器）
                cur.execute(
                    sql.SQL("SELECT DISTINCT upload_date FROM {} ORDER BY upload_date DESC LIMIT 365").format(
                        sql.Identifier(UPLOAD_SESSIONS_TABLE)
                    )
                )
                available_dates = [r[0].isoformat() for r in cur.fetchall() if r[0]]
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"读取上传历史失败：{e!s}") from e

    return JSONResponse(content={
        "ok": True,
        "items": items,
        "available_dates": available_dates,
        "start_date": start_date.isoformat() if start_date else None,
        "end_date": end_date.isoformat() if end_date else None,
    })


@app.get("/api/config/list")
async def get_saved_configs() -> JSONResponse:
    """获取已保存的数据配置列表（用于下拉加载）。"""
    dsn = get_pg_dsn()
    try:
        with connect(dsn) as conn:
            with conn.cursor() as cur:
                # 查询所有有配置名称的会话（distinct by config_name）
                cur.execute(
                    sql.SQL("""
                    SELECT DISTINCT config_name, display_names, column_types, chart_types, selected_columns, filename, upload_time
                    FROM {}
                    WHERE config_name IS NOT NULL AND config_name != ''
                    ORDER BY upload_time DESC
                    LIMIT 50
                    """).format(sql.Identifier(UPLOAD_SESSIONS_TABLE))
                )
                rows = cur.fetchall()
                
                configs = []
                for r in rows:
                    if r[0]:  # config_name存在
                        configs.append({
                            "config_name": r[0],
                            "display_names": parse_json_value(r[1]) or {},
                            "column_types": parse_json_value(r[2]) or {},
                            "chart_types": parse_json_value(r[3]) or {},
                            "selected_columns": r[4] or "",
                            "filename": r[5] or "",
                            "upload_time": r[6].isoformat() if r[6] else None,
                        })
                
                return JSONResponse(content={
                    "ok": True,
                    "configs": configs,
                })
    except Exception as e:
        return JSONResponse(content={
            "ok": False,
            "configs": [],
            "error": str(e),
        })


@app.get("/api/session/config/{session_id}")
async def get_session_config(session_id: int) -> JSONResponse:
    """获取指定会话的配置信息。"""
    dsn = get_pg_dsn()
    try:
        with connect(dsn) as conn:
            with conn.cursor() as cur:
                cur.execute(
                    sql.SQL("""
                    SELECT id, config_name, display_names, column_types, chart_types, selected_columns, sheet_name, filename
                    FROM {}
                    WHERE id = %s
                    """).format(sql.Identifier(UPLOAD_SESSIONS_TABLE)),
                    (session_id,)
                )
                row = cur.fetchone()
                if not row:
                    return JSONResponse(content={"ok": False, "error": "会话不存在"})
                
                return JSONResponse(content={
                    "ok": True,
                    "session_id": row[0],
                    "config_name": row[1] or "",
                    "display_names": parse_json_value(row[2]) or {},
                    "column_types": parse_json_value(row[3]) or {},
                    "chart_types": parse_json_value(row[4]) or {},
                    "selected_columns": row[5] or "",
                    "sheet_name": row[6] or "",
                    "filename": row[7] or "",
                })
    except Exception as e:
        return JSONResponse(content={"ok": False, "error": str(e)})


@app.get("/api/upload/latest")
async def get_latest_upload() -> JSONResponse:
    """获取最新一次上传的数据。"""
    dsn = get_pg_dsn()
    try:
        with connect(dsn) as conn:
            with conn.cursor() as cur:
                # 获取最新会话（包含完整配置字段）
                cur.execute(
                    sql.SQL("""
                    SELECT id, upload_date, upload_time, filename, row_count, col_count, columns_json, has_workload_analysis,
                           sheet_name, selected_columns, display_names, column_types, chart_types, config_name
                    FROM {}
                    ORDER BY upload_date DESC, upload_time DESC
                    LIMIT 1
                    """).format(sql.Identifier(UPLOAD_SESSIONS_TABLE))
                )
                session = cur.fetchone()
                if not session:
                    return JSONResponse(content={"ok": False, "message": "暂无历史数据"})

                session_id = session[0]
                columns = parse_json_value(session[6]) or []
                # 解析图表配置字段
                sheet_name = session[8] or ""
                selected_columns = session[9] or ""
                display_names_json = parse_json_value(session[10]) or {}
                column_types_json = parse_json_value(session[11]) or {}
                chart_types_json = parse_json_value(session[12]) or {}
                config_name = session[13] or ""

                # 获取该会话的所有数据行
                cur.execute(
                    sql.SQL("""
                    SELECT row_data
                    FROM {}
                    WHERE session_id = %s
                    ORDER BY row_index ASC
                    """).format(sql.Identifier(UPLOAD_DATA_TABLE)),
                    (session_id,)
                )
                raw_rows = cur.fetchall()
                data_rows = [parse_json_value(r[0]) for r in raw_rows if r[0]]

                if not data_rows:
                    return JSONResponse(content={"ok": False, "message": "会话数据为空"})

                # 构建返回 payload
                df = pd.DataFrame(data_rows)
                payload = _dataframe_to_payload(df)
                payload["all_rows"] = data_rows
                payload["all_rows_truncated"] = False
                payload["session_info"] = {
                    "session_id": session_id,
                    "upload_date": session[1].isoformat() if session[1] else None,
                    "upload_time": session[2].isoformat() if session[2] else None,
                    "filename": session[3],
                    "sheet_name": sheet_name,
                    "config_name": config_name,
                }
                # 添加图表配置（关键：前端需要这些字段渲染自定义图表）
                payload["chart_type_config"] = chart_types_json
                payload["display_name_config"] = display_names_json
                payload["column_type_config"] = column_types_json
                payload["selected_columns"] = selected_columns
                payload["sheet_name"] = sheet_name
                payload["ok"] = True
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"读取最新数据失败：{e!s}") from e

    return JSONResponse(content=payload)


@app.get("/api/upload/session/{session_id}")
async def get_upload_session(session_id: int) -> JSONResponse:
    """获取指定会话的数据。"""
    dsn = get_pg_dsn()
    try:
        with connect(dsn) as conn:
            with conn.cursor() as cur:
                # 获取会话信息（包含图表配置等完整字段）
                cur.execute(
                    sql.SQL("""
                    SELECT id, upload_date, upload_time, filename, row_count, col_count, columns_json, has_workload_analysis,
                           sheet_name, selected_columns, display_names, column_types, chart_types, config_name
                    FROM {}
                    WHERE id = %s
                    """).format(sql.Identifier(UPLOAD_SESSIONS_TABLE)),
                    (session_id,)
                )
                session = cur.fetchone()
                if not session:
                    raise HTTPException(status_code=404, detail="会话不存在")

                columns = parse_json_value(session[6]) or []
                # 解析图表配置字段（索引: sheet_name=8, selected_columns=9, display_names=10, column_types=11, chart_types=12, config_name=13）
                sheet_name = session[8] or ""
                selected_columns = session[9] or ""
                display_names_json = parse_json_value(session[10]) or {}
                column_types_json = parse_json_value(session[11]) or {}
                chart_types_json = parse_json_value(session[12]) or {}
                config_name = session[13] or ""

                # 获取数据行
                cur.execute(
                    sql.SQL("""
                    SELECT row_data
                    FROM {}
                    WHERE session_id = %s
                    ORDER BY row_index ASC
                    """).format(sql.Identifier(UPLOAD_DATA_TABLE)),
                    (session_id,)
                )
                raw_rows = cur.fetchall()
                data_rows = [parse_json_value(r[0]) for r in raw_rows if r[0]]

                if not data_rows:
                    raise HTTPException(status_code=404, detail="会话数据为空")

                # 构建返回 payload
                df = pd.DataFrame(data_rows)
                payload = _dataframe_to_payload(df)
                payload["all_rows"] = data_rows
                payload["all_rows_truncated"] = False
                payload["session_info"] = {
                    "session_id": session_id,
                    "upload_date": session[1].isoformat() if session[1] else None,
                    "upload_time": session[2].isoformat() if session[2] else None,
                    "filename": session[3],
                    "sheet_name": sheet_name,
                    "config_name": config_name,
                }
                # 添加图表配置（关键：前端需要这些字段渲染自定义图表）
                payload["chart_type_config"] = chart_types_json
                payload["display_name_config"] = display_names_json
                payload["column_type_config"] = column_types_json
                payload["selected_columns"] = selected_columns
                payload["sheet_name"] = sheet_name
                payload["ok"] = True
    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"读取会话数据失败：{e!s}") from e

    return JSONResponse(content=payload)


@app.post("/api/upload/delete/{session_id}")
async def delete_upload_session(session_id: int) -> JSONResponse:
    """删除指定上传会话及其数据。"""
    dsn = get_pg_dsn()
    try:
        with connect(dsn) as conn:
            with conn.cursor() as cur:
                # 先删除关联的数据行
                cur.execute(
                    sql.SQL("DELETE FROM {} WHERE session_id = %s").format(
                        sql.Identifier(UPLOAD_DATA_TABLE)
                    ),
                    (session_id,)
                )
                # 再删除会话记录
                cur.execute(
                    sql.SQL("DELETE FROM {} WHERE id = %s RETURNING id").format(
                        sql.Identifier(UPLOAD_SESSIONS_TABLE)
                    ),
                    (session_id,)
                )
                deleted = cur.fetchone()
                if not deleted:
                    raise HTTPException(status_code=404, detail="会话不存在")
                conn.commit()
    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"删除会话失败：{e!s}") from e

    return JSONResponse(content={"ok": True, "deleted_session_id": session_id})


@app.post("/api/custom-mode/save")
async def save_custom_mode(req: SaveCustomModeRequest) -> JSONResponse:
    if not req.rows:
        raise HTTPException(status_code=400, detail="rows 不能为空")
    if len(req.rows) > MAX_SAVE_ROWS:
        raise HTTPException(status_code=400, detail=f"单次最多保存 {MAX_SAVE_ROWS} 行")

    selected_columns = req.selected_columns or list(req.rows[0].keys())
    if not selected_columns:
        raise HTTPException(status_code=400, detail="未选择任何列")

    missing = [c for c in selected_columns if c not in req.rows[0]]
    if missing:
        raise HTTPException(status_code=400, detail=f"列不存在: {missing}")

    # 表名添加时间戳后缀，格式：custom_mode_{slugified_name}_{YYYYMMDD_HHMMSS}
    timestamp_suffix = datetime.now().strftime("%Y%m%d_%H%M%S")
    table_name = f"custom_mode_{slugify_mode_name(req.mode_name)}_{timestamp_suffix}"
    
    # 对列名进行安全处理，转换为安全的 SQL 列名
    def safe_column_name(col: str) -> str:
        # 移除特殊字符，保留中文、字母、数字、下划线
        safe_name = re.sub(r'[^\w\u4e00-\u9fff]', '_', col.strip())
        # 如果首字符不是字母或下划线，添加前缀
        if safe_name and not (safe_name[0].isalpha() or safe_name[0] == '_' or '\u4e00' <= safe_name[0] <= '\u9fff'):
            safe_name = 'col_' + safe_name
        return safe_name or 'col_unknown'
    
    safe_column_mapping = {col: safe_column_name(col) for col in selected_columns}
    column_types = {
        safe_column_mapping[col]: infer_sql_type([row.get(col) for row in req.rows])
        for col in selected_columns
    }

    dsn = os.getenv("PG_DSN", "").strip()
    if not dsn:
        raise HTTPException(status_code=500, detail="缺少 PG_DSN 环境变量，请先配置 PostgreSQL 连接串")

    try:
        with connect(dsn) as conn:
            with conn.cursor() as cur:
                # 自动记录模式名和写入时间，方便后续按模式检索
                create_cols: list[sql.SQL] = [
                    sql.SQL("id BIGSERIAL PRIMARY KEY"),
                    sql.SQL("mode_name TEXT NOT NULL"),
                    sql.SQL("created_at TIMESTAMPTZ NOT NULL DEFAULT NOW()"),
                ]
                for col in selected_columns:
                    safe_col = safe_column_mapping[col]
                    create_cols.append(
                        sql.SQL("{} {}").format(
                            sql.Identifier(safe_col),
                            sql.SQL(column_types[safe_col]),
                        )
                    )
                create_stmt = sql.SQL("CREATE TABLE IF NOT EXISTS {} ({})").format(
                    sql.Identifier(table_name),
                    sql.SQL(", ").join(create_cols),
                )
                cur.execute(create_stmt)

                # 插入数据时使用安全的列名
                insert_safe_cols = ["mode_name"] + [safe_column_mapping[col] for col in selected_columns]
                insert_stmt = sql.SQL("INSERT INTO {} ({}) VALUES ({})").format(
                    sql.Identifier(table_name),
                    sql.SQL(", ").join(sql.Identifier(c) for c in insert_safe_cols),
                    sql.SQL(", ").join(sql.Placeholder() for _ in insert_safe_cols),
                )
                # 使用逐行插入替代executemany，避免%符号导致的问题
                for row in req.rows:
                    row_values = [req.mode_name]
                    for col in selected_columns:
                        safe_col = safe_column_mapping[col]
                        row_values.append(normalize_cell_for_insert(row.get(col), column_types[safe_col]))
                    cur.execute(insert_stmt, tuple(row_values))
                conn.commit()
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"写入 PostgreSQL 失败：{e!s}") from e

    return JSONResponse(
        content={
            "ok": True,
            "table_name": table_name,
            "row_count": len(req.rows),
            "selected_columns": selected_columns,
            "column_types": {col: column_types[safe_column_mapping[col]] for col in selected_columns},
        }
    )


@app.get("/api/custom-mode/list")
async def list_custom_modes() -> JSONResponse:
    dsn = get_pg_dsn()
    try:
        with connect(dsn) as conn:
            with conn.cursor() as cur:
                cur.execute(
                    """
                    SELECT table_name
                    FROM information_schema.tables
                    WHERE table_schema = 'public'
                      AND table_name LIKE 'custom_mode_%'
                    ORDER BY table_name;
                    """
                )
                tables = [r[0] for r in cur.fetchall()]
                items: list[dict[str, Any]] = []
                for table in tables:
                    query = sql.SQL(
                        """
                        SELECT mode_name, COUNT(*)::BIGINT AS row_count, MAX(created_at) AS last_saved_at
                        FROM {}
                        GROUP BY mode_name
                        ORDER BY last_saved_at DESC NULLS LAST;
                        """
                    ).format(sql.Identifier(table))
                    cur.execute(query)
                    for mode_name, row_count, last_saved_at in cur.fetchall():
                        items.append(
                            {
                                "table_name": table,
                                "mode_name": mode_name,
                                "row_count": int(row_count or 0),
                                "last_saved_at": last_saved_at.isoformat() if last_saved_at else None,
                            }
                        )
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"读取 PostgreSQL 模式列表失败：{e!s}") from e

    return JSONResponse(content={"ok": True, "items": items})


@app.get("/api/custom-mode/load/{table_name}")
async def load_custom_mode_data(table_name: str) -> JSONResponse:
    """加载指定自定义模式表的数据。"""
    dsn = get_pg_dsn()
    try:
        with connect(dsn) as conn:
            with conn.cursor() as cur:
                # 获取表结构
                cur.execute(
                    """
                    SELECT column_name, data_type 
                    FROM information_schema.columns 
                    WHERE table_name = %s AND table_schema = 'public'
                    ORDER BY ordinal_position;
                    """,
                    (table_name,)
                )
                columns_info = cur.fetchall()
                columns = [col[0] for col in columns_info if col[0] not in ('id', 'mode_name', 'created_at')]
                
                # 获取数据
                cur.execute(
                    sql.SQL("SELECT {} FROM {} ORDER BY created_at DESC LIMIT 500").format(
                        sql.SQL(", ").join(sql.Identifier(c) for c in columns),
                        sql.Identifier(table_name)
                    )
                )
                rows = cur.fetchall()
                
                # 转换为字典格式
                data_rows = []
                for row in rows:
                    row_dict = {}
                    for i, col in enumerate(columns):
                        value = row[i]
                        if value is not None:
                            row_dict[col] = value
                        else:
                            row_dict[col] = ""
                    data_rows.append(row_dict)
                
                logger.info(f"加载自定义模式成功: {table_name}, {len(data_rows)} 行")
                
                # 构建预览payload
                import pandas as pd
                if data_rows:
                    df = pd.DataFrame(data_rows)
                    payload = {
                        "ok": True,
                        "table_name": table_name,
                        "columns": columns,
                        "rows": data_rows,
                        "row_count": len(data_rows),
                        "shape": {"rows": len(data_rows), "cols": len(columns)},
                        "preview_rows": data_rows[:20],
                        "dtypes": {col: str(df.dtypes.get(col, "object")) for col in columns},
                        "workload_analysis": None,  # 自定义模式不进行工作量分析
                    }
                else:
                    payload = {
                        "ok": True,
                        "table_name": table_name,
                        "columns": columns,
                        "rows": [],
                        "row_count": 0,
                    }
                    
    except Exception as e:
        logger.error(f"加载自定义模式失败: {e}")
        raise HTTPException(status_code=500, detail=f"加载自定义模式数据失败：{e!s}") from e

    return JSONResponse(content=payload)


# ========== 数据库健康检查 API ==========

@app.get("/api/db/health")
async def check_db_health() -> JSONResponse:
    """检查数据库健康状态，返回各表的统计信息。"""
    dsn = get_pg_dsn()
    try:
        with connect(dsn) as conn:
            from db.schema import check_db_health
            health_info = check_db_health(conn)
            return JSONResponse(content=health_info)
    except Exception as e:
        return JSONResponse(content={
            "healthy": False,
            "message": f"数据库连接失败: {e!s}",
            "tables": {}
        })


@app.post("/api/custom-mode/delete")
async def delete_custom_mode(req: DeleteCustomModeRequest) -> JSONResponse:
    """删除指定模式名的所有相关表（匹配 custom_mode_{slugified_name}_* 格式）。"""
    dsn = get_pg_dsn()
    slug = slugify_mode_name(req.mode_name)
    table_prefix = f"custom_mode_{slug}"
    
    try:
        with connect(dsn) as conn:
            with conn.cursor() as cur:
                # 先查找匹配的所有表
                cur.execute(
                    """
                    SELECT table_name FROM information_schema.tables
                    WHERE table_schema = 'public' AND table_name LIKE %s
                    """,
                    (table_prefix + "_%",)
                )
                matching_tables = [r[0] for r in cur.fetchall()]
                
                if not matching_tables:
                    raise HTTPException(status_code=404, detail=f"未找到模式 '{req.mode_name}' 对应的表")
                
                # 删除所有匹配的表
                deleted_tables = []
                for table_name in matching_tables:
                    cur.execute(
                        sql.SQL("DROP TABLE IF EXISTS {}").format(sql.Identifier(table_name))
                    )
                    deleted_tables.append(table_name)
                
                conn.commit()
    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"删除模式表失败：{e!s}") from e

    return JSONResponse(content={
        "ok": True, 
        "deleted_tables": deleted_tables, 
        "mode_name": req.mode_name,
        "deleted_count": len(deleted_tables)
    })


# ========== 列映射配置API ==========

@app.get("/api/column-mapping/list")
async def list_column_mappings() -> JSONResponse:
    """获取所有列映射配置。"""
    dsn = get_pg_dsn()
    try:
        with connect(dsn) as conn:
            _ensure_upload_tables_exist(conn)
            with conn.cursor() as cur:
                cur.execute(
                    sql.SQL("""
                    SELECT id, mapping_name, is_default, created_at, updated_at,
                        name_aliases, oncall_open_aliases, pending_ticket_aliases,
                        new_issue_yesterday_aliases, governance_issue_aliases, kernel_issue_aliases,
                        consult_issue_aliases, escalation_help_aliases, issue_ticket_output_aliases,
                        requirement_ticket_output_aliases, wiki_output_aliases, analysis_report_output_aliases
                    FROM {}
                    ORDER BY is_default DESC, updated_at DESC
                    """).format(sql.Identifier(COLUMN_MAPPING_TABLE))
                )
                rows = cur.fetchall()
                items = []
                for r in rows:
                    items.append({
                        "id": r[0],
                        "mapping_name": r[1],
                        "is_default": r[2],
                        "created_at": r[3].isoformat() if r[3] else None,
                        "updated_at": r[4].isoformat() if r[4] else None,
                        "aliases": {
                            "name": parse_json_value(r[5]) or [],
                            "oncall_open": parse_json_value(r[6]) or [],
                            "pending_ticket": parse_json_value(r[7]) or [],
                            "new_issue_yesterday": parse_json_value(r[8]) or [],
                            "governance_issue": parse_json_value(r[9]) or [],
                            "kernel_issue": parse_json_value(r[10]) or [],
                            "consult_issue": parse_json_value(r[11]) or [],
                            "escalation_help": parse_json_value(r[12]) or [],
                            "issue_ticket_output": parse_json_value(r[13]) or [],
                            "requirement_ticket_output": parse_json_value(r[14]) or [],
                            "wiki_output": parse_json_value(r[15]) or [],
                            "analysis_report_output": parse_json_value(r[16]) or [],
                        }
                    })
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"读取列映射配置失败：{e!s}") from e

    return JSONResponse(content={"ok": True, "items": items})


@app.get("/api/column-mapping/default")
async def get_default_column_mapping() -> JSONResponse:
    """获取默认列映射配置。"""
    return JSONResponse(content={
        "ok": True,
        "aliases": DEFAULT_COLUMN_ALIASES,
        "mapping_name": "默认配置"
    })


@app.post("/api/column-mapping/save")
async def save_column_mapping(req: SaveColumnMappingRequest) -> JSONResponse:
    """保存列映射配置。"""
    dsn = get_pg_dsn()
    try:
        with connect(dsn) as conn:
            _ensure_upload_tables_exist(conn)
            with conn.cursor() as cur:
                # 检查是否已存在同名配置
                cur.execute(
                    sql.SQL("SELECT id FROM {} WHERE mapping_name = %s").format(
                        sql.Identifier(COLUMN_MAPPING_TABLE)
                    ),
                    (req.mapping_name,)
                )
                existing = cur.fetchone()
                
                if existing:
                    # 更新现有配置
                    cur.execute(
                        sql.SQL("""
                        UPDATE {} SET
                            name_aliases = %s,
                            oncall_open_aliases = %s,
                            pending_ticket_aliases = %s,
                            new_issue_yesterday_aliases = %s,
                            governance_issue_aliases = %s,
                            kernel_issue_aliases = %s,
                            consult_issue_aliases = %s,
                            escalation_help_aliases = %s,
                            issue_ticket_output_aliases = %s,
                            requirement_ticket_output_aliases = %s,
                            wiki_output_aliases = %s,
                            analysis_report_output_aliases = %s,
                            updated_at = NOW()
                        WHERE id = %s
                        """).format(sql.Identifier(COLUMN_MAPPING_TABLE)),
                        (
                            json.dumps(req.name_aliases),
                            json.dumps(req.oncall_open_aliases),
                            json.dumps(req.pending_ticket_aliases),
                            json.dumps(req.new_issue_yesterday_aliases),
                            json.dumps(req.governance_issue_aliases),
                            json.dumps(req.kernel_issue_aliases),
                            json.dumps(req.consult_issue_aliases),
                            json.dumps(req.escalation_help_aliases),
                            json.dumps(req.issue_ticket_output_aliases),
                            json.dumps(req.requirement_ticket_output_aliases),
                            json.dumps(req.wiki_output_aliases),
                            json.dumps(req.analysis_report_output_aliases),
                            existing[0]
                        )
                    )
                    mapping_id = existing[0]
                else:
                    # 创建新配置
                    cur.execute(
                        sql.SQL("""
                        INSERT INTO {} (mapping_name, name_aliases, oncall_open_aliases, pending_ticket_aliases,
                            new_issue_yesterday_aliases, governance_issue_aliases, kernel_issue_aliases, consult_issue_aliases,
                            escalation_help_aliases, issue_ticket_output_aliases, requirement_ticket_output_aliases,
                            wiki_output_aliases, analysis_report_output_aliases)
                        VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
                        RETURNING id
                        """).format(sql.Identifier(COLUMN_MAPPING_TABLE)),
                        (
                            req.mapping_name,
                            json.dumps(req.name_aliases),
                            json.dumps(req.oncall_open_aliases),
                            json.dumps(req.pending_ticket_aliases),
                            json.dumps(req.new_issue_yesterday_aliases),
                            json.dumps(req.governance_issue_aliases),
                            json.dumps(req.kernel_issue_aliases),
                            json.dumps(req.consult_issue_aliases),
                            json.dumps(req.escalation_help_aliases),
                            json.dumps(req.issue_ticket_output_aliases),
                            json.dumps(req.requirement_ticket_output_aliases),
                            json.dumps(req.wiki_output_aliases),
                            json.dumps(req.analysis_report_output_aliases),
                        )
                    )
                    mapping_id = cur.fetchone()[0]
                conn.commit()
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"保存列映射配置失败：{e!s}") from e

    return JSONResponse(content={
        "ok": True,
        "mapping_id": mapping_id,
        "mapping_name": req.mapping_name
    })


@app.post("/api/column-mapping/delete")
async def delete_column_mapping(req: DeleteColumnMappingRequest) -> JSONResponse:
    """删除列映射配置（不能删除默认配置）。"""
    dsn = get_pg_dsn()
    try:
        with connect(dsn) as conn:
            _ensure_upload_tables_exist(conn)
            with conn.cursor() as cur:
                # 检查是否是默认配置
                cur.execute(
                    sql.SQL("SELECT is_default FROM {} WHERE id = %s").format(
                        sql.Identifier(COLUMN_MAPPING_TABLE)
                    ),
                    (req.mapping_id,)
                )
                row = cur.fetchone()
                if not row:
                    raise HTTPException(status_code=404, detail="配置不存在")
                if row[0]:  # is_default
                    raise HTTPException(status_code=400, detail="不能删除默认配置")
                
                cur.execute(
                    sql.SQL("DELETE FROM {} WHERE id = %s RETURNING id").format(
                        sql.Identifier(COLUMN_MAPPING_TABLE)
                    ),
                    (req.mapping_id,)
                )
                deleted = cur.fetchone()
                conn.commit()
    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"删除列映射配置失败：{e!s}") from e

    return JSONResponse(content={"ok": True, "deleted_id": req.mapping_id})


@app.get("/api/column-mapping/suggest")
async def suggest_column_mapping() -> JSONResponse:
    """根据Excel列名建议映射配置。"""
    return JSONResponse(content={
        "ok": True,
        "message": "请在上传Excel后调用此功能",
        "default_aliases": DEFAULT_COLUMN_ALIASES,
        "alias_labels": {
            "name": "姓名列",
            "oncall_open": "oncall未闭环",
            "pending_ticket": "待处理工单",
            "new_issue_yesterday": "昨日新增问题",
            "governance_issue": "管控问题",
            "kernel_issue": "内核问题",
            "consult_issue": "咨询问题",
            "escalation_help": "透传求助",
            "issue_ticket_output": "问题单产出",
            "requirement_ticket_output": "需求单产出",
            "wiki_output": "wiki产出",
            "analysis_report_output": "分析报告产出",
        }
    })


@app.get("/", response_class=HTMLResponse)
async def index() -> str:
    html_path = BASE_DIR / "static" / "index.html"
    return html_path.read_text(encoding="utf-8")


app.mount(
    "/static",
    StaticFiles(directory=str(BASE_DIR / "static")),
    name="static",
)
