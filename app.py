"""
Excel 上传与分析 API：读取表格、返回预览与基础统计，支持自动入库和历史记录管理。
"""

from __future__ import annotations

import io
import json
import os
import re
from datetime import datetime
from pathlib import Path
from typing import Any

import pandas as pd
from dotenv import load_dotenv
from pydantic import BaseModel, Field
from psycopg import connect
from psycopg import sql

from datetime import date as date_type

from fastapi import FastAPI, File, HTTPException, UploadFile, Query
from fastapi.responses import HTMLResponse, JSONResponse
from fastapi.staticfiles import StaticFiles

BASE_DIR = Path(__file__).resolve().parent
load_dotenv(BASE_DIR / ".env")

app = FastAPI(title="Excel 分析")

MAX_PREVIEW_ROWS = 200
UPLOAD_SESSIONS_TABLE = "upload_sessions"
UPLOAD_DATA_TABLE = "upload_data"
COLUMN_MAPPING_TABLE = "column_mappings"
DEFAULT_WEIGHTS = {
    "oncall_open": 0.9,
    "pending_ticket": 0.8,
    "new_issue_yesterday": 1.4,
    "governance_issue": 1.0,
    "kernel_issue": 1.35,
    "consult_issue": 0.85,
    "escalation_help": -0.6,
    "issue_ticket_output": 1.25,
    "requirement_ticket_output": 1.05,
    "wiki_output": 1.3,
    "analysis_report_output": 1.2,
}

# 默认列映射配置 - 每个指标对应的列名别名列表
DEFAULT_COLUMN_ALIASES = {
    "name": ["姓名", "名字", "人员", "同学", "name", "员工姓名"],
    "oncall_open": ["oncall接单未闭环的数量", "oncall未闭环", "接单未闭环", "oncall_open", "未闭环数量"],
    "pending_ticket": ["名下的待处理工单数", "待处理工单", "待处理工单数", "pending_ticket", "名下工单"],
    "new_issue_yesterday": ["昨日新增多少个问题", "昨日新增问题", "昨日新增", "new_issue_yesterday", "new_issues"],
    "governance_issue": ["多少个管控的问题", "管控问题", "管控", "governance_issue", "管控类问题"],
    "kernel_issue": ["多少个内核的问题", "内核问题", "内核", "kernel_issue", "内核类问题"],
    "consult_issue": ["多少个咨询问题", "咨询问题", "咨询", "consult_issue", "咨询类问题"],
    "escalation_help": ["透传求助了多少个", "透传求助", "透传", "escalation_help", "求助数量"],
    "issue_ticket_output": ["问题单数量", "提了多少问题单", "问题单", "issue_ticket_output", "问题单产出"],
    "requirement_ticket_output": ["需求单数量", "提了多少需求单", "需求单", "requirement_ticket_output", "需求单产出"],
    "wiki_output": ["wiki输出数量", "输出多少wiki", "wiki", "wiki_output", "wiki产出"],
    "analysis_report_output": ["问题分析报告数量", "输出多少问题分析报告", "分析报告", "analysis_report_output", "报告产出"],
}

MAX_SAVE_ROWS = 5000


def _normalize_col_name(name: str) -> str:
    return str(name).replace(" ", "").replace("\t", "").replace("(", "").replace(")", "").replace("（", "").replace("）", "")


def _find_col(columns: list[str], aliases: list[str]) -> str | None:
    normalized = {col: _normalize_col_name(col) for col in columns}
    for alias in aliases:
        key = _normalize_col_name(alias)
        for col, col_key in normalized.items():
            if col_key == key or key in col_key:
                return col
    return None


def _to_float(value: Any) -> float:
    try:
        return float(value)
    except (TypeError, ValueError):
        return 0.0


class SaveCustomModeRequest(BaseModel):
    mode_name: str = Field(min_length=1, max_length=64)
    selected_columns: list[str] = Field(default_factory=list)
    rows: list[dict[str, Any]] = Field(default_factory=list)


class DeleteCustomModeRequest(BaseModel):
    mode_name: str = Field(min_length=1, max_length=64)


def _parse_json_value(value: Any) -> Any:
    """安全解析 JSON 值，兼容已解析和未解析的情况。"""
    if value is None:
        return None
    if isinstance(value, (dict, list)):
        # psycopg 已自动解析 JSONB
        return value
    if isinstance(value, (str, bytes, bytearray)):
        # 需要手动解析
        return json.loads(value)
    return value
    slug = re.sub(r"[^a-zA-Z0-9_]+", "_", mode_name.strip().lower())
    slug = slug.strip("_")
    if not slug:
        slug = "custom_mode"
    return slug[:48]


def _infer_sql_type(values: list[Any]) -> str:
    non_empty = [v for v in values if str(v).strip() != ""]
    if not non_empty:
        return "TEXT"
    numeric_count = 0
    for v in non_empty:
        try:
            float(v)
            numeric_count += 1
        except (TypeError, ValueError):
            pass
    if numeric_count == len(non_empty):
        return "DOUBLE PRECISION"
    return "TEXT"


def _normalize_cell_for_insert(value: Any, sql_type: str) -> Any:
    if value is None or str(value).strip() == "":
        return None
    if sql_type == "DOUBLE PRECISION":
        try:
            return float(value)
        except (TypeError, ValueError):
            return None
    return str(value)


def _get_pg_dsn() -> str:
    dsn = os.getenv("PG_DSN", "").strip()
    if not dsn:
        raise HTTPException(status_code=500, detail="缺少 PG_DSN 环境变量，请先配置 PostgreSQL 连接串")
    return dsn


def _ensure_upload_tables_exist(conn) -> None:
    """确保上传历史表、数据表和列映射配置表存在。"""
    with conn.cursor() as cur:
        # 上传会话表：记录每次上传的元信息
        cur.execute(
            sql.SQL("""
            CREATE TABLE IF NOT EXISTS {} (
                id BIGSERIAL PRIMARY KEY,
                upload_date DATE NOT NULL,
                upload_time TIMESTAMPTZ NOT NULL DEFAULT NOW(),
                filename TEXT,
                row_count BIGINT NOT NULL,
                col_count BIGINT NOT NULL,
                columns_json JSONB,
                has_workload_analysis BOOLEAN DEFAULT FALSE,
                notes TEXT,
                column_mapping_id BIGINT
            )
            """).format(sql.Identifier(UPLOAD_SESSIONS_TABLE))
        )
        # 上传数据表：存储每次上传的具体数据行
        cur.execute(
            sql.SQL("""
            CREATE TABLE IF NOT EXISTS {} (
                id BIGSERIAL PRIMARY KEY,
                session_id BIGINT NOT NULL REFERENCES {}(id) ON DELETE CASCADE,
                row_index BIGINT NOT NULL,
                row_data JSONB NOT NULL
            )
            """).format(
                sql.Identifier(UPLOAD_DATA_TABLE),
                sql.Identifier(UPLOAD_SESSIONS_TABLE)
            )
        )
        # 列映射配置表：存储自定义列名映射
        cur.execute(
            sql.SQL("""
            CREATE TABLE IF NOT EXISTS {} (
                id BIGSERIAL PRIMARY KEY,
                mapping_name TEXT NOT NULL UNIQUE,
                name_aliases JSONB NOT NULL DEFAULT '[]',
                oncall_open_aliases JSONB NOT NULL DEFAULT '[]',
                pending_ticket_aliases JSONB NOT NULL DEFAULT '[]',
                new_issue_yesterday_aliases JSONB NOT NULL DEFAULT '[]',
                governance_issue_aliases JSONB NOT NULL DEFAULT '[]',
                kernel_issue_aliases JSONB NOT NULL DEFAULT '[]',
                consult_issue_aliases JSONB NOT NULL DEFAULT '[]',
                escalation_help_aliases JSONB NOT NULL DEFAULT '[]',
                issue_ticket_output_aliases JSONB NOT NULL DEFAULT '[]',
                requirement_ticket_output_aliases JSONB NOT NULL DEFAULT '[]',
                wiki_output_aliases JSONB NOT NULL DEFAULT '[]',
                analysis_report_output_aliases JSONB NOT NULL DEFAULT '[]',
                is_default BOOLEAN DEFAULT FALSE,
                created_at TIMESTAMPTZ NOT NULL DEFAULT NOW(),
                updated_at TIMESTAMPTZ NOT NULL DEFAULT NOW()
            )
            """).format(sql.Identifier(COLUMN_MAPPING_TABLE))
        )
        # 创建索引加速查询
        cur.execute(
            sql.SQL("CREATE INDEX IF NOT EXISTS idx_upload_sessions_date ON {} (upload_date DESC)").format(
                sql.Identifier(UPLOAD_SESSIONS_TABLE)
            )
        )
        cur.execute(
            sql.SQL("CREATE INDEX IF NOT EXISTS idx_upload_data_session ON {} (session_id)").format(
                sql.Identifier(UPLOAD_DATA_TABLE)
            )
        )
        cur.execute(
            sql.SQL("CREATE INDEX IF NOT EXISTS idx_column_mappings_name ON {} (mapping_name)").format(
                sql.Identifier(COLUMN_MAPPING_TABLE)
            )
        )
        # 插入默认列映射配置（如果不存在）
        cur.execute(
            sql.SQL("""
            INSERT INTO {} (mapping_name, name_aliases, oncall_open_aliases, pending_ticket_aliases,
                new_issue_yesterday_aliases, governance_issue_aliases, kernel_issue_aliases, consult_issue_aliases,
                escalation_help_aliases, issue_ticket_output_aliases, requirement_ticket_output_aliases,
                wiki_output_aliases, analysis_report_output_aliases, is_default)
            VALUES ('默认配置', %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, TRUE)
            ON CONFLICT (mapping_name) DO NOTHING
            """).format(sql.Identifier(COLUMN_MAPPING_TABLE)),
            (
                json.dumps(DEFAULT_COLUMN_ALIASES["name"]),
                json.dumps(DEFAULT_COLUMN_ALIASES["oncall_open"]),
                json.dumps(DEFAULT_COLUMN_ALIASES["pending_ticket"]),
                json.dumps(DEFAULT_COLUMN_ALIASES["new_issue_yesterday"]),
                json.dumps(DEFAULT_COLUMN_ALIASES["governance_issue"]),
                json.dumps(DEFAULT_COLUMN_ALIASES["kernel_issue"]),
                json.dumps(DEFAULT_COLUMN_ALIASES["consult_issue"]),
                json.dumps(DEFAULT_COLUMN_ALIASES["escalation_help"]),
                json.dumps(DEFAULT_COLUMN_ALIASES["issue_ticket_output"]),
                json.dumps(DEFAULT_COLUMN_ALIASES["requirement_ticket_output"]),
                json.dumps(DEFAULT_COLUMN_ALIASES["wiki_output"]),
                json.dumps(DEFAULT_COLUMN_ALIASES["analysis_report_output"]),
            )
        )
        conn.commit()


def _save_upload_to_db(
    conn,
    filename: str,
    rows: list[dict[str, Any]],
    columns: list[str],
    has_workload_analysis: bool,
    column_mapping_id: int | None = None,
) -> int:
    """保存上传数据到数据库，返回 session_id。"""
    today = datetime.now().date()
    row_count = len(rows)
    col_count = len(columns)

    with conn.cursor() as cur:
        # 插入会话记录
        cur.execute(
            sql.SQL("""
            INSERT INTO {} (upload_date, filename, row_count, col_count, columns_json, has_workload_analysis, column_mapping_id)
            VALUES (%s, %s, %s, %s, %s, %s, %s)
            RETURNING id
            """).format(sql.Identifier(UPLOAD_SESSIONS_TABLE)),
            (today, filename, row_count, col_count, json.dumps(columns), has_workload_analysis, column_mapping_id)
        )
        session_id = cur.fetchone()[0]

        # 批量插入数据行
        insert_stmt = sql.SQL("INSERT INTO {} (session_id, row_index, row_data) VALUES (%s, %s, %s)").format(
            sql.Identifier(UPLOAD_DATA_TABLE)
        )
        batch_values = [(session_id, idx, json.dumps(row)) for idx, row in enumerate(rows)]
        cur.executemany(insert_stmt, batch_values)
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


def _get_column_aliases_from_config(mapping_config: dict[str, Any] | None) -> dict[str, list[str]]:
    """从配置中提取列别名映射，如果无配置则使用默认值。"""
    if mapping_config is None:
        return DEFAULT_COLUMN_ALIASES.copy()
    
    aliases = {}
    for key in DEFAULT_COLUMN_ALIASES.keys():
        config_key = f"{key}_aliases"
        if config_key in mapping_config:
            aliases[key] = _parse_json_value(mapping_config[config_key]) or []
        else:
            aliases[key] = DEFAULT_COLUMN_ALIASES[key]
    return aliases


def _build_workload_analysis(
    df: pd.DataFrame, 
    column_aliases: dict[str, list[str]] | None = None
) -> dict[str, Any] | None:
    """构建工作量分析，支持自定义列别名。"""
    cols = [str(c) for c in df.columns.tolist()]
    
    # 使用传入的别名或默认别名
    if column_aliases is None:
        column_aliases = DEFAULT_COLUMN_ALIASES.copy()
    
    name_col = _find_col(cols, column_aliases.get("name", ["姓名"]))
    if not name_col:
        return None

    col_map = {
        "oncall_open": _find_col(cols, column_aliases.get("oncall_open", ["oncall接单未闭环的数量"])),
        "pending_ticket": _find_col(cols, column_aliases.get("pending_ticket", ["名下的待处理工单数"])),
        "new_issue_yesterday": _find_col(cols, column_aliases.get("new_issue_yesterday", ["昨日新增多少个问题"])),
        "governance_issue": _find_col(cols, column_aliases.get("governance_issue", ["多少个管控的问题"])),
        "kernel_issue": _find_col(cols, column_aliases.get("kernel_issue", ["多少个内核的问题"])),
        "consult_issue": _find_col(cols, column_aliases.get("consult_issue", ["多少个咨询问题"])),
        "escalation_help": _find_col(cols, column_aliases.get("escalation_help", ["透传求助了多少个"])),
        "issue_ticket_output": _find_col(cols, column_aliases.get("issue_ticket_output", ["问题单数量"])),
        "requirement_ticket_output": _find_col(cols, column_aliases.get("requirement_ticket_output", ["需求单数量"])),
        "wiki_output": _find_col(cols, column_aliases.get("wiki_output", ["wiki输出数量"])),
        "analysis_report_output": _find_col(cols, column_aliases.get("analysis_report_output", ["问题分析报告数量"])),
    }

    weights = DEFAULT_WEIGHTS.copy()

    people: list[dict[str, Any]] = []
    for row in df.fillna("").to_dict(orient="records"):
        person = str(row.get(name_col, "")).strip()
        if not person:
            continue
        metrics = {k: _to_float(row.get(v, 0)) if v else 0.0 for k, v in col_map.items()}
        daily_issue_total = metrics["governance_issue"] + metrics["kernel_issue"] + metrics["consult_issue"]
        score = sum(metrics[k] * w for k, w in weights.items())
        item = {
            "name": person,
            **{k: round(v, 2) for k, v in metrics.items()},
            "daily_issue_total": round(daily_issue_total, 2),
            "workload_score": round(score, 2),
        }
        item.update(_build_person_risk(item))
        people.append(item)

    if not people:
        return None

    by_score = sorted(people, key=lambda x: x["workload_score"], reverse=True)
    by_escalation = sorted(people, key=lambda x: x["escalation_help"], reverse=True)
    by_risk = sorted(people, key=lambda x: x["risk_score"], reverse=True)
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


@app.post("/api/upload")
async def upload_excel(
    file: UploadFile = File(...),
    column_mapping_id: int | None = Query(None, description="列映射配置ID，可选")
) -> JSONResponse:
    name = (file.filename or "").lower()
    if not name.endswith((".xlsx", ".xls")):
        raise HTTPException(
            status_code=400,
            detail="请上传 .xlsx 或 .xls 文件",
        )
    raw = await file.read()
    if not raw:
        raise HTTPException(status_code=400, detail="文件为空")

    try:
        bio = io.BytesIO(raw)
        engine = "openpyxl" if name.endswith(".xlsx") else "xlrd"
        df = pd.read_excel(bio, engine=engine)
    except Exception as e:
        raise HTTPException(
            status_code=422,
            detail=f"无法解析 Excel：{e!s}",
        ) from e

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
        except Exception:
            pass  # 配置加载失败时使用默认配置
    
    if column_aliases is None:
        column_aliases = DEFAULT_COLUMN_ALIASES.copy()
        mapping_name = "默认配置"

    payload = _dataframe_to_payload(df, column_aliases)
    all_rows = json.loads(df.fillna("").to_json(orient="records", force_ascii=False))
    if len(all_rows) > MAX_SAVE_ROWS:
        all_rows = all_rows[:MAX_SAVE_ROWS]
    payload["all_rows"] = all_rows
    payload["all_rows_truncated"] = len(df) > MAX_SAVE_ROWS
    payload["column_mapping_id"] = column_mapping_id
    payload["column_mapping_name"] = mapping_name

    # 自动保存到 PostgreSQL（如果配置了 PG_DSN）
    if dsn:
        columns = payload["columns"]
        has_workload = payload["workload_analysis"] is not None
        try:
            with connect(dsn) as conn:
                _ensure_upload_tables_exist(conn)
                session_id = _save_upload_to_db(
                    conn, file.filename or "unknown.xlsx", all_rows, columns, has_workload,
                    column_mapping_id
                )
                payload["saved_to_db"] = True
                payload["session_id"] = session_id
        except Exception as e:
            # 入库失败不影响返回数据，只记录警告
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
    dsn = _get_pg_dsn()
    try:
        with connect(dsn) as conn:
            with conn.cursor() as cur:
                # 构建带日期筛选的查询
                base_query = sql.SQL("""
                    SELECT 
                        id, upload_date, upload_time, filename, row_count, col_count,
                        has_workload_analysis, notes
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
                        "notes": r[7],
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


@app.get("/api/upload/latest")
async def get_latest_upload() -> JSONResponse:
    """获取最新一次上传的数据。"""
    dsn = _get_pg_dsn()
    try:
        with connect(dsn) as conn:
            with conn.cursor() as cur:
                # 获取最新会话
                cur.execute(
                    sql.SQL("""
                    SELECT id, upload_date, upload_time, filename, row_count, col_count, columns_json, has_workload_analysis
                    FROM {}
                    ORDER BY upload_date DESC, upload_time DESC
                    LIMIT 1
                    """).format(sql.Identifier(UPLOAD_SESSIONS_TABLE))
                )
                session = cur.fetchone()
                if not session:
                    return JSONResponse(content={"ok": False, "message": "暂无历史数据"})

                session_id = session[0]
                columns = _parse_json_value(session[6]) or []

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
                data_rows = [_parse_json_value(r[0]) for r in raw_rows if r[0]]

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
                }
                payload["ok"] = True
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"读取最新数据失败：{e!s}") from e

    return JSONResponse(content=payload)


@app.get("/api/upload/session/{session_id}")
async def get_upload_session(session_id: int) -> JSONResponse:
    """获取指定会话的数据。"""
    dsn = _get_pg_dsn()
    try:
        with connect(dsn) as conn:
            with conn.cursor() as cur:
                # 获取会话信息
                cur.execute(
                    sql.SQL("""
                    SELECT id, upload_date, upload_time, filename, row_count, col_count, columns_json, has_workload_analysis
                    FROM {}
                    WHERE id = %s
                    """).format(sql.Identifier(UPLOAD_SESSIONS_TABLE)),
                    (session_id,)
                )
                session = cur.fetchone()
                if not session:
                    raise HTTPException(status_code=404, detail="会话不存在")

                columns = _parse_json_value(session[6]) or []

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
                data_rows = [_parse_json_value(r[0]) for r in raw_rows if r[0]]

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
                }
                payload["ok"] = True
    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"读取会话数据失败：{e!s}") from e

    return JSONResponse(content=payload)


@app.post("/api/upload/delete/{session_id}")
async def delete_upload_session(session_id: int) -> JSONResponse:
    """删除指定上传会话及其数据。"""
    dsn = _get_pg_dsn()
    try:
        with connect(dsn) as conn:
            with conn.cursor() as cur:
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

    table_name = f"custom_mode_{_slugify_mode_name(req.mode_name)}"
    column_types = {
        col: _infer_sql_type([row.get(col) for row in req.rows])
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
                    create_cols.append(
                        sql.SQL("{} {}").format(
                            sql.Identifier(col),
                            sql.SQL(column_types[col]),
                        )
                    )
                create_stmt = sql.SQL("CREATE TABLE IF NOT EXISTS {} ({})").format(
                    sql.Identifier(table_name),
                    sql.SQL(", ").join(create_cols),
                )
                cur.execute(create_stmt)

                insert_cols = ["mode_name", *selected_columns]
                insert_stmt = sql.SQL("INSERT INTO {} ({}) VALUES ({})").format(
                    sql.Identifier(table_name),
                    sql.SQL(", ").join(sql.Identifier(c) for c in insert_cols),
                    sql.SQL(", ").join(sql.Placeholder() for _ in insert_cols),
                )
                values_batch = []
                for row in req.rows:
                    row_values = [req.mode_name]
                    for col in selected_columns:
                        row_values.append(_normalize_cell_for_insert(row.get(col), column_types[col]))
                    values_batch.append(tuple(row_values))
                cur.executemany(insert_stmt, values_batch)
                conn.commit()
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"写入 PostgreSQL 失败：{e!s}") from e

    return JSONResponse(
        content={
            "ok": True,
            "table_name": table_name,
            "row_count": len(req.rows),
            "selected_columns": selected_columns,
            "column_types": column_types,
        }
    )


@app.get("/api/custom-mode/list")
async def list_custom_modes() -> JSONResponse:
    dsn = _get_pg_dsn()
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


@app.post("/api/custom-mode/delete")
async def delete_custom_mode(req: DeleteCustomModeRequest) -> JSONResponse:
    dsn = _get_pg_dsn()
    table_name = f"custom_mode_{_slugify_mode_name(req.mode_name)}"
    try:
        with connect(dsn) as conn:
            with conn.cursor() as cur:
                cur.execute(
                    sql.SQL("DROP TABLE IF EXISTS {}").format(sql.Identifier(table_name))
                )
                conn.commit()
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"删除模式表失败：{e!s}") from e

    return JSONResponse(content={"ok": True, "table_name": table_name, "mode_name": req.mode_name})


# ========== 列映射配置API ==========

class SaveColumnMappingRequest(BaseModel):
    mapping_name: str = Field(min_length=1, max_length=64)
    name_aliases: list[str] = Field(default_factory=list)
    oncall_open_aliases: list[str] = Field(default_factory=list)
    pending_ticket_aliases: list[str] = Field(default_factory=list)
    new_issue_yesterday_aliases: list[str] = Field(default_factory=list)
    governance_issue_aliases: list[str] = Field(default_factory=list)
    kernel_issue_aliases: list[str] = Field(default_factory=list)
    consult_issue_aliases: list[str] = Field(default_factory=list)
    escalation_help_aliases: list[str] = Field(default_factory=list)
    issue_ticket_output_aliases: list[str] = Field(default_factory=list)
    requirement_ticket_output_aliases: list[str] = Field(default_factory=list)
    wiki_output_aliases: list[str] = Field(default_factory=list)
    analysis_report_output_aliases: list[str] = Field(default_factory=list)


class DeleteColumnMappingRequest(BaseModel):
    mapping_id: int


@app.get("/api/column-mapping/list")
async def list_column_mappings() -> JSONResponse:
    """获取所有列映射配置。"""
    dsn = _get_pg_dsn()
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
                            "name": _parse_json_value(r[5]) or [],
                            "oncall_open": _parse_json_value(r[6]) or [],
                            "pending_ticket": _parse_json_value(r[7]) or [],
                            "new_issue_yesterday": _parse_json_value(r[8]) or [],
                            "governance_issue": _parse_json_value(r[9]) or [],
                            "kernel_issue": _parse_json_value(r[10]) or [],
                            "consult_issue": _parse_json_value(r[11]) or [],
                            "escalation_help": _parse_json_value(r[12]) or [],
                            "issue_ticket_output": _parse_json_value(r[13]) or [],
                            "requirement_ticket_output": _parse_json_value(r[14]) or [],
                            "wiki_output": _parse_json_value(r[15]) or [],
                            "analysis_report_output": _parse_json_value(r[16]) or [],
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
    dsn = _get_pg_dsn()
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
    dsn = _get_pg_dsn()
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
