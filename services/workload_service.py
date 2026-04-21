"""
工作量分析服务
"""

from typing import Any

import pandas as pd

from models import DEFAULT_WEIGHTS, DEFAULT_COLUMN_ALIASES
from utils import find_col, to_float


def build_person_risk(person: dict[str, Any]) -> dict[str, Any]:
    """构建人员风险评估。"""
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


def build_workload_analysis(
    df: pd.DataFrame, 
    column_aliases: dict[str, list[str]] | None = None
) -> dict[str, Any] | None:
    """构建工作量分析，支持自定义列别名。"""
    cols = [str(c) for c in df.columns.tolist()]
    
    if column_aliases is None:
        column_aliases = DEFAULT_COLUMN_ALIASES.copy()
    
    name_col = find_col(cols, column_aliases.get("name", ["姓名"]))
    if not name_col:
        return None

    col_map = {
        "oncall_open": find_col(cols, column_aliases.get("oncall_open", ["oncall接单未闭环的数量"])),
        "pending_ticket": find_col(cols, column_aliases.get("pending_ticket", ["名下的待处理工单数"])),
        "new_issue_yesterday": find_col(cols, column_aliases.get("new_issue_yesterday", ["昨日新增多少个问题"])),
        "governance_issue": find_col(cols, column_aliases.get("governance_issue", ["多少个管控的问题"])),
        "kernel_issue": find_col(cols, column_aliases.get("kernel_issue", ["多少个内核的问题"])),
        "consult_issue": find_col(cols, column_aliases.get("consult_issue", ["多少个咨询问题"])),
        "escalation_help": find_col(cols, column_aliases.get("escalation_help", ["透传求助了多少个"])),
        "issue_ticket_output": find_col(cols, column_aliases.get("issue_ticket_output", ["问题单数量"])),
        "requirement_ticket_output": find_col(cols, column_aliases.get("requirement_ticket_output", ["需求单数量"])),
        "wiki_output": find_col(cols, column_aliases.get("wiki_output", ["wiki输出数量"])),
        "analysis_report_output": find_col(cols, column_aliases.get("analysis_report_output", ["问题分析报告数量"])),
    }

    weights = DEFAULT_WEIGHTS.copy()

    people: list[dict[str, Any]] = []
    for row in df.fillna("").to_dict(orient="records"):
        person = str(row.get(name_col, "")).strip()
        if not person:
            continue
        metrics = {k: to_float(row.get(v, 0)) if v else 0.0 for k, v in col_map.items()}
        daily_issue_total = metrics["governance_issue"] + metrics["kernel_issue"] + metrics["consult_issue"]
        score = sum(metrics[k] * w for k, w in weights.items())
        item = {
            "name": person,
            **{k: round(v, 2) for k, v in metrics.items()},
            "daily_issue_total": round(daily_issue_total, 2),
            "workload_score": round(score, 2),
        }
        item.update(build_person_risk(item))
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


def dataframe_to_payload(
    df: pd.DataFrame, 
    column_aliases: dict[str, list[str]] | None = None,
    max_preview_rows: int = 200
) -> dict[str, Any]:
    """将DataFrame转换为API响应payload。"""
    import json
    from utils import MAX_PREVIEW_ROWS
    
    df = df.fillna("")
    preview = df.head(max_preview_rows)
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
        "preview_truncated": len(df) > max_preview_rows,
        "numeric_describe": describe,
        "workload_analysis": build_workload_analysis(df, column_aliases),
        "column_aliases": column_aliases or DEFAULT_COLUMN_ALIASES,
    }