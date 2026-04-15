"""
工具函数：数据处理、类型转换、SQL辅助
"""

import re
import json
from typing import Any

import pandas as pd


# 常量配置
MAX_PREVIEW_ROWS = 200
MAX_SAVE_ROWS = 5000
UPLOAD_SESSIONS_TABLE = "upload_sessions"
UPLOAD_DATA_TABLE = "upload_data"
COLUMN_MAPPING_TABLE = "column_mappings"


def normalize_col_name(name: str) -> str:
    """规范化列名，去除空格和括号"""
    return str(name).replace(" ", "").replace("\t", "").replace("(", "").replace(")", "").replace("（", "").replace("）", "")


def find_col(columns: list[str], aliases: list[str]) -> str | None:
    """根据别名查找列名"""
    normalized = {col: normalize_col_name(col) for col in columns}
    for alias in aliases:
        key = normalize_col_name(alias)
        for col, col_key in normalized.items():
            if col_key == key or key in col_key:
                return col
    return None


def to_float(value: Any) -> float:
    """安全转换为浮点数"""
    try:
        return float(value)
    except (TypeError, ValueError):
        return 0.0


def parse_json_value(value: Any) -> Any:
    """安全解析 JSON 值，兼容已解析和未解析的情况"""
    if value is None:
        return None
    if isinstance(value, (dict, list)):
        return value  # psycopg 已自动解析 JSONB
    if isinstance(value, (str, bytes, bytearray)):
        return json.loads(value)
    return value


def slugify_mode_name(mode_name: str) -> str:
    """将模式名转换为安全的SQL表名"""
    slug = re.sub(r"[^a-zA-Z0-9_]+", "_", mode_name.strip().lower())
    slug = slug.strip("_")
    if not slug:
        slug = "custom_mode"
    return slug[:48]


def infer_sql_type(values: list[Any]) -> str:
    """推断SQL类型"""
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


def normalize_cell_for_insert(value: Any, sql_type: str) -> Any:
    """规范化单元格数据用于插入"""
    if value is None or str(value).strip() == "":
        return None
    if sql_type == "DOUBLE PRECISION":
        try:
            return float(value)
        except (TypeError, ValueError):
            return None
    return str(value)


def dataframe_to_preview(df: pd.DataFrame, max_rows: int = MAX_PREVIEW_ROWS) -> dict[str, Any]:
    """将DataFrame转换为预览数据"""
    df = df.fillna("")
    preview = df.head(max_rows)
    rows = json.loads(preview.to_json(orient="records", force_ascii=False))
    
    dtypes = {str(c): str(t) for c, t in df.dtypes.items()}
    missing = {str(k): int(v) for k, v in df.isna().sum().astype(int).to_dict().items()}
    
    numeric = df.select_dtypes(include=["number"])
    describe = None
    if not numeric.empty:
        desc = numeric.describe().round(4)
        describe = json.loads(desc.to_json(orient="index", force_ascii=False))
    
    return {
        "shape": {"rows": int(df.shape[0]), "cols": int(df.shape[1])},
        "columns": [str(c) for c in df.columns.tolist()],
        "dtypes": dtypes,
        "missing_counts": missing,
        "preview_rows": rows,
        "preview_truncated": len(df) > max_rows,
        "numeric_describe": describe,
    }