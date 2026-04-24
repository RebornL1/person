"""
工具函数：数据处理、类型转换、SQL辅助、日期识别
"""

import re
import json
from datetime import datetime, date
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


def safe_cell_value(value: Any) -> Any:
    """
    安全处理单元格值，兼容：
    1. 空值（None, NaN, 空字符串） -> 返回空字符串或0
    2. 公式错误（pandas读取公式时可能得到错误标记） -> 返回错误提示
    3. 正常值 -> 返回原始值
    """
    # 处理 None
    if value is None:
        return ""
    
    # 处理 pandas NaN
    try:
        if pd.isna(value):
            return ""
    except (TypeError, ValueError):
        pass
    
    # 处理字符串类型
    if isinstance(value, str):
        s = value.strip()
        if s == "":
            return ""
        # 检查是否是公式错误标记
        if s.startswith("#") and any(err in s.upper() for err in ["N/A", "REF", "VALUE", "DIV", "NUM", "NAME", "NULL"]):
            return s  # 保持公式错误标记可见
        return s
    
    # 处理数值类型
    if isinstance(value, (int, float)):
        # 处理特殊浮点值
        if isinstance(value, float):
            try:
                if value == float('inf') or value == float('-inf'):
                    return "∞"
                if pd.isna(value):  # 再次检查 NaN
                    return ""
            except (TypeError, ValueError):
                pass
        return value
    
    # 处理其他类型（如 datetime）
    return str(value)


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
    """将DataFrame转换为预览数据，兼容空单元格和公式错误"""
    # 替换NaN为空字符串
    df = df.fillna("")
    
    # 处理公式错误标记（Excel读取时可能保留的错误值）
    for col in df.columns:
        df[col] = df[col].apply(safe_cell_value)
    
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


def parse_sheet_name_to_date(sheet_name: str, reference_year: int = None) -> dict[str, Any]:
    """
    将sheet名称解析为日期，支持多种格式：
    
    支持的格式：
    1. "xx月xx日" 或 "x月x日" -> 如 "4月21日", "12月1日"
    2. "x-xx" 或 "xx-xx" -> 如 "4-21", "04-21"
    3. "xxxx" 或 "xxxxxx" -> 如 "0423", "20260423"
    4. "xxxx年xx月xx日" -> 如 "2026年4月21日"
    5. "xxxx.xx.xx" -> 如 "2026.04.21"
    6. "Month Day" -> 如 "April 21" (英文格式)
    
    Args:
        sheet_name: sheet页名称
        reference_year: 参考年份（用于无年份的格式），默认当前年份
    
    Returns:
        {
            "original_name": 原始sheet名称,
            "parsed_date": 解析后的日期字符串 (YYYY-MM-DD) 或 None,
            "date_obj": datetime对象 或 None,
            "is_date": 是否成功识别为日期,
            "format_type": 识别的格式类型,
            "display_name": 用于显示的日期名称（如 "4月21日"）
        }
    """
    if not sheet_name:
        return {"original_name": sheet_name, "parsed_date": None, "is_date": False, "format_type": None}
    
    sheet_name = str(sheet_name).strip()
    if reference_year is None:
        reference_year = datetime.now().year
    
    result = {
        "original_name": sheet_name,
        "parsed_date": None,
        "date_obj": None,
        "is_date": False,
        "format_type": None,
        "display_name": sheet_name,
    }
    
    # 格式1: xx月xx日 或 x月x日
    pattern1 = re.match(r'^(\d{1,2})月(\d{1,2})日$', sheet_name)
    if pattern1:
        month = int(pattern1.group(1))
        day = int(pattern1.group(2))
        try:
            date_obj = date(reference_year, month, day)
            result["parsed_date"] = date_obj.isoformat()
            result["date_obj"] = date_obj
            result["is_date"] = True
            result["format_type"] = "xx月xx日"
            result["display_name"] = f"{month}月{day}日"
            return result
        except ValueError:
            pass
    
    # 格式4: xxxx年xx月xx日
    pattern4 = re.match(r'^(\d{4})年(\d{1,2})月(\d{1,2})日$', sheet_name)
    if pattern4:
        year = int(pattern4.group(1))
        month = int(pattern4.group(2))
        day = int(pattern4.group(3))
        try:
            date_obj = date(year, month, day)
            result["parsed_date"] = date_obj.isoformat()
            result["date_obj"] = date_obj
            result["is_date"] = True
            result["format_type"] = "xxxx年xx月xx日"
            result["display_name"] = f"{month}月{day}日"
            return result
        except ValueError:
            pass
    
    # 格式2: x-xx 或 xx-xx 或 xxxx-xx-xx
    pattern2 = re.match(r'^(\d{4})-(\d{1,2})-(\d{1,2})$', sheet_name)  # YYYY-MM-DD
    if pattern2:
        year = int(pattern2.group(1))
        month = int(pattern2.group(2))
        day = int(pattern2.group(3))
        try:
            date_obj = date(year, month, day)
            result["parsed_date"] = date_obj.isoformat()
            result["date_obj"] = date_obj
            result["is_date"] = True
            result["format_type"] = "YYYY-MM-DD"
            result["display_name"] = f"{month}月{day}日"
            return result
        except ValueError:
            pass
    
    pattern2b = re.match(r'^(\d{1,2})-(\d{1,2})$', sheet_name)  # MM-DD
    if pattern2b:
        month = int(pattern2b.group(1))
        day = int(pattern2b.group(2))
        try:
            date_obj = date(reference_year, month, day)
            result["parsed_date"] = date_obj.isoformat()
            result["date_obj"] = date_obj
            result["is_date"] = True
            result["format_type"] = "MM-DD"
            result["display_name"] = f"{month}月{day}日"
            return result
        except ValueError:
            pass
    
    # 格式3: xxxx (如0423表示4月23日) 或 xxxxxx (如260423表示26年4月23日)
    pattern3a = re.match(r'^(\d{6})$', sheet_name)  # YYMMDD
    if pattern3a:
        digits = sheet_name
        year_part = int(digits[:2])
        month = int(digits[2:4])
        day = int(digits[4:6])
        # 26表示2026年
        year = 2000 + year_part if year_part >= 26 else 1900 + year_part
        try:
            date_obj = date(year, month, day)
            result["parsed_date"] = date_obj.isoformat()
            result["date_obj"] = date_obj
            result["is_date"] = True
            result["format_type"] = "YYMMDD"
            result["display_name"] = f"{month}月{day}日"
            return result
        except ValueError:
            pass
    
    pattern3b = re.match(r'^(\d{8})$', sheet_name)  # YYYYMMDD
    if pattern3b:
        digits = sheet_name
        year = int(digits[:4])
        month = int(digits[4:6])
        day = int(digits[6:8])
        try:
            date_obj = date(year, month, day)
            result["parsed_date"] = date_obj.isoformat()
            result["date_obj"] = date_obj
            result["is_date"] = True
            result["format_type"] = "YYYYMMDD"
            result["display_name"] = f"{month}月{day}日"
            return result
        except ValueError:
            pass
    
    pattern3c = re.match(r'^(\d{4})$', sheet_name)  # MMDD (如0423)
    if pattern3c:
        digits = sheet_name
        month = int(digits[:2])
        day = int(digits[2:4])
        try:
            date_obj = date(reference_year, month, day)
            result["parsed_date"] = date_obj.isoformat()
            result["date_obj"] = date_obj
            result["is_date"] = True
            result["format_type"] = "MMDD"
            result["display_name"] = f"{month}月{day}日"
            return result
        except ValueError:
            pass
    
    # 格式5: xxxx.xx.xx 或 xx.xx
    pattern5a = re.match(r'^(\d{4})\.(\d{1,2})\.(\d{1,2})$', sheet_name)  # YYYY.MM.DD
    if pattern5a:
        year = int(pattern5a.group(1))
        month = int(pattern5a.group(2))
        day = int(pattern5a.group(3))
        try:
            date_obj = date(year, month, day)
            result["parsed_date"] = date_obj.isoformat()
            result["date_obj"] = date_obj
            result["is_date"] = True
            result["format_type"] = "YYYY.MM.DD"
            result["display_name"] = f"{month}月{day}日"
            return result
        except ValueError:
            pass
    
    pattern5b = re.match(r'^(\d{1,2})\.(\d{1,2})$', sheet_name)  # MM.DD
    if pattern5b:
        month = int(pattern5b.group(1))
        day = int(pattern5b.group(2))
        try:
            date_obj = date(reference_year, month, day)
            result["parsed_date"] = date_obj.isoformat()
            result["date_obj"] = date_obj
            result["is_date"] = True
            result["format_type"] = "MM.DD"
            result["display_name"] = f"{month}月{day}日"
            return result
        except ValueError:
            pass
    
    # 格式6: 英文月份名（如 April 21, Apr 21）
    month_names = {
        'january': 1, 'jan': 1, 'february': 2, 'feb': 2, 'march': 3, 'mar': 3,
        'april': 4, 'apr': 4, 'may': 5, 'june': 6, 'jun': 6,
        'july': 7, 'jul': 7, 'august': 8, 'aug': 8, 'september': 9, 'sep': 9,
        'october': 10, 'oct': 10, 'november': 11, 'nov': 11, 'december': 12, 'dec': 12
    }
    pattern6 = re.match(r'^([A-Za-z]+)\s+(\d{1,2})$', sheet_name)
    if pattern6:
        month_str = pattern6.group(1).lower()
        day = int(pattern6.group(2))
        if month_str in month_names:
            month = month_names[month_str]
            try:
                date_obj = date(reference_year, month, day)
                result["parsed_date"] = date_obj.isoformat()
                result["date_obj"] = date_obj
                result["is_date"] = True
                result["format_type"] = "Month Day"
                result["display_name"] = f"{month}月{day}日"
                return result
            except ValueError:
                pass
    
    # 无法识别为日期
    return result


def parse_all_sheet_dates(sheet_names: list[str], reference_year: int = None) -> list[dict[str, Any]]:
    """
    解析所有sheet名称为日期，并返回结果列表
    
    Args:
        sheet_names: sheet名称列表
        reference_year: 参考年份
    
    Returns:
        [
            {
                "original_name": 原始名称,
                "parsed_date": 解析后的日期字符串,
                "is_date": 是否识别为日期,
                ...
            },
            ...
        ]
    """
    results = []
    for name in sheet_names:
        parsed = parse_sheet_name_to_date(name, reference_year)
        results.append(parsed)
    return results


def get_date_display_order(date_info_list: list[dict[str, Any]]) -> list[dict[str, Any]]:
    """
    按日期排序sheet信息（仅对成功识别为日期的sheet排序）
    
    Args:
        date_info_list: parse_all_sheet_dates返回的列表
    
    Returns:
        按日期排序后的列表（未识别的排在最后）
    """
    dated = [d for d in date_info_list if d.get("is_date") and d.get("date_obj")]
    undated = [d for d in date_info_list if not d.get("is_date")]
    
    dated.sort(key=lambda x: x.get("date_obj"))
    
    return dated + undated