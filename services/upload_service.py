"""
上传数据处理服务
"""

import json
import re
from datetime import datetime
from typing import Any

from psycopg import sql, extras

from config.settings import (
    logger,
    UPLOAD_SESSIONS_TABLE,
    UPLOAD_DATA_TABLE,
    MAX_SAVE_ROWS,
)
from db.schema import ensure_upload_tables_exist
from models import DEFAULT_COLUMN_ALIASES
from utils import (
    slugify_mode_name,
    infer_sql_type,
    normalize_cell_for_insert,
    parse_json_value,
)


def save_upload_to_db(
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
        # 插入会话记录
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
        insert_sql = sql.SQL("INSERT INTO {} (session_id, row_index, row_data) VALUES (%s, %s, %s)").format(table_name)
        
        for batch_start in range(0, len(rows), batch_size):
            batch_end = min(batch_start + batch_size, len(rows))
            batch_values = [
                (session_id, idx, json.dumps(rows[idx]))
                for idx in range(batch_start, batch_end)
            ]
            extras.execute_values(cur, insert_sql, batch_values, template=None, page_size=1000)
        
        conn.commit()
        logger.info(f"保存上传数据成功: session_id={session_id}, rows={row_count}")

    return session_id


def delete_upload_session(conn, session_id: int) -> bool:
    """删除指定上传会话及其关联数据。"""
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
        if deleted:
            conn.commit()
            logger.info(f"删除会话成功: session_id={session_id}")
            return True
        return False


def save_custom_mode_to_db(conn, mode_name: str, rows: list[dict[str, Any]], selected_columns: list[str]) -> dict:
    """保存自定义模式到数据库。"""
    if len(rows) > MAX_SAVE_ROWS:
        raise ValueError(f"单次最多保存 {MAX_SAVE_ROWS} 行")

    timestamp_suffix = datetime.now().strftime("%Y%m%d_%H%M%S")
    table_name = f"custom_mode_{slugify_mode_name(mode_name)}_{timestamp_suffix}"
    
    # 对列名进行安全处理
    def safe_column_name(col: str) -> str:
        safe_name = re.sub(r'[^\w\u4e00-\u9fff]', '_', col.strip())
        if safe_name and not (safe_name[0].isalpha() or safe_name[0] == '_' or '\u4e00' <= safe_name[0] <= '\u9fff'):
            safe_name = 'col_' + safe_name
        return safe_name or 'col_unknown'
    
    safe_column_mapping = {col: safe_column_name(col) for col in selected_columns}
    column_types = {
        safe_column_mapping[col]: infer_sql_type([row.get(col) for row in rows])
        for col in selected_columns
    }

    with conn.cursor() as cur:
        # 创建表
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

        # 插入数据
        insert_safe_cols = ["mode_name"] + [safe_column_mapping[col] for col in selected_columns]
        insert_stmt = sql.SQL("INSERT INTO {} ({}) VALUES ({})").format(
            sql.Identifier(table_name),
            sql.SQL(", ").join(sql.Identifier(c) for c in insert_safe_cols),
            sql.SQL(", ").join(sql.Placeholder() for _ in insert_safe_cols),
        )
        for row in rows:
            row_values = [mode_name]
            for col in selected_columns:
                safe_col = safe_column_mapping[col]
                row_values.append(normalize_cell_for_insert(row.get(col), column_types[safe_col]))
            cur.execute(insert_stmt, tuple(row_values))
        conn.commit()
        logger.info(f"保存自定义模式成功: table={table_name}, rows={len(rows)}")

    return {
        "table_name": table_name,
        "row_count": len(rows),
        "selected_columns": selected_columns,
        "column_types": {col: column_types[safe_column_mapping[col]] for col in selected_columns},
    }


def get_column_aliases_from_config(mapping_config: dict[str, Any] | None) -> dict[str, list[str]]:
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