"""
PostgreSQL 连接示例脚本。

可直接运行：
    source .venv/bin/activate
    python scripts/pg_demo.py

或作为模块导入：
    from scripts.pg_demo import get_connection, list_tables, query_table
"""

from __future__ import annotations

import os
from pathlib import Path
from typing import Any

from dotenv import load_dotenv
from psycopg import connect

# 加载环境变量
BASE_DIR = Path(__file__).resolve().parents[1]
load_dotenv(BASE_DIR / ".env")


def get_dsn() -> str:
    """获取 PostgreSQL 连接串。"""
    dsn = os.getenv("PG_DSN", "").strip()
    if not dsn:
        raise ValueError("缺少 PG_DSN 环境变量，请在 .env 中配置")
    return dsn


def get_connection() -> Any:
    """获取数据库连接。"""
    return connect(get_dsn())


def list_tables() -> list[str]:
    """列出数据库中所有表。"""
    with get_connection() as conn:
        with conn.cursor() as cur:
            cur.execute(
                """
                SELECT table_name
                FROM information_schema.tables
                WHERE table_schema = 'public'
                ORDER BY table_name;
                """
            )
            return [row[0] for row in cur.fetchall()]


def get_table_columns(table_name: str) -> list[dict[str, str]]:
    """获取表的列信息。"""
    with get_connection() as conn:
        with conn.cursor() as cur:
            cur.execute(
                """
                SELECT column_name, data_type, is_nullable
                FROM information_schema.columns
                WHERE table_name = %s AND table_schema = 'public'
                ORDER BY ordinal_position;
                """,
                (table_name,),
            )
            return [
                {"name": row[0], "type": row[1], "nullable": row[2] == "YES"}
                for row in cur.fetchall()
            ]


def get_table_count(table_name: str) -> int:
    """获取表的行数。"""
    with get_connection() as conn:
        with conn.cursor() as cur:
            cur.execute(f'SELECT COUNT(*) FROM "{table_name}";')
            return cur.fetchone()[0]


def query_table(table_name: str, limit: int = 10) -> list[dict[str, Any]]:
    """查询表数据。"""
    with get_connection() as conn:
        with conn.cursor() as cur:
            cur.execute(f'SELECT * FROM "{table_name}" LIMIT {limit};')
            columns = [desc[0] for desc in cur.description]
            rows = cur.fetchall()
            return [dict(zip(columns, row)) for row in rows]


def drop_table(table_name: str) -> bool:
    """删除表。"""
    with get_connection() as conn:
        with conn.cursor() as cur:
            cur.execute(f'DROP TABLE IF EXISTS "{table_name}";')
            conn.commit()
            return True


def print_db_info() -> None:
    """打印数据库概览信息。"""
    print(f"连接串: {get_dsn()}")
    print()

    tables = list_tables()
    if not tables:
        print("数据库中没有表")
        return

    print(f"共有 {len(tables)} 个表:")
    for table in tables:
        count = get_table_count(table)
        print(f"  - {table} ({count} 行)")

    print()

    for table in tables:
        print(f"=== {table} ===")
        columns = get_table_columns(table)
        print("列:")
        for col in columns:
            nullable = "可空" if col["nullable"] else "非空"
            print(f"  - {col['name']} ({col['type']}, {nullable})")

        rows = query_table(table, limit=5)
        if rows:
            print("数据预览:")
            for row in rows:
                print(f"  {row}")
        print()


if __name__ == "__main__":
    print_db_info()