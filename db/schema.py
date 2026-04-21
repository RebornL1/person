"""
数据库表结构管理
"""

import json
from psycopg import sql

from config.settings import (
    logger,
    UPLOAD_SESSIONS_TABLE,
    UPLOAD_DATA_TABLE,
    COLUMN_MAPPING_TABLE,
)
from models import DEFAULT_COLUMN_ALIASES

def ensure_upload_tables_exist(conn) -> None:
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
                has_analysis BOOLEAN DEFAULT TRUE,
                sheet_name TEXT,
                selected_columns TEXT,
                display_names JSONB,
                column_types JSONB,
                chart_types JSONB,
                notes TEXT,
                column_mapping_id BIGINT,
                config_name TEXT
            )
            """).format(sql.Identifier(UPLOAD_SESSIONS_TABLE))
        )
        # 添加新字段（如果表已存在）
        try:
            cur.execute(sql.SQL("ALTER TABLE {} ADD COLUMN IF NOT EXISTS has_analysis BOOLEAN DEFAULT TRUE").format(sql.Identifier(UPLOAD_SESSIONS_TABLE)))
            cur.execute(sql.SQL("ALTER TABLE {} ADD COLUMN IF NOT EXISTS sheet_name TEXT").format(sql.Identifier(UPLOAD_SESSIONS_TABLE)))
            cur.execute(sql.SQL("ALTER TABLE {} ADD COLUMN IF NOT EXISTS selected_columns TEXT").format(sql.Identifier(UPLOAD_SESSIONS_TABLE)))
            cur.execute(sql.SQL("ALTER TABLE {} ADD COLUMN IF NOT EXISTS display_names JSONB").format(sql.Identifier(UPLOAD_SESSIONS_TABLE)))
            cur.execute(sql.SQL("ALTER TABLE {} ADD COLUMN IF NOT EXISTS column_types JSONB").format(sql.Identifier(UPLOAD_SESSIONS_TABLE)))
            cur.execute(sql.SQL("ALTER TABLE {} ADD COLUMN IF NOT EXISTS chart_types JSONB").format(sql.Identifier(UPLOAD_SESSIONS_TABLE)))
            cur.execute(sql.SQL("ALTER TABLE {} ADD COLUMN IF NOT EXISTS config_name TEXT").format(sql.Identifier(UPLOAD_SESSIONS_TABLE)))
        except Exception as e:
            logger.warning(f"添加新字段时出错（可能已存在）: {e}")
        
        # 上传数据表：存储每次上传的具体数据行
        cur.execute(
            sql.SQL("""
            CREATE TABLE IF NOT EXISTS {} (
                id BIGSERIAL PRIMARY KEY,
                session_id BIGINT NOT NULL,
                row_index BIGINT NOT NULL,
                row_data JSONB NOT NULL
            )
            """).format(sql.Identifier(UPLOAD_DATA_TABLE))
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
            SELECT '默认配置', %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, TRUE
            WHERE NOT EXISTS (SELECT 1 FROM {} WHERE mapping_name = '默认配置')
            """).format(sql.Identifier(COLUMN_MAPPING_TABLE), sql.Identifier(COLUMN_MAPPING_TABLE)),
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