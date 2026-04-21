"""
数据库表结构管理

设计原则：
- 无外键约束：避免级联删除问题，提高兼容性
- 无唯一约束：防止插入冲突，使用逻辑判断替代
- 使用 IF NOT EXISTS：确保幂等性，可重复执行
- 所有字段都有默认值：避免插入时必填字段报错
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
    """
    确保所有必要的表都存在。
    在全新数据库中执行不会报错，可重复执行。
    """
    with conn.cursor() as cur:
        # ========== 1. 上传会话表 ==========
        # 记录每次上传的元信息
        # 注意：JSONB 默认值使用双花括号 {{}} 转义，避免被 format() 解析为占位符
        cur.execute(
            sql.SQL("""
            CREATE TABLE IF NOT EXISTS {} (
                id BIGSERIAL PRIMARY KEY,
                upload_date DATE NOT NULL DEFAULT CURRENT_DATE,
                upload_time TIMESTAMPTZ NOT NULL DEFAULT NOW(),
                filename TEXT DEFAULT '',
                row_count BIGINT NOT NULL DEFAULT 0,
                col_count BIGINT NOT NULL DEFAULT 0,
                columns_json JSONB DEFAULT '{{}}',
                has_workload_analysis BOOLEAN DEFAULT FALSE,
                has_analysis BOOLEAN DEFAULT TRUE,
                sheet_name TEXT DEFAULT '',
                selected_columns TEXT DEFAULT '',
                display_names JSONB DEFAULT '{{}}',
                column_types JSONB DEFAULT '{{}}',
                chart_types JSONB DEFAULT '{{}}',
                notes TEXT DEFAULT '',
                column_mapping_id BIGINT DEFAULT NULL,
                config_name TEXT DEFAULT ''
            )
            """).format(sql.Identifier(UPLOAD_SESSIONS_TABLE))
        )
        
        # ========== 2. 上传数据表 ==========
        # 存储每次上传的具体数据行（无外键约束）
        cur.execute(
            sql.SQL("""
            CREATE TABLE IF NOT EXISTS {} (
                id BIGSERIAL PRIMARY KEY,
                session_id BIGINT NOT NULL DEFAULT 0,
                row_index BIGINT NOT NULL DEFAULT 0,
                row_data JSONB DEFAULT '{{}}'
            )
            """).format(sql.Identifier(UPLOAD_DATA_TABLE))
        )
        
        # ========== 3. 列映射配置表 ==========
        # 存储自定义列名映射（无唯一约束）
        cur.execute(
            sql.SQL("""
            CREATE TABLE IF NOT EXISTS {} (
                id BIGSERIAL PRIMARY KEY,
                mapping_name TEXT NOT NULL DEFAULT '',
                name_aliases JSONB DEFAULT '[]',
                oncall_open_aliases JSONB DEFAULT '[]',
                pending_ticket_aliases JSONB DEFAULT '[]',
                new_issue_yesterday_aliases JSONB DEFAULT '[]',
                governance_issue_aliases JSONB DEFAULT '[]',
                kernel_issue_aliases JSONB DEFAULT '[]',
                consult_issue_aliases JSONB DEFAULT '[]',
                escalation_help_aliases JSONB DEFAULT '[]',
                issue_ticket_output_aliases JSONB DEFAULT '[]',
                requirement_ticket_output_aliases JSONB DEFAULT '[]',
                wiki_output_aliases JSONB DEFAULT '[]',
                analysis_report_output_aliases JSONB DEFAULT '[]',
                is_default BOOLEAN DEFAULT FALSE,
                created_at TIMESTAMPTZ NOT NULL DEFAULT NOW(),
                updated_at TIMESTAMPTZ NOT NULL DEFAULT NOW()
            )
            """).format(sql.Identifier(COLUMN_MAPPING_TABLE))
        )
        
        # ========== 4. 创建索引（加速查询）==========
        # 上传会话表索引
        cur.execute(
            sql.SQL("CREATE INDEX IF NOT EXISTS idx_upload_sessions_date ON {} (upload_date DESC)").format(
                sql.Identifier(UPLOAD_SESSIONS_TABLE)
            )
        )
        cur.execute(
            sql.SQL("CREATE INDEX IF NOT EXISTS idx_upload_sessions_time ON {} (upload_time DESC)").format(
                sql.Identifier(UPLOAD_SESSIONS_TABLE)
            )
        )
        
        # 上传数据表索引
        cur.execute(
            sql.SQL("CREATE INDEX IF NOT EXISTS idx_upload_data_session ON {} (session_id)").format(
                sql.Identifier(UPLOAD_DATA_TABLE)
            )
        )
        cur.execute(
            sql.SQL("CREATE INDEX IF NOT EXISTS idx_upload_data_row_index ON {} (session_id, row_index)").format(
                sql.Identifier(UPLOAD_DATA_TABLE)
            )
        )
        
        # 列映射配置表索引
        cur.execute(
            sql.SQL("CREATE INDEX IF NOT EXISTS idx_column_mappings_name ON {} (mapping_name)").format(
                sql.Identifier(COLUMN_MAPPING_TABLE)
            )
        )
        cur.execute(
            sql.SQL("CREATE INDEX IF NOT EXISTS idx_column_mappings_default ON {} (is_default)").format(
                sql.Identifier(COLUMN_MAPPING_TABLE)
            )
        )
        
        # ========== 5. 插入默认列映射配置 ==========
        # 使用逻辑判断替代唯一约束，检查是否已存在默认配置
        cur.execute(
            sql.SQL("SELECT COUNT(*) FROM {} WHERE is_default = TRUE").format(
                sql.Identifier(COLUMN_MAPPING_TABLE)
            )
        )
        row = cur.fetchone()
        default_count = row[0] if row else 0
        
        if default_count == 0:
            # 插入默认配置
            cur.execute(
                sql.SQL("""
                INSERT INTO {} (
                    mapping_name,
                    name_aliases,
                    oncall_open_aliases,
                    pending_ticket_aliases,
                    new_issue_yesterday_aliases,
                    governance_issue_aliases,
                    kernel_issue_aliases,
                    consult_issue_aliases,
                    escalation_help_aliases,
                    issue_ticket_output_aliases,
                    requirement_ticket_output_aliases,
                    wiki_output_aliases,
                    analysis_report_output_aliases,
                    is_default
                ) VALUES (
                    '默认配置',
                    %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s,
                    TRUE
                )
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
            logger.info("已插入默认列映射配置")
        
        conn.commit()
        logger.info("数据库表结构初始化完成")


def get_table_list(conn) -> list[str]:
    """获取所有应用相关的表列表。"""
    with conn.cursor() as cur:
        cur.execute(
            """
            SELECT table_name FROM information_schema.tables
            WHERE table_schema = 'public'
            AND table_name IN (%s, %s, %s)
            OR table_name LIKE 'custom_mode_%'
            ORDER BY table_name
            """,
            (UPLOAD_SESSIONS_TABLE, UPLOAD_DATA_TABLE, COLUMN_MAPPING_TABLE)
        )
        rows = cur.fetchall()
        return [r[0] for r in rows] if rows else []


def check_db_health(conn) -> dict:
    """检查数据库健康状态，返回各表的统计信息。"""
    result = {
        "tables": {},
        "healthy": True,
        "message": "数据库正常"
    }
    
    try:
        with conn.cursor() as cur:
            # 检查上传会话表
            try:
                cur.execute(
                    sql.SQL("SELECT COUNT(*) FROM {}").format(
                        sql.Identifier(UPLOAD_SESSIONS_TABLE)
                    )
                )
                row = cur.fetchone()
                result["tables"][UPLOAD_SESSIONS_TABLE] = {"count": row[0] if row else 0, "exists": True}
            except Exception as e:
                result["tables"][UPLOAD_SESSIONS_TABLE] = {"count": 0, "exists": False, "error": str(e)}
                result["healthy"] = False
            
            # 检查上传数据表
            try:
                cur.execute(
                    sql.SQL("SELECT COUNT(*) FROM {}").format(
                        sql.Identifier(UPLOAD_DATA_TABLE)
                    )
                )
                row = cur.fetchone()
                result["tables"][UPLOAD_DATA_TABLE] = {"count": row[0] if row else 0, "exists": True}
            except Exception as e:
                result["tables"][UPLOAD_DATA_TABLE] = {"count": 0, "exists": False, "error": str(e)}
                result["healthy"] = False
            
            # 检查列映射配置表
            try:
                cur.execute(
                    sql.SQL("SELECT COUNT(*) FROM {}").format(
                        sql.Identifier(COLUMN_MAPPING_TABLE)
                    )
                )
                row = cur.fetchone()
                result["tables"][COLUMN_MAPPING_TABLE] = {"count": row[0] if row else 0, "exists": True}
            except Exception as e:
                result["tables"][COLUMN_MAPPING_TABLE] = {"count": 0, "exists": False, "error": str(e)}
                result["healthy"] = False
            
            # 检查自定义模式表数量
            try:
                cur.execute(
                    """
                    SELECT COUNT(*) FROM information_schema.tables
                    WHERE table_schema = 'public' AND table_name LIKE 'custom_mode_%'
                    """
                )
                row = cur.fetchone()
                result["tables"]["custom_modes"] = {"count": row[0] if row else 0, "exists": True}
            except Exception as e:
                result["tables"]["custom_modes"] = {"count": 0, "exists": False, "error": str(e)}
            
            # 更新消息
            if result["healthy"]:
                result["message"] = "数据库正常，所有表都已创建"
            else:
                missing_tables = [name for name, info in result["tables"].items() 
                                  if not info.get("exists", True)]
                result["message"] = f"缺少表: {', '.join(missing_tables)}"
                
    except Exception as e:
        result["healthy"] = False
        result["message"] = f"数据库检查失败: {e}"
        logger.error(f"数据库健康检查失败: {e}")
    
    return result