"""
应用配置管理
"""

import logging
import os
from pathlib import Path
from dotenv import load_dotenv

# 项目根目录
BASE_DIR = Path(__file__).resolve().parent.parent

# 加载环境变量
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

# 文件上传配置
MAX_FILE_SIZE = 50 * 1024 * 1024  # 50MB
MAX_FILE_SIZE_MB = 50
ALLOWED_EXTENSIONS = {".xlsx", ".xls"}

# 数据限制配置
MAX_PREVIEW_ROWS = 200
MAX_SAVE_ROWS = 5000

# 数据库表名
UPLOAD_SESSIONS_TABLE = "upload_sessions"
UPLOAD_DATA_TABLE = "upload_data"
COLUMN_MAPPING_TABLE = "column_mappings"

# 获取数据库连接串
def get_pg_dsn() -> str:
    """获取PostgreSQL连接串"""
    dsn = os.getenv("PG_DSN", "").strip()
    return dsn

def get_pg_dsn_or_raise() -> str:
    """获取PostgreSQL连接串，如果未配置则抛出异常"""
    dsn = get_pg_dsn()
    if not dsn:
        from fastapi import HTTPException
        raise HTTPException(status_code=500, detail="缺少 PG_DSN 环境变量，请先配置 PostgreSQL 连接串")
    return dsn