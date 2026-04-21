"""
数据库连接管理
"""

import os
from typing import Generator
from psycopg import connect
from fastapi import HTTPException

from config.settings import logger, get_pg_dsn

def get_connection() -> Generator:
    """获取数据库连接"""
    dsn = get_pg_dsn()
    if not dsn:
        raise HTTPException(status_code=500, detail="缺少 PG_DSN 环境变量")
    return connect(dsn)