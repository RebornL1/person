#!/usr/bin/env python3
"""
数据库初始化脚本

用于在全新的 PostgreSQL 数据库中创建所有必要的表结构。
可以在启动应用前单独运行，确保数据库已准备好。

使用方法:
    python db/init_db.py

或者配合环境变量:
    PG_DSN=postgresql://user:password@host:port/database python db/init_db.py
"""

import os
import sys
from pathlib import Path

# 添加项目根目录到路径
PROJECT_DIR = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(PROJECT_DIR))

from dotenv import load_dotenv

# 加载 .env
load_dotenv(PROJECT_DIR / ".env")

from psycopg import connect


def init_database():
    """初始化数据库表结构"""
    dsn = os.getenv("PG_DSN", "").strip()
    
    if not dsn:
        print("错误: 未找到 PG_DSN 环境变量")
        print("请在项目根目录的 .env 文件中配置:")
        print("  PG_DSN=postgresql://user:password@host:port/database")
        sys.exit(1)
    
    print(f"连接数据库: {dsn.split('@')[1] if '@' in dsn else dsn}")
    print("-" * 50)
    
    try:
        with connect(dsn) as conn:
            # 导入并执行表结构初始化
            from db.schema import ensure_upload_tables_exist, check_db_health
            
            print("正在创建表结构...")
            ensure_upload_tables_exist(conn)
            
            print("-" * 50)
            print("检查数据库状态...")
            health = check_db_health(conn)
            
            if health["healthy"]:
                print("数据库初始化成功!")
                print("-" * 50)
                for table_name, info in health["tables"].items():
                    print(f"  - {table_name}: {info['count']} 条记录")
            else:
                print(f"数据库检查失败: {health['message']}")
                sys.exit(1)
                
    except ImportError as e:
        print(f"错误: 无法导入模块 - {e}")
        print("请确保已安装 psycopg: pip install psycopg[binary]")
        sys.exit(1)
    except Exception as e:
        print(f"数据库初始化失败: {e}")
        sys.exit(1)


def main():
    print("=" * 50)
    print("J1 - Excel 分析平台数据库初始化")
    print("=" * 50)
    init_database()


if __name__ == "__main__":
    main()