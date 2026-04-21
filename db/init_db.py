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


def init_database():
    """初始化数据库表结构"""
    dsn = os.getenv("PG_DSN", "").strip()
    
    if not dsn:
        print("错误: 未找到 PG_DSN 环境变量")
        print("请在项目根目录的 .env 文件中配置:")
        print("  PG_DSN=postgresql://user:password@host:port/database")
        sys.exit(1)
    
    # 安全地提取数据库连接信息（避免 tuple index out of range）
    try:
        if '@' in dsn:
            parts = dsn.split('@')
            if len(parts) >= 2:
                db_info = parts[1]
            else:
                db_info = dsn
        else:
            db_info = dsn
        print(f"连接数据库: {db_info}")
    except Exception:
        print(f"连接数据库...")
    
    print("-" * 50)
    
    try:
        from psycopg import connect
        
        with connect(dsn) as conn:
            # 导入并执行表结构初始化
            from db.schema import ensure_upload_tables_exist, check_db_health
            
            print("正在创建表结构...")
            ensure_upload_tables_exist(conn)
            
            print("-" * 50)
            print("检查数据库状态...")
            health = check_db_health(conn)
            
            print("-" * 50)
            if health["healthy"]:
                print("✅ 数据库初始化成功!")
                print("-" * 50)
                print("表统计:")
                for table_name, info in health["tables"].items():
                    exists_mark = "✓" if info.get("exists", True) else "✗"
                    count = info.get("count", 0)
                    print(f"  [{exists_mark}] {table_name}: {count} 条记录")
            else:
                print(f"⚠️ 数据库状态异常: {health['message']}")
                print("-" * 50)
                print("表统计:")
                for table_name, info in health["tables"].items():
                    exists_mark = "✓" if info.get("exists", True) else "✗"
                    count = info.get("count", 0)
                    error = info.get("error", "")
                    if error:
                        print(f"  [{exists_mark}] {table_name}: 错误 - {error}")
                    else:
                        print(f"  [{exists_mark}] {table_name}: {count} 条记录")
                print("-" * 50)
                print("提示: 请检查数据库连接或手动运行建表语句")
                sys.exit(1)
                
    except ImportError as e:
        print(f"错误: 无法导入模块 - {e}")
        print("请确保已安装 psycopg: pip install psycopg[binary]")
        sys.exit(1)
    except Exception as e:
        print(f"数据库初始化失败: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)


def main():
    print("=" * 50)
    print("J1 - Excel 分析平台数据库初始化")
    print("=" * 50)
    init_database()


if __name__ == "__main__":
    main()