#!/usr/bin/env python3
"""
PostgreSQL 连接测试脚本
使用方法: python3 connect.py
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


def main():
    dsn = os.getenv("PG_DSN", "").strip()
    
    if not dsn:
        print("错误: 未找到 PG_DSN 环境变量")
        print("请在项目根目录的 .env 文件中配置:")
        print("  PG_DSN=postgresql://user:password@host:port/database")
        sys.exit(1)
    
    print(f"DSN: {dsn}")
    print("-" * 50)
    
    try:
        from psycopg import connect, sql
        
        with connect(dsn) as conn:
            with conn.cursor() as cur:
                # 获取 PostgreSQL 版本
                cur.execute("SELECT version();")
                version = cur.fetchone()[0]
                print(f"PostgreSQL 版本: {version}")
                
                # 获取当前数据库
                cur.execute("SELECT current_database();")
                db = cur.fetchone()[0]
                print(f"当前数据库: {db}")
                
                # 获取当前用户
                cur.execute("SELECT current_user;")
                user = cur.fetchone()[0]
                print(f"当前用户: {user}")
                
                # 列出所有自定义模式表
                cur.execute("""
                    SELECT table_name 
                    FROM information_schema.tables 
                    WHERE table_schema = 'public' 
                      AND table_name LIKE 'custom_mode_%%'
                    ORDER BY table_name;
                """)
                tables = cur.fetchall()
                
                print("-" * 50)
                if tables:
                    print(f"自定义模式表 ({len(tables)} 个):")
                    for t in tables:
                        # 使用 sql.Identifier 防止 SQL 注入
                        cur.execute(
                            sql.SQL("SELECT COUNT(*) FROM {}").format(sql.Identifier(t[0]))
                        )
                        count = cur.fetchone()[0]
                        print(f"  - {t[0]} ({count} 行)")
                else:
                    print("暂无自定义模式表")
                    
                print("-" * 50)
                print("连接成功!")
                
    except ImportError:
        print("错误: 未安装 psycopg 库")
        print("请运行: pip install psycopg[binary]")
        sys.exit(1)
    except Exception as e:
        print(f"连接失败: {e}")
        sys.exit(1)


if __name__ == "__main__":
    main()