#!/bin/bash
# PostgreSQL 连接脚本
# 使用方法: ./connect.sh

# 切换到项目根目录
SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
PROJECT_DIR="$(dirname "$SCRIPT_DIR")"

# 读取 .env 文件中的 PG_DSN
if [ -f "$PROJECT_DIR/.env" ]; then
    # 提取 PG_DSN 的值
    PG_DSN=$(grep "^PG_DSN=" "$PROJECT_DIR/.env" | cut -d'=' -f2-)
    
    if [ -n "$PG_DSN" ]; then
        echo "正在连接 PostgreSQL..."
        echo "DSN: $PG_DSN"
        echo "---"
        psql "$PG_DSN"
    else
        echo "错误: .env 文件中未找到 PG_DSN 配置"
        exit 1
    fi
else
    echo "错误: 未找到 .env 文件: $PROJECT_DIR/.env"
    exit 1
fi