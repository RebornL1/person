# PostgreSQL 连接指南

本目录提供通过终端命令连接 PostgreSQL 数据库的方法。

## 连接信息

根据项目 `.env` 文件配置：

```
主机: 127.0.0.1
端口: 5432
数据库: j1_analytics
用户: apple
```

## 方式一：使用 psql 命令行工具

### 1. 直接连接

```bash
psql -h 127.0.0.1 -p 5432 -U apple -d j1_analytics
```

### 2. 使用连接串

```bash
psql "postgresql://apple@127.0.0.1:5432/j1_analytics"
```

### 3. 从项目 .env 读取配置

```bash
# 进入项目根目录
cd /Users/apple/Documents/AI/J1

# 读取 .env 并连接
source .env && psql "$PG_DSN"
```

## 方式二：使用 Python 脚本

```bash
cd /Users/apple/Documents/AI/J1/pg
python3 connect.py
```

## 方式三：使用快捷脚本

```bash
cd /Users/apple/Documents/AI/J1/pg
./connect.sh
```

## 常用命令

连接成功后，可使用以下 psql 命令：

| 命令 | 说明 |
|------|------|
| `\l` | 列出所有数据库 |
| `\dt` | 列出当前数据库的所有表 |
| `\d 表名` | 查看表结构 |
| `\du` | 列出所有用户 |
| `\c 数据库名` | 切换数据库 |
| `\q` | 退出 |
| `\?` | 帮助 |

## 查看应用创建的表

应用保存的自定义模式表都以 `custom_mode_` 开头：

```sql
-- 查看所有自定义模式表
SELECT table_name 
FROM information_schema.tables 
WHERE table_schema = 'public' 
  AND table_name LIKE 'custom_mode_%';

-- 查看某个模式表的数据
SELECT * FROM custom_mode_xxx LIMIT 10;
```

## 环境要求

确保已安装：

1. **PostgreSQL 客户端工具**（psql）
   ```bash
   # macOS 使用 Homebrew 安装
   brew install libpq
   
   # 或安装完整 PostgreSQL
   brew install postgresql
   ```

2. **Python 依赖**（用于 Python 脚本方式）
   ```bash
   pip install psycopg[binary]
   ```

## 故障排查

### 连接被拒绝

检查 PostgreSQL 服务是否运行：

```bash
# macOS Homebrew
brew services list
brew services start postgresql

# 或使用 pg_ctl
pg_ctl -D /usr/local/var/postgres status
```

### 权限问题

确保用户有访问数据库的权限：

```sql
-- 以 postgres 超级用户登录后执行
GRANT ALL PRIVILEGES ON DATABASE j1_analytics TO apple;
```

### 数据库不存在

创建数据库：

```bash
createdb -U postgres j1_analytics
```