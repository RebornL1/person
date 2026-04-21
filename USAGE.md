# Excel 分析网页 — 使用指南

本文档说明如何在本地安装依赖、启动服务、访问页面与测试。**凡新增或变更启动方式、端口、依赖、接口或用户可见行为时，请同步更新本文件**（见文末维护约定）。

---

## 环境要求

- **Python**：推荐 **3.13**（项目内已用 `.venv` 验证）。若使用 Homebrew，可将 `python3` 指向 3.13（参见你本机 `~/.zshrc` 中的 `python@3.13` 配置）。
- **操作系统**：macOS / Linux / Windows 均可；以下为类 Unix 命令示例。
- **PostgreSQL（可选）**：若要启用"自定义模式入库"，需提供 `PG_DSN`（例如 `postgresql://user:password@127.0.0.1:5432/dbname`）。项目使用 **psycopg 3.x**（非 psycopg2）。
- **`.env` 配置**：应用会自动读取项目根目录 `.env`（例如 `PG_DSN=...`）。

---

## 首次安装

### 方式一：一键启动（推荐）

项目提供了跨平台启动脚本，首次运行会自动创建虚拟环境并安装依赖：

**Windows**：双击 `start.bat` 文件

**macOS/Linux**：
```bash
./start.sh
# 或
python3 start.py
```

### 方式二：手动安装

**macOS/Linux**：

```bash
cd /path/to/J1
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

**Windows (CMD)**：

```cmd
cd C:\path\to\J1
python -m venv .venv
.venv\Scripts\activate.bat
pip install -r requirements.txt
```

**Windows (PowerShell)**：

```powershell
cd C:\path\to\J1
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
```

---

## PostgreSQL 快速配置（用于自定义模式入库）

macOS（Homebrew）：

```bash
brew install postgresql@16
brew services start postgresql@16
createdb j1_analytics
```

Windows（建议使用 PostgreSQL 官方安装包）：

1. 安装 PostgreSQL（带 pgAdmin 可选）
2. 新建数据库 `j1_analytics`
3. 记录用户名、密码、端口（默认 `5432`）

在项目根目录创建 `.env`：

```bash
PG_DSN=postgresql://<user>:<password>@127.0.0.1:5432/j1_analytics
```

PowerShell 可用：

```powershell
Set-Content -Path .env -Value "PG_DSN=postgresql://<user>:<password>@127.0.0.1:5432/j1_analytics"
```

**初始化数据库表结构**（可选，服务启动时会自动创建）：

```bash
python db/init_db.py
```

---

## 启动网页服务

### 方式一：使用启动脚本（最简单）

**Windows**：双击 `start.bat`

**macOS/Linux**：
```bash
./start.sh
```

启动脚本参数（可选）：
```bash
python3 start.py --host 127.0.0.1 --port 8080
python3 start.py --no-reload   # 生产模式，禁用自动重载
python3 start.py --install     # 仅安装依赖
```

### 方式二：手动启动

**必须在项目根目录执行**，且建议先激活虚拟环境：

**macOS/Linux**：
```bash
cd /path/to/J1
source .venv/bin/activate
uvicorn app:app --reload --host 127.0.0.1 --port 8000
```

**Windows (CMD)**：
```cmd
cd C:\path\to\J1
.venv\Scripts\activate.bat
uvicorn app:app --reload --host 127.0.0.1 --port 8000
```

**Windows (PowerShell)**：
```powershell
cd C:\path\to\J1
.\.venv\Scripts\Activate.ps1
uvicorn app:app --reload --host 127.0.0.1 --port 8000
```

参数说明：
- **`--reload`**：改代码后自动重载（仅开发环境使用）
- **`--host 127.0.0.1`**：只本机可访问；若需局域网访问可改为 `0.0.0.0`（注意防火墙与安全风险）
- **`--port 8000`**：端口可改，例如 `--port 8080`

**不激活 venv 时**，可直接指定解释器：

**macOS/Linux**：
```bash
.venv/bin/python -m uvicorn app:app --reload --host 127.0.0.1 --port 8000
```

**Windows**：
```cmd
.venv\Scripts\python.exe -m uvicorn app:app --reload --host 127.0.0.1 --port 8000
```

启动成功后，终端会看到类似：`Uvicorn running on http://127.0.0.1:8000`。

---

## 访问地址

| 用途 | URL |
|------|-----|
| 上传与分析页面（主界面） | http://127.0.0.1:8000/ |
| API 文档（Swagger） | http://127.0.0.1:8000/docs |
| 备用 API 文档（ReDoc） | http://127.0.0.1:8000/redoc |

静态资源目录挂载在 **`/static`**（与 `static/` 文件夹对应）。

---

## 使用步骤

1. 按上文启动 `uvicorn`。
2. 浏览器打开 http://127.0.0.1:8000/ 。
3. 点击上传区或拖拽 **`.xlsx` / `.xls`** 文件。
4. **导入配置界面**（新功能）：
   - 选择要导入的 Sheet
   - 勾选需要的列（可多选）
   - 为每列设置图表类型：`bar`（柱状图）、`line`（折线图）、`pie`（饼图）、`table`（表格）
   - 设置各列的显示名称（可选）
   - 设置各列的数据类型：`text`、`numeric`、`datetime`
   - 选择列映射配置（可选，用于复用列名映射）
5. 点击「开始导入」后，页面展示：
   - **自定义图表展示区域**：按配置自动生成图表和表格
   - **GaussDB 现网工作量看板**：综合评分、透传排序、风险分层、问题结构、鼓励项产出
   - 行列统计、列类型与缺失、数值列描述统计、表格预览
6. 如需查看模型公式或调权重，点击「**打开模型口径与权重配置**」按钮，在弹窗中完成调整（会实时重算排名与图表）。
7. 如需保存自定义模式到 PG：输入模式名、勾选要保存的 Excel 列，点击「**保存当前模式到 PostgreSQL**」。
8. **历史数据管理**：
   - 点击历史记录的「加载」按钮可加载历史数据
   - 加载后会自动恢复当时的图表配置
   - 可通过日期选择器筛选历史数据
9. 若刚新增或修改 `.env`，请重启服务后再测试入库。

### 示例数据

项目内自带多 sheet 示例表：

- 路径：`samples/multi_sheet_test.xlsx`
- Sheet 内容：
  - `团队A工作量`：8人工作量数据
  - `团队B工作量`：8人工作量数据
  - `产品线X数据`：5人工作量数据（不同列名）
  - `项目进度表`：项目进度数据（非工作量类型）
- 单 sheet 示例：`samples/sample_10_people.xlsx`
- 重新生成：

  ```bash
  .venv/bin/python scripts/generate_sample_excel.py
  ```

---

## 自定义图表展示功能

上传 Excel 时可为每列配置不同的展示方式：

| 图表类型 | 说明 | 适用数据 |
|----------|------|----------|
| `bar` | 柱状图 | 数值型数据，按姓名分组统计 |
| `line` | 折线图 | 数值型数据，适合趋势展示 |
| `pie` | 饼图 | 数值型数据，占比展示 |
| `table` | 数据表格 | 文本型数据或需要查看原始值 |
| `none` | 不展示 | 不需要展示的列 |

**注意事项**：
- 数值型列会自动按「姓名」列分组统计
- 文本型列会统计各值出现次数（最多展示10个分类）
- 图表配置会随数据一起保存到数据库，加载历史数据时自动恢复

---

## 接口说明

### 文件上传相关

| 方法 | 路径 | 说明 |
|------|------|------|
| `POST` | `/api/upload/preview` | 预览 Excel，返回所有 sheet 信息（列名、类型、预览行） |
| `POST` | `/api/upload` | 上传 Excel 并保存，支持参数：`sheet_name`、`selected_columns`、`chart_types`、`display_names`、`column_types`、`column_mapping_id` |
| `GET` | `/api/upload/history` | 获取历史上传记录，按日期分组返回 |
| `GET` | `/api/upload/latest` | 获取最新上传数据（含 `chart_type_config`、`display_name_config`） |
| `GET` | `/api/upload/session/{id}` | 获取指定会话数据（含图表配置） |
| `POST` | `/api/upload/delete/{id}` | 删除指定上传会话及其数据 |

### 自定义模式相关

| 方法 | 路径 | 说明 |
|------|------|------|
| `GET` | `/api/custom-mode/list` | 获取所有已保存模式 |
| `GET` | `/api/custom-mode/load/{table_name}` | 加载指定模式数据 |
| `POST` | `/api/custom-mode/save` | 保存自定义模式（创建独立表） |
| `POST` | `/api/custom-mode/delete` | 删除模式（删除对应表） |

### 列映射配置相关

| 方法 | 路径 | 说明 |
|------|------|------|
| `GET` | `/api/column-mapping/list` | 获取列映射配置列表 |
| `POST` | `/api/column-mapping/save` | 保存列映射配置 |
| `POST` | `/api/column-mapping/delete` | 删除列映射配置 |

### 数据库相关

| 方法 | 路径 | 说明 |
|------|------|------|
| `GET` | `/api/db/health` | 检查数据库连接和表状态 |

详细请求/响应可在 http://127.0.0.1:8000/docs 中查看与调试。

---

## 数据库表结构

项目使用 PostgreSQL 存储上传数据和自定义模式。主要表结构：

### upload_sessions（上传会话表）

| 字段 | 类型 | 说明 |
|------|------|------|
| `id` | BIGSERIAL | 主键 |
| `upload_date` | DATE | 上传日期 |
| `upload_time` | TIMESTAMPTZ | 上传时间 |
| `filename` | TEXT | 文件名 |
| `row_count` | BIGINT | 行数 |
| `col_count` | BIGINT | 列数 |
| `columns_json` | JSONB | 列名数组 |
| `chart_types` | JSONB | 图表类型配置 |
| `display_names` | JSONB | 显示名称配置 |
| `column_types` | JSONB | 列类型配置 |
| `sheet_name` | TEXT | Sheet 名称 |
| `selected_columns` | TEXT | 选中的列 |

### upload_data（上传数据表）

| 字段 | 类型 | 说明 |
|------|------|------|
| `id` | BIGSERIAL | 主键 |
| `session_id` | BIGINT | 会话 ID |
| `row_index` | BIGINT | 行序号 |
| `row_data` | JSONB | 行数据 |

---

## 生产环境提示（可选）

开发命令中的 `--reload` **不要**用于生产。可改用多 worker（需先去掉 reload，并按部署平台调整）：

```bash
.venv/bin/uvicorn app:app --host 0.0.0.0 --port 8000 --workers 2
```

实际生产建议配合反向代理（Nginx/Caddy）、HTTPS 与进程守护（systemd、Docker 等），按你的运维标准配置。

---

## 常见问题

| 现象 | 处理 |
|------|------|
| `Address already in use` | 更换端口：`--port 8080`，或结束占用 8000 的进程。 |
| `ModuleNotFoundError: fastapi` | 确认已 `source .venv/bin/activate` 且执行过 `pip install -r requirements.txt`。 |
| 上传后报错「无法解析」 | 确认文件为有效 Excel；`.xlsx` 依赖 `openpyxl`，`.xls` 依赖 `xlrd`。 |
| 保存自定义模式时报 PG 连接错误 | 确认已设置 `PG_DSN`，且数据库用户有 `CREATE TABLE` 与 `INSERT` 权限。 |
| 修改 `.env` 后仍提示缺少 `PG_DSN` | 重启 `uvicorn` 进程（已运行进程不会自动读取新环境变量）。 |
| 图表不显示数据 | 检查浏览器控制台(F12)是否有错误；确认 Chart.js CDN 加载成功。 |
| 历史数据加载后图表无内容 | 已修复：API 现正确返回 `chart_type_config`。刷新页面重试。 |
| `tuple index out of range` 错误 | 已修复：JSONB 默认值转义问题。运行 `python db/init_db.py` 重建表。 |

---

## 工作量评分模型（当前版本）

综合工作量分（每人）使用如下方向：

- 正相关：`oncall未闭环`、`待处理工单`、`昨日新增问题`
- 正相关：`管控/内核/咨询`，且 **内核权重更高**
- 负相关：`透传求助`
- 励项（正向）：`问题单`、`需求单`、`wiki`、`问题分析报告`
- 约束：**昨日新增问题正向权重大于透传求助负向权重**
- 派生指标：`每日的问题数量 = 管控 + 内核 + 咨询`
- 新增风险层：基于 `透传求助 + 待处理工单 + oncall未闭环` 计算风险分，并分为 `high / medium / low`
- 新增现网聚焦指数：`gaussdb_focus_index`，强调内核问题处置与知识沉淀，同时将透传求助作为负向项

> 具体权重以 `services/workload_service.py` 中返回的 `weights` 为准，页面会直接展示当前模型公式。
> 页面支持在弹窗内直接调节权重（不改源数据），会实时重算排名、图表和风险看板，用于模拟不同运维策略下的团队表现。

---

## 项目结构（与运行相关）

```
J1/
├── app.py                 # FastAPI 应用入口
├── models.py              # Pydantic 模型定义
├── utils.py               # 工具函数
├── start.py               # 跨平台启动脚本（Python）
├── start.bat              # Windows 快捷启动
├── start.sh               # macOS/Linux 快捷启动
├── requirements.txt       # Python 依赖
├── config/
│   ├── settings.py        # 配置管理（日志、常量、DSN）
├── db/
│   ├── connection.py      # 数据库连接
│   ├── schema.py          # 表结构管理
│   ├── init_db.py         # 初始化脚本
├── services/
│   ├── upload_service.py  # 上传服务
│   ├── workload_service.py # 工作量分析
├── static/
│   ├── index.html         # 前端页面
│   ├── app.js             # 前端逻辑
│   └── styles.css         # 样式
├── samples/               # 示例 Excel
├── scripts/               # 脚本工具
├── pg/
│   └── connect.py         # PG 连接测试
├── .env                   # 环境变量
├── USAGE.md               # 本指南
└── README.md              # 项目说明
```

---

## 维护约定（给开发者与 AI 助手）

- 变更**启动命令、端口、环境变量、依赖、路由、上传限制、预览行数**或任何**用户需要在指南里知道的行为**时，**同时修改 `USAGE.md`**。
- 仓库内已配置 Cursor 规则：在改动本项目的 `app.py`、`static/`、`requirements.txt`、`scripts/` 等时，应检查并更新本指南，避免文档与代码脱节。

---

*文档版本与 `config/settings.py` 中 `MAX_PREVIEW_ROWS = 200` 及当前路由保持一致；若不一致，以代码为准并请及时改文档。*