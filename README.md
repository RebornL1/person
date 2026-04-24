# J1 - Excel 工作量分析 Web 应用

一个基于 FastAPI 的 Excel 工作量分析工具，支持上传表格、自动评分、风险分层、数据入库和自定义图表展示。

## 功能特性

- **Excel 上传与预览** - 支持 `.xlsx` / `.xls` 格式，最多预览 200 行
- **数据导入配置** - 多 sheet 选择、列勾选、显示名称配置、列类型配置
- **多日期数据汇总** - 自动识别多个sheet为日期数据，按姓名聚合汇总，计算平均值
- **日期智能识别** - 支持多种日期格式识别（xx月xx日、MM-DD、MMDD、YYYYMMDD等）
- **自定义图表展示** - 支持柱状图(bar)、折线图(line)、饼图(pie)、表格(table)四种展示类型
- **工作量评分模型** - 综合评分、透传排序、风险分层、问题结构分析
- **智能风险预测** - 自动分析高压人员、透传异常、出勤问题、内核问题集中等风险点
- **团队建议生成** - 基于整体数据生成团队层面的改进建议
- **GaussDB 现网聚焦指数** - 强调内核问题处置与知识沉淀能力
- **风险建议** - 自动生成运维改进建议
- **自动数据入库** - 上传时自动保存到 PostgreSQL，按日期整理历史数据
- **历史数据管理** - 支持加载最新数据、查看历史记录、删除历史会话
- **日期筛选功能** - 日期选择器 + 快捷按钮（近7天/30天/90天）筛选历史数据
- **自定义模式入库** - 保存分析结果到 PostgreSQL（可选）
- **列映射配置** - 编辑和保存列名映射关系
- **权重配置弹窗** - 实时调整权重，模拟不同运维策略
- **入库状态提示** - 上传后显示数据入库状态（成功/失败/未配置）

## 最近更新

| 日期 | 更新内容 |
|------|----------|
| 2026-04-23 | 新增多日期智能识别功能（支持xx月xx日、MM-DD、MMDD、YYYYMMDD等格式） |
| 2026-04-23 | 新增智能风险预测功能（高压人员、透传异常、出勤问题、内核问题集中分析） |
| 2026-04-23 | 新增团队建议生成功能（基于整体数据的改进建议） |
| 2026-04-23 | 改进空单元格处理（NaN按数字0处理）和姓名聚合逻辑（不依赖行位置） |
| 2026-04-21 | 修复自定义图表渲染问题：Session/Latest API 缺少图表配置字段、Canvas DOM 更新时序问题 |
| 2026-04-21 | 重构代码结构：拆分 config、db、services 模块，改进数据库连接和表结构管理 |
| 2026-04-21 | 修复数据库初始化：移除唯一/外键约束，修复 JSONB 默认值转义问题 |
| 2026-04-17 | 完善数据导入配置功能：添加列映射编辑、每个图表单独配置展示类型 |
| 2026-04-16 | 添加数据导入配置功能：支持多sheet选择、列勾选、图表类型配置 |
| 2026-04-16 | 核心指标侧边栏优化：宽度240px、默认打开、添加关闭按钮、内容左右布局 |
| 2026-04-16 | 将核心指标总览改为左侧可折叠侧边栏，指标竖向陈列，支持点击展开/收起 |
| 2026-04-16 | 修复上传功能 JavaScript 错误，添加入库状态显示 |
| 2026-04-16 | 添加 favicon 图标，修复日期管理功能事件绑定 |
| 2026-04-16 | 修复数据库表结构，添加 column_mapping_id 列 |

## 技术栈

| 类别 | 技术 |
|------|------|
| 后端 | FastAPI + Uvicorn |
| 数据处理 | Pandas + openpyxl + xlrd |
| 数据库 | PostgreSQL（可选，使用 psycopg 3.x） |
| 前端 | 原生 HTML/CSS/JS + Chart.js |
| 图表 | Chart.js 4.4.3 |

## 快速开始

### 方式一：一键启动（推荐）

**Windows**：双击 `start.bat` 文件

**macOS/Linux**：
```bash
# 方式1：双击 start.sh 或运行
./start.sh

# 方式2：直接运行Python脚本
python3 start.py
```

首次运行会自动创建虚拟环境并安装依赖。

### 方式二：手动安装

**macOS/Linux**：
```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
uvicorn app:app --reload --host 127.0.0.1 --port 8000
```

**Windows (CMD)**：
```cmd
python -m venv .venv
.venv\Scripts\activate.bat
pip install -r requirements.txt
uvicorn app:app --reload --host 127.0.0.1 --port 8000
```

**Windows (PowerShell)**：
```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
uvicorn app:app --reload --host 127.0.0.1 --port 8000
```

启动后访问 http://127.0.0.1:8000/

## 使用方法

1. 打开页面，上传 `.xlsx` / `.xls` 文件
2. 在导入配置界面：
   - 选择 Sheet
   - 勾选需要的列
   - 设置各列的图表类型（bar/line/pie/table）
   - 设置显示名称和列类型
3. 查看自定义图表展示区域（按配置自动生成图表）
4. 查看工作量看板（综合评分、风险分层、透传排名等）
5. 点击「打开模型口径与权重配置」调整权重
6. 可选：保存自定义模式到 PostgreSQL

**示例数据**：`samples/multi_sheet_test.xlsx` 包含多个 sheet 可用于测试。

## API 接口

| 方法 | 路径 | 说明 |
|------|------|------|
| GET | `/` | 主页面 |
| POST | `/api/upload/preview` | 预览 Excel，返回所有 sheet 信息 |
| POST | `/api/upload` | 上传 Excel，自动保存到数据库 |
| GET | `/api/upload/history` | 获取历史上传记录（按日期分组） |
| GET | `/api/upload/latest` | 获取最新一次上传的数据（含图表配置） |
| GET | `/api/upload/session/{id}` | 获取指定会话的数据（含图表配置） |
| POST | `/api/upload/delete/{id}` | 删除指定上传会话 |
| GET | `/api/custom-mode/list` | 获取已保存模式 |
| GET | `/api/custom-mode/load/{table_name}` | 加载指定模式数据 |
| POST | `/api/custom-mode/save` | 保存自定义模式 |
| POST | `/api/custom-mode/delete` | 删除模式 |
| GET | `/api/column-mapping/list` | 获取列映射配置列表 |
| POST | `/api/column-mapping/save` | 保存列映射配置 |
| POST | `/api/column-mapping/delete` | 删除列映射配置 |
| GET | `/api/db/health` | 检查数据库健康状态 |

API 文档：http://127.0.0.1:8000/docs

## 项目结构

```
J1/
├── app.py                 # FastAPI 应用入口
├── models.py              # Pydantic 模型定义
├── utils.py               # 工具函数
├── start.py               # 跨平台启动脚本
├── start.bat              # Windows 快捷启动
├── start.sh               # macOS/Linux 快捷启动
├── requirements.txt       # Python 依赖
├── config/
│   ├── __init__.py
│   └── settings.py        # 配置管理（日志、常量、DSN）
├── db/
│   ├── __init__.py
│   ├── connection.py      # 数据库连接管理
│   ├── schema.py          # 表结构定义和健康检查
│   └── init_db.py         # 数据库初始化脚本
├── services/
│   ├── __init__.py
│   ├── upload_service.py  # 上传数据保存服务
│   └── workload_service.py # 工作量分析服务
├── routes/
│   └── __init__.py        # 路由模块（预留）
├── static/
│   ├── index.html         # 前端页面
│   ├── app.js             # 前端逻辑
│   └── styles.css         # 样式文件
├── samples/
│   ├── sample_10_people.xlsx   # 示例数据
│   └── multi_sheet_test.xlsx   # 多 sheet 示例
├── scripts/
│   └── generate_sample_excel.py  # 生成示例脚本
├── pg/
│   └── connect.py         # PostgreSQL 连接测试
├── .env                   # 环境变量配置
├── USAGE.md               # 详细使用指南
└── README.md              # 本文件
```

## 工作量评分模型

综合评分基于以下维度：

| 指标类型 | 指标 | 权重方向 |
|----------|------|----------|
| 负荷指标 | oncall未闭环、待处理工单、昨日新增问题 | 正相关 |
| 问题类型 | 管控、内核、咨询 | 正相关（内核权重更高） |
| 负向指标 | 透传求助 | 负相关 |
| 鼓励产出 | 问题单、需求单、wiki、分析报告 | 正相关 |

**风险分层**：基于透传求助 + 待处理工单 + oncall未闭环计算风险分，分为 `high / medium / low` 三档。

## PostgreSQL 配置（可选）

如需启用自定义模式入库功能：

```bash
# .env 文件
PG_DSN=postgresql://user:password@127.0.0.1:5432/j1_analytics
```

初始化数据库：
```bash
python db/init_db.py
```

## 详细文档

完整安装、配置、API 说明请参阅 [USAGE.md](./USAGE.md)