# J1 - Excel 工作量分析 Web 应用

一个基于 FastAPI 的 Excel 工作量分析工具，支持上传表格、自动评分、风险分层和数据入库。

## 功能特性

- **Excel 上传与预览** - 支持 `.xlsx` / `.xls` 格式，最多预览 200 行
- **工作量评分模型** - 综合评分、透传排序、风险分层、问题结构分析
- **GaussDB 现网聚焦指数** - 强调内核问题处置与知识沉淀能力
- **风险建议** - 自动生成运维改进建议
- **自动数据入库** - 上传时自动保存到 PostgreSQL，按日期整理历史数据
- **历史数据管理** - 支持加载最新数据、查看历史记录、删除历史会话
- **日期筛选功能** - 日期选择器 + 快捷按钮（近7天/30天/90天）筛选历史数据
- **自定义模式入库** - 保存分析结果到 PostgreSQL（可选）
- **权重配置弹窗** - 实时调整权重，模拟不同运维策略
- **入库状态提示** - 上传后显示数据入库状态（成功/失败/未配置）

## 最近更新

- 2026-04-17: 完善数据导入配置功能：添加列映射编辑、每个图表单独配置展示类型
- 2026-04-16: 添加数据导入配置功能：支持多sheet选择、列勾选、图表类型配置
- 2026-04-16: 核心指标侧边栏优化：宽度240px、默认打开、添加关闭按钮、内容左右布局
- 2026-04-16: 将核心指标总览改为左侧可折叠侧边栏，指标竖向陈列，支持点击展开/收起
- 2026-04-16: 修复上传功能 JavaScript 错误，添加入库状态显示
- 2026-04-16: 添加 favicon 图标，修复日期管理功能事件绑定
- 2026-04-16: 修复数据库表结构，添加 column_mapping_id 列

## 技术栈

| 类别 | 技术 |
|------|------|
| 后端 | FastAPI + Uvicorn |
| 数据处理 | Pandas + openpyxl + xlrd |
| 数据库 | PostgreSQL（可选） |
| 前端 | 原生 HTML/CSS/JS |

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
2. 查看工作量看板（综合评分、风险分层、透传排名等）
3. 点击「打开模型口径与权重配置」调整权重
4. 可选：保存自定义模式到 PostgreSQL

**示例数据**：`samples/sample_10_people.xlsx` 可用于测试。

## API 接口

|| 方法 | 路径 | 说明 |
||------|------|------|
|| GET | `/` | 主页面 |
|| POST | `/api/upload` | 上传 Excel，自动保存到数据库 |
|| GET | `/api/upload/history` | 获取历史上传记录（按日期分组） |
|| GET | `/api/upload/latest` | 获取最新一次上传的数据 |
|| GET | `/api/upload/session/{id}` | 获取指定会话的数据 |
|| POST | `/api/upload/delete/{id}` | 删除指定上传会话 |
|| GET | `/api/custom-mode/list` | 获取已保存模式 |
|| POST | `/api/custom-mode/save` | 保存自定义模式 |
|| POST | `/api/custom-mode/delete` | 删除模式 |

API 文档：http://127.0.0.1:8000/docs

## 项目结构

```
J1/
├── app.py              # FastAPI 应用入口
├── start.py            # 跨平台启动脚本
├── start.bat           # Windows 快捷启动
├── start.sh            # macOS/Linux 快捷启动
├── requirements.txt    # Python 依赖
├── static/
│   └── index.html      # 前端页面
├── samples/
│   └── sample_10_people.xlsx  # 示例数据
├── scripts/
│   └── generate_sample_excel.py  # 生成示例脚本
├── USAGE.md            # 详细使用指南
└── README.md           # 本文件
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

## 详细文档

完整安装、配置、API 说明请参阅 [USAGE.md](./USAGE.md)