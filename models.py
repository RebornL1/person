"""
数据模型定义：请求/响应模型、默认配置
"""

from typing import Any
from pydantic import BaseModel, Field


# 新格式的默认权重配置（基于新的列结构）
DEFAULT_WEIGHTS = {
    # 正向指标（工作负荷）
    "oncall_open": 0.9,          # 未闭环oncall接单数（累加，反映持续压力）
    "pending_ticket": 0.8,       # 名下待处理工单数（累加，反映持续压力）
    "new_ticket_yesterday": 1.4, # 昨日新增工单数（当日，反映当日工作量）
    "consult_issue": 0.85,       # 咨询问题
    "kernel_support": 1.35,      # 内核技术支持（权重较高）
    "governance_support": 1.0,   # 管控技术支持
    
    # 负向指标（需要降低工作衡量）
    "escalation_help": -0.6,     # 透传问题数量（负向）
    "irregular_ticket": -0.3,    # 不规范走单数量（负向）
    
    # 正向产出指标
    "ticket_demand_output": 1.25,   # 昨日提单/需求数量
    "case_summary_output": 1.3,     # 昨日案例总结数量
}

# 新格式的默认列映射配置
DEFAULT_COLUMN_ALIASES = {
    # 基础字段
    "name": ["名称", "姓名", "名字", "人员", "同学", "name", "员工姓名"],
    "workload_count": ["接单数工作量", "工作量", "接单数"],
    
    # 累加指标（反映持续压力）
    "oncall_open": ["未闭环oncall接单数", "oncall未闭环", "接单未闭环", "oncall_open", "未闭环数量"],
    "pending_ticket": ["名下待处理工单数", "待处理工单", "待处理工单数", "pending_ticket", "名下工单"],
    
    # 当日指标（反映当日工作量）
    "new_ticket_yesterday": ["昨日新增工单数", "昨日新增问题", "昨日新增", "new_issue_yesterday"],
    "consult_issue": ["咨询问题", "多少个咨询问题", "咨询", "consult_issue", "咨询类问题"],
    "kernel_support": ["内核技术支持", "多少个内核的问题", "内核问题", "内核", "kernel_issue", "kernel_support"],
    "governance_support": ["管控技术支持", "多少个管控的问题", "管控问题", "管控", "governance_issue", "governance_support"],
    
    # 负向指标
    "escalation_help": ["透传问题数量", "透传求助了多少个", "透传求助", "透传", "escalation_help", "求助数量"],
    "irregular_ticket": ["不规范走单数量", "不规范走单", "不规范", "irregular_ticket"],
    
    # 正向产出
    "ticket_demand_output": ["昨日提单/需求数量", "提单数量", "需求数量", "问题单数量", "issue_ticket_output"],
    "case_summary_output": ["昨日案例总结数量", "案例总结", "wiki输出数量", "wiki_output", "analysis_report_output"],
    
    # 文本字段（特殊展示）
    "attendance": ["参会情况", "参会"],
    "date": ["日期", "date"],
    "case_record": ["案例/提单记录", "案例记录", "提单记录"],
    "public_affairs": ["昨日公共事务", "公共事务"],
    "other_notes": ["昨日其他说明", "其他说明"],
    "risk_handover": ["今日风险问题/交接", "风险问题", "交接"],
}

# 新格式的权重标签（前端显示用）
WEIGHT_LABELS = {
    "oncall_open": "未闭环oncall接单(累加)",
    "pending_ticket": "名下待处理工单(累加)",
    "new_ticket_yesterday": "昨日新增工单(当日)",
    "consult_issue": "咨询问题",
    "kernel_support": "内核技术支持",
    "governance_support": "管控技术支持",
    "escalation_help": "透传问题(负向)",
    "irregular_ticket": "不规范走单(负向)",
    "ticket_demand_output": "提单/需求产出(正向)",
    "case_summary_output": "案例总结产出(正向)",
}

# 文本字段的展示类型配置
TEXT_COLUMN_DISPLAY_CONFIG = {
    "attendance": {
        "display_name": "参会情况",
        "show_type": "badge",  # 使用标签形式展示
        "colors": {
            "已参会": "good",
            "迟到": "warn",
            "调休": "muted",
            "请假一周": "danger",
            "请假几周": "danger",
            "未参会": "danger",
        }
    },
    "case_record": {
        "display_name": "案例/提单记录",
        "show_type": "link",  # 链接形式展示
        "link_prefix": "",
    },
    "risk_handover": {
        "display_name": "今日风险问题/交接",
        "show_type": "alert",  # 警告形式单独展示
        "highlight": True,
    },
    "public_affairs": {
        "display_name": "昨日公共事务",
        "show_type": "text",
    },
    "other_notes": {
        "display_name": "昨日其他说明",
        "show_type": "text",
    },
}


class SaveCustomModeRequest(BaseModel):
    """保存自定义模式请求"""
    mode_name: str = Field(min_length=1, max_length=64)
    selected_columns: list[str] = Field(default_factory=list)
    rows: list[dict[str, Any]] = Field(default_factory=list)


class DeleteCustomModeRequest(BaseModel):
    """删除自定义模式请求"""
    mode_name: str = Field(min_length=1, max_length=64)


class SaveColumnMappingRequest(BaseModel):
    """保存列映射配置请求（新格式）"""
    mapping_name: str = Field(min_length=1, max_length=64)
    # 基础字段
    name_aliases: list[str] = Field(default_factory=list)
    workload_count_aliases: list[str] = Field(default_factory=list)
    # 累加指标
    oncall_open_aliases: list[str] = Field(default_factory=list)
    pending_ticket_aliases: list[str] = Field(default_factory=list)
    # 当日指标
    new_ticket_yesterday_aliases: list[str] = Field(default_factory=list)
    consult_issue_aliases: list[str] = Field(default_factory=list)
    kernel_support_aliases: list[str] = Field(default_factory=list)
    governance_support_aliases: list[str] = Field(default_factory=list)
    # 负向指标
    escalation_help_aliases: list[str] = Field(default_factory=list)
    irregular_ticket_aliases: list[str] = Field(default_factory=list)
    # 正向产出
    ticket_demand_output_aliases: list[str] = Field(default_factory=list)
    case_summary_output_aliases: list[str] = Field(default_factory=list)
    # 文本字段
    attendance_aliases: list[str] = Field(default_factory=list)
    date_aliases: list[str] = Field(default_factory=list)
    case_record_aliases: list[str] = Field(default_factory=list)
    public_affairs_aliases: list[str] = Field(default_factory=list)
    other_notes_aliases: list[str] = Field(default_factory=list)
    risk_handover_aliases: list[str] = Field(default_factory=list)


class DeleteColumnMappingRequest(BaseModel):
    """删除列映射配置请求"""
    mapping_id: int