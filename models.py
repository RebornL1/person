"""
数据模型定义：请求/响应模型、默认配置
"""

from typing import Any
from pydantic import BaseModel, Field


# 默认权重配置
DEFAULT_WEIGHTS = {
    "oncall_open": 0.9,
    "pending_ticket": 0.8,
    "new_issue_yesterday": 1.4,
    "governance_issue": 1.0,
    "kernel_issue": 1.35,
    "consult_issue": 0.85,
    "escalation_help": -0.6,
    "issue_ticket_output": 1.25,
    "requirement_ticket_output": 1.05,
    "wiki_output": 1.3,
    "analysis_report_output": 1.2,
}

# 默认列映射配置
DEFAULT_COLUMN_ALIASES = {
    "name": ["姓名", "名字", "人员", "同学", "name", "员工姓名"],
    "oncall_open": ["oncall接单未闭环的数量", "oncall未闭环", "接单未闭环", "oncall_open", "未闭环数量"],
    "pending_ticket": ["名下的待处理工单数", "待处理工单", "待处理工单数", "pending_ticket", "名下工单"],
    "new_issue_yesterday": ["昨日新增多少个问题", "昨日新增问题", "昨日新增", "new_issue_yesterday", "new_issues"],
    "governance_issue": ["多少个管控的问题", "管控问题", "管控", "governance_issue", "管控类问题"],
    "kernel_issue": ["多少个内核的问题", "内核问题", "内核", "kernel_issue", "内核类问题"],
    "consult_issue": ["多少个咨询问题", "咨询问题", "咨询", "consult_issue", "咨询类问题"],
    "escalation_help": ["透传求助了多少个", "透传求助", "透传", "escalation_help", "求助数量"],
    "issue_ticket_output": ["问题单数量", "提了多少问题单", "问题单", "issue_ticket_output", "问题单产出"],
    "requirement_ticket_output": ["需求单数量", "提了多少需求单", "需求单", "requirement_ticket_output", "需求单产出"],
    "wiki_output": ["wiki输出数量", "输出多少wiki", "wiki", "wiki_output", "wiki产出"],
    "analysis_report_output": ["问题分析报告数量", "输出多少问题分析报告", "分析报告", "analysis_report_output", "报告产出"],
}

# 权重标签（前端显示用）
WEIGHT_LABELS = {
    "oncall_open": "oncall未闭环",
    "pending_ticket": "待处理工单",
    "new_issue_yesterday": "昨日新增问题",
    "governance_issue": "管控问题",
    "kernel_issue": "内核问题",
    "consult_issue": "咨询问题",
    "escalation_help": "透传求助(负向建议为负数)",
    "issue_ticket_output": "问题单产出",
    "requirement_ticket_output": "需求单产出",
    "wiki_output": "wiki产出",
    "analysis_report_output": "分析报告产出",
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
    """保存列映射配置请求"""
    mapping_name: str = Field(min_length=1, max_length=64)
    name_aliases: list[str] = Field(default_factory=list)
    oncall_open_aliases: list[str] = Field(default_factory=list)
    pending_ticket_aliases: list[str] = Field(default_factory=list)
    new_issue_yesterday_aliases: list[str] = Field(default_factory=list)
    governance_issue_aliases: list[str] = Field(default_factory=list)
    kernel_issue_aliases: list[str] = Field(default_factory=list)
    consult_issue_aliases: list[str] = Field(default_factory=list)
    escalation_help_aliases: list[str] = Field(default_factory=list)
    issue_ticket_output_aliases: list[str] = Field(default_factory=list)
    requirement_ticket_output_aliases: list[str] = Field(default_factory=list)
    wiki_output_aliases: list[str] = Field(default_factory=list)
    analysis_report_output_aliases: list[str] = Field(default_factory=list)


class DeleteColumnMappingRequest(BaseModel):
    """删除列映射配置请求"""
    mapping_id: int