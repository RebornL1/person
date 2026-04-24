"""
生成新格式的工作量Excel测试文件
包含完整的工作量分析字段和文本记录字段
"""

import pandas as pd
import random
from datetime import datetime, timedelta
from pathlib import Path

# 配置
OUTPUT_PATH = Path(__file__).parent.parent / "samples" / "new_workload.xlsx"
PEOPLE_NAMES = ["张伟", "李娜", "王强", "赵敏", "陈晨", "刘洋", "周杰", "吴婷", "郑凯", "孙悦", "杨帆", "黄丽"]

# 新的列定义（按照用户需求）
COLUMNS = [
    "接单数工作量",
    "序号",
    "名称",
    "参会情况",
    "日期",
    "未闭环oncall接单数",
    "名下待处理工单数",
    "昨日新增工单数",
    "咨询问题",
    "内核技术支持",
    "管控技术支持",
    "透传问题数量",
    "不规范走单数量",
    "昨日提单/需求数量",
    "昨日案例总结数量",
    "案例/提单记录",
    "昨日公共事务",
    "昨日其他说明",
    "今日风险问题/交接",
]

# 参会情况选项
ATTENDANCE_OPTIONS = ["已参会", "迟到", "调休", "请假一周", "请假几周", "未参会"]

# 公共事务选项
PUBLIC_AFFAIRS_OPTIONS = [
    "协助排查问题",
    "文档整理",
    "流程优化讨论",
    "知识分享",
    "值班",
    "",
]

# 其他说明选项
OTHER_NOTES_OPTIONS = [
    "",
    "调休一天",
    "请假处理私事",
    "加班处理紧急问题",
    "",
    "",
]

# 风险问题选项
RISK_OPTIONS = [
    "",
    "有未闭环的内核问题需跟进",
    "需交接管控问题给同事",
    "待处理工单积压，需关注",
    "透传问题较多，需加强学习",
    "",
]


def generate_daily_data(date_str: str, people: list[str], day_offset: int) -> pd.DataFrame:
    """生成一天的工作量数据"""
    rows = []
    
    for idx, person in enumerate(people, 1):
        # 基础随机数据，累加值随天数增加
        base_oncall = random.randint(2, 15) + day_offset * random.randint(0, 3)
        base_tickets = random.randint(3, 20) + day_offset * random.randint(0, 2)
        
        # 参会情况（大部分已参会，偶尔有其他情况）
        attendance = random.choices(
            ATTENDANCE_OPTIONS, 
            weights=[70, 10, 10, 3, 3, 4],
            k=1
        )[0]
        
        # 其他说明（与参会情况关联）
        other_notes = ""
        if attendance == "调休":
            other_notes = "调休一天"
        elif attendance == "请假一周":
            other_notes = "请假处理私事"
        elif attendance == "请假几周":
            other_notes = "请假处理私事，预计下周返回"
        
        # 风险问题（偶尔有风险）
        risk = random.choices(RISK_OPTIONS, weights=[50, 10, 10, 10, 5, 15], k=1)[0]
        
        # 公共事务（随机）
        public_affairs = random.choice(PUBLIC_AFFAIRS_OPTIONS)
        
        # 案例/提单记录（偶尔有）
        case_records = ""
        if random.random() > 0.7:
            case_records = f"https://wiki.example.com/{person}/{date_str}/case-{random.randint(1, 5)}"
        
        row = {
            "接单数工作量": random.randint(5, 25),
            "序号": idx,
            "名称": person,
            "参会情况": attendance,
            "日期": date_str,
            "未闭环oncall接单数": base_oncall,
            "名下待处理工单数": base_tickets,
            "昨日新增工单数": random.randint(0, 8),  # 当日工作量
            "咨询问题": random.randint(0, 6),
            "内核技术支持": random.randint(0, 5),
            "管控技术支持": random.randint(0, 4),
            "透传问题数量": random.randint(0, 3),  # 负向
            "不规范走单数量": random.randint(0, 2),  # 负向
            "昨日提单/需求数量": random.randint(0, 3),  # 正向
            "昨日案例总结数量": random.randint(0, 2),  # 正向
            "案例/提单记录": case_records,
            "昨日公共事务": public_affairs,
            "昨日其他说明": other_notes,
            "今日风险问题/交接": risk,
        }
        rows.append(row)
    
    df = pd.DataFrame(rows)
    return df


def main():
    """生成多日期Excel文件"""
    # 创建输出目录
    OUTPUT_PATH.parent.mkdir(parents=True, exist_ok=True)
    
    # 生成3天的数据：4月21日、4月22日、4月23日
    dates = [
        datetime(2026, 4, 21).strftime("%Y-%m-%d"),
        datetime(2026, 4, 22).strftime("%Y-%m-%d"),
        datetime(2026, 4, 23).strftime("%Y-%m-%d"),
    ]
    
    # 创建ExcelWriter
    with pd.ExcelWriter(OUTPUT_PATH, engine="openpyxl") as writer:
        for day_offset, date_str in enumerate(dates):
            # 每天随机选择人员（模拟请假等情况）
            daily_people = random.sample(PEOPLE_NAMES, random.randint(10, 12))
            df = generate_daily_data(date_str, sorted(daily_people), day_offset)
            
            # Sheet名称使用日期格式
            month = int(date_str.split("-")[1])
            day = int(date_str.split("-")[2])
            sheet_name = f"{month}月{day}日"
            df.to_excel(writer, sheet_name=sheet_name, index=False)
            
            print(f"生成sheet: {sheet_name}, 人员数: {len(daily_people)}")
    
    print(f"\n文件已生成: {OUTPUT_PATH}")
    print(f"包含 {len(dates)} 个日期的数据")
    print(f"\n列结构:")
    for i, col in enumerate(COLUMNS, 1):
        print(f"  {i}. {col}")


if __name__ == "__main__":
    main()