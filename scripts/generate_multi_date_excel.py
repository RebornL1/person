"""
生成多日期测试Excel文件
每个sheet代表一天的工作量数据，用于测试多日汇总功能
"""

import pandas as pd
import random
from datetime import datetime, timedelta
from pathlib import Path

# 配置
OUTPUT_PATH = Path(__file__).parent.parent / "samples" / "multi_date_workload.xlsx"
PEOPLE_NAMES = ["张伟", "李娜", "王强", "赵敏", "陈晨", "刘洋", "周杰", "吴婷", "郑凯", "孙悦", "杨帆", "黄丽"]

# 工作量列定义
COLUMNS = [
    "姓名",
    "oncall接单未闭环的数量",
    "名下的待处理工单数",
    "昨日新增多少个问题",
    "多少个管控的问题",
    "多少个内核的问题",
    "多少个咨询问题",
    "透传求助了多少个",
    "问题单数量",
    "需求单数量",
    "wiki输出数量",
    "问题分析报告数量"
]


def generate_daily_data(date_str: str, people: list[str]) -> pd.DataFrame:
    """生成一天的工作量数据"""
    rows = []
    for person in people:
        # 基础随机数据，但保持一定的规律性
        base_value = random.randint(0, 5)
        
        row = {
            "姓名": person,
            "oncall接单未闭环的数量": random.randint(0, 15),
            "名下的待处理工单数": random.randint(2, 20),
            "昨日新增多少个问题": random.randint(0, 8),
            "多少个管控的问题": random.randint(0, 4),
            "多少个内核的问题": random.randint(0, 5),
            "多少个咨询问题": random.randint(0, 6),
            "透传求助了多少个": random.randint(0, 3),
            "问题单数量": random.randint(0, 8),
            "需求单数量": random.randint(0, 5),
            "wiki输出数量": random.randint(0, 4),
            "问题分析报告数量": random.randint(0, 3)
        }
        rows.append(row)
    
    df = pd.DataFrame(rows)
    return df


def main():
    """生成多日期Excel文件"""
    # 创建输出目录
    OUTPUT_PATH.parent.mkdir(parents=True, exist_ok=True)
    
    # 生成3天的数据
    base_date = datetime(2026, 4, 21)
    dates = [
        (base_date - timedelta(days=1)).strftime("%Y-%m-%d"),  # 4-20
        base_date.strftime("%Y-%m-%d"),  # 4-21
        (base_date + timedelta(days=1)).strftime("%Y-%m-%d"),  # 4-22
    ]
    
    # 每天的人员可能略有不同，模拟实际场景
    all_people = PEOPLE_NAMES.copy()
    
    # 创建ExcelWriter
    with pd.ExcelWriter(OUTPUT_PATH, engine="openpyxl") as writer:
        for date_str in dates:
            # 每天随机选择10-12人（模拟请假等情况）
            daily_people = random.sample(all_people, random.randint(10, 12))
            df = generate_daily_data(date_str, sorted(daily_people))
            
            # Sheet名称使用日期格式（如 "4月20日"）
            month = int(date_str.split("-")[1])
            day = int(date_str.split("-")[2])
            sheet_name = f"{month}月{day}日"
            df.to_excel(writer, sheet_name=sheet_name, index=False)
            
            print(f"生成sheet: {sheet_name}, 人员数: {len(daily_people)}")
    
    print(f"\n文件已生成: {OUTPUT_PATH}")
    print(f"包含 {len(dates)} 个日期的数据，可用于测试多日汇总功能")


if __name__ == "__main__":
    main()