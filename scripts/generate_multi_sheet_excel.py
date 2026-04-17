#!/usr/bin/env python3
"""
生成包含多个sheet页的测试Excel文件，用于测试数据导入配置功能。
"""

import pandas as pd
from pathlib import Path


def generate_multi_sheet_excel():
    """生成包含3个sheet页的测试Excel文件。"""
    
    # Sheet1: 团队A工作量数据
    team_a_data = {
        "姓名": ["张三", "李四", "王五", "赵六", "陈七", "周八", "吴九", "郑十"],
        "oncall接单未闭环的数量": [3, 2, 5, 1, 4, 2, 3, 1],
        "名下的待处理工单数": [5, 3, 8, 2, 6, 4, 5, 2],
        "昨日新增多少个问题": [2, 1, 3, 0, 2, 1, 2, 0],
        "多少个管控的问题": [1, 2, 1, 0, 2, 1, 1, 0],
        "多少个内核的问题": [2, 1, 3, 1, 2, 1, 2, 1],
        "多少个咨询问题": [1, 2, 0, 3, 1, 2, 1, 2],
        "透传求助了多少个": [1, 2, 4, 0, 3, 1, 2, 0],
        "问题单数量": [2, 1, 0, 3, 1, 2, 1, 3],
        "需求单数量": [1, 2, 1, 2, 0, 1, 2, 1],
        "wiki输出数量": [3, 2, 1, 5, 2, 3, 2, 4],
        "问题分析报告数量": [1, 0, 0, 2, 1, 1, 0, 2],
    }
    df_team_a = pd.DataFrame(team_a_data)
    
    # Sheet2: 团队B工作量数据
    team_b_data = {
        "姓名": ["钱一", "孙二", "周三", "李四", "王五", "赵六", "冯七", "蒋八"],
        "oncall接单未闭环的数量": [4, 1, 6, 2, 3, 5, 1, 2],
        "名下的待处理工单数": [7, 2, 10, 3, 5, 8, 2, 4],
        "昨日新增多少个问题": [3, 0, 5, 1, 2, 4, 0, 1],
        "多少个管控的问题": [2, 0, 2, 1, 1, 2, 0, 1],
        "多少个内核的问题": [3, 1, 4, 1, 2, 3, 1, 1],
        "多少个咨询问题": [2, 3, 1, 4, 2, 3, 3, 2],
        "透传求助了多少个": [2, 0, 5, 1, 2, 4, 0, 1],
        "问题单数量": [1, 3, 0, 2, 1, 0, 3, 2],
        "需求单数量": [2, 1, 0, 3, 2, 1, 1, 2],
        "wiki输出数量": [2, 4, 1, 3, 2, 1, 4, 3],
        "问题分析报告数量": [0, 2, 0, 1, 1, 0, 2, 1],
    }
    df_team_b = pd.DataFrame(team_b_data)
    
    # Sheet3: 产品线X数据（不同列名格式）
    product_x_data = {
        "员工姓名": ["小明", "小红", "小刚", "小芳", "小强"],
        "未闭环oncall": [2, 1, 4, 1, 3],
        "待处理工单": [4, 2, 7, 2, 5],
        "新增问题": [1, 0, 3, 0, 2],
        "管控类": [1, 0, 2, 0, 1],
        "内核类": [1, 1, 3, 1, 2],
        "咨询类": [2, 3, 1, 4, 2],
        "求助透传": [0, 0, 3, 0, 1],
        "提问题单": [3, 4, 0, 5, 2],
        "提需求单": [2, 3, 1, 4, 1],
        "写wiki": [4, 5, 2, 6, 3],
        "写分析报告": [2, 3, 0, 4, 1],
    }
    df_product_x = pd.DataFrame(product_x_data)
    
    # Sheet4: 普通数据表（非工作量分析格式）
    normal_data = {
        "项目名称": ["项目A", "项目B", "项目C", "项目D", "项目E"],
        "负责人": ["张三", "李四", "王五", "赵六", "陈七"],
        "进度(%)": [80, 60, 90, 40, 70],
        "预算(万元)": [100, 50, 200, 30, 80],
        "已用预算(万元)": [75, 35, 180, 15, 55],
        "状态": ["进行中", "进行中", "已完成", "待启动", "进行中"],
        "预计完成日期": ["2026-05-01", "2026-06-15", "2026-04-10", "2026-07-01", "2026-05-20"],
    }
    df_normal = pd.DataFrame(normal_data)
    
    # 创建输出目录
    output_dir = Path(__file__).resolve().parent.parent / "samples"
    output_dir.mkdir(parents=True, exist_ok=True)
    
    output_file = output_dir / "multi_sheet_test.xlsx"
    
    # 写入Excel文件，包含多个sheet
    with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
        df_team_a.to_excel(writer, sheet_name="团队A工作量", index=False)
        df_team_b.to_excel(writer, sheet_name="团队B工作量", index=False)
        df_product_x.to_excel(writer, sheet_name="产品线X数据", index=False)
        df_normal.to_excel(writer, sheet_name="项目进度表", index=False)
    
    print(f"已生成测试文件: {output_file}")
    print(f"包含4个Sheet页:")
    print(f"  1. 团队A工作量 - 8人数据（标准列名）")
    print(f"  2. 团队B工作量 - 8人数据（标准列名）")
    print(f"  3. 产品线X数据 - 5人数据（不同列名格式，测试列映射）")
    print(f"  4. 项目进度表 - 5项目数据（非工作量分析格式）")
    
    return output_file


if __name__ == "__main__":
    generate_multi_sheet_excel()