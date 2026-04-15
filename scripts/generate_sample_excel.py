"""生成 8 人工作量分析示例 Excel，供上传测试。"""
from pathlib import Path

import pandas as pd

ROWS = [
    ("张伟", 12, 9, 6, 2, 1, 3, 6, 2, 4, 3, 2, 1),
    ("李娜", 8, 6, 4, 1, 2, 4, 7, 1, 2, 5, 3, 2),
    ("王强", 15, 11, 7, 3, 1, 2, 6, 1, 6, 2, 1, 1),
    ("刘洋", 7, 5, 3, 1, 0, 5, 6, 2, 2, 2, 4, 1),
    ("陈静", 10, 8, 5, 2, 1, 3, 6, 2, 3, 3, 2, 2),
    ("杨磊", 13, 10, 8, 2, 2, 3, 7, 1, 5, 4, 1, 1),
    ("赵敏", 9, 7, 6, 1, 1, 4, 6, 2, 3, 2, 3, 2),
    ("周杰", 11, 9, 5, 2, 3, 2, 7, 1, 4, 2, 2, 3),
]

COLUMNS = [
    "姓名",
    "oncall接单未闭环的数量",
    "名下的待处理工单数",
    "昨日新增多少个问题",
    "多少个管控的问题",
    "多少个内核的问题",
    "多少个咨询问题",
    "每日的问题数量",
    "透传求助了多少个",
    "问题单数量",
    "需求单数量",
    "wiki输出数量",
    "问题分析报告数量",
]


def main() -> None:
    root = Path(__file__).resolve().parents[1]
    out = root / "samples" / "sample_8_people.xlsx"
    out.parent.mkdir(parents=True, exist_ok=True)
    df = pd.DataFrame(ROWS, columns=COLUMNS)
    df.to_excel(out, index=False, engine="openpyxl")
    print(out)


if __name__ == "__main__":
    main()
