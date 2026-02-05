"""生成示例 Excel，用于测试填表程序。支持单表与双列排列两种格式。"""
from pathlib import Path

import pandas as pd


def main() -> None:
    base = Path(__file__).parent

    # 1) 单表示例
    out_single = base / "sample_data.xlsx"
    df = pd.DataFrame([
        {"姓名": "张三", "学号": "2024001", "成绩": "85", "备注": "优秀"},
        {"姓名": "李四", "学号": "2024002", "成绩": "78", "备注": "良好"},
        {"姓名": "王五", "学号": "2024003", "成绩": "92", "备注": "优秀"},
    ])
    df.to_excel(out_single, index=False, engine="openpyxl")
    print(f"已生成单表示例：{out_single}")

    # 2) 双列排列示例（中间空列分隔）：用 openpyxl 直接写，保证表头为 姓名,学号,成绩,空,姓名,学号,成绩
    out_double = base / "sample_data_double_column.xlsx"
    from openpyxl import Workbook
    wb = Workbook()
    ws = wb.active
    header = ["姓名", "学号", "成绩", "", "姓名", "学号", "成绩"]
    ws.append(header)
    ws.append(["张三", "2024001", "85", "", "李四", "2024002", "78"])
    ws.append(["王五", "2024003", "92", "", "赵六", "2024004", "88"])
    wb.save(out_double)
    print(f"已生成双列示例（空列分隔）：{out_double}")

    # 3) 平时成绩 + 考试成绩 两列示例（分别同步进系统）
    out_two_scores = base / "sample_data_usual_exam.xlsx"
    df_two = pd.DataFrame([
        {"姓名": "张三", "学号": "2024001", "平时成绩": 88, "考试成绩": 82, "备注": "优秀"},
        {"姓名": "李四", "学号": "2024002", "平时成绩": 75, "考试成绩": 80, "备注": "良好"},
        {"姓名": "王五", "学号": "2024003", "平时成绩": 90, "考试成绩": 94, "备注": "优秀"},
    ])
    df_two.to_excel(out_two_scores, index=False, engine="openpyxl")
    print(f"已生成平时成绩+考试成绩双列示例：{out_two_scores}")


if __name__ == "__main__":
    main()
