"""
Phase 2：任务下发 + 强校验

- 从 Phase 1 拿到 (records, meta)。
- 强校验：未同时识别「平时成绩」与「考试成绩」列 → 停止自动填表，输出诊断。
- 通过后生成「姓名 | 平时成绩(目标值) | 考试成绩(目标值)」任务文本，供 Agent 执行。
"""
import argparse
import sys
from pathlib import Path

from excel_reader import (
    USUAL_SCORE_KEY,
    EXAM_SCORE_KEY,
    read_excel_to_records,
    records_to_task_text,
    validate_records_for_fill,
)


def build_task(records: list, url: str, page_size: int | None = None) -> str:
    """构造填表任务文案：比对优先，数值已是 int，字段与网页语义对应。"""
    total = len(records)
    page_size = page_size or total
    intro = (
        f"打开录入页：{url}\n\n"
        "将下方表格中每行的「平时成绩」「考试成绩」**分别**填入系统对应列。\n"
        "对每一行（按姓名匹配）：页面值为空 → 填目标值；页面值存在且整数一致 → 跳过；不一致 → 允许修改。\n"
        "本页内批量填写所有需要修改的行，只在「本页发生过修改」时点击一次【提交本页】。\n\n"
    )
    if page_size < total:
        intro += f"本批共 {total} 条，按每 {page_size} 条一页处理；翻页后继续当前游标。\n\n"
    intro += "若出现「预览/确认/二次确认」页面：不再比对成绩，只点击【确认/提交】，然后继续流程。\n\n"
    body = records_to_task_text(records)
    return intro + "数据：\n" + body


def main() -> None:
    parser = argparse.ArgumentParser(description="Excel → 强校验 → 任务生成（browser-use 自动成绩录入）")
    parser.add_argument("-e", "--excel", required=True, type=Path, help="Excel 文件路径")
    parser.add_argument("-u", "--url", required=True, help="录入页 URL")
    parser.add_argument("--sheet", default=None, help="工作表名或索引")
    parser.add_argument("--page-size", type=int, default=None, help="每批条数")
    parser.add_argument("--max-rows", type=int, default=None, help="最多处理行数")
    parser.add_argument("--dry-run", action="store_true", help="只做读取+校验+打印任务，不调 Agent")
    args = parser.parse_args()

    records, meta = read_excel_to_records(args.excel, sheet=args.sheet)
    if args.max_rows:
        records = records[: args.max_rows]

    ok, msg = validate_records_for_fill(records, meta)
    if not ok:
        print("【强校验未通过】禁止自动填表。", file=sys.stderr)
        print(msg, file=sys.stderr)
        sys.exit(1)

    print(f"已读取 {len(records)} 条记录，已同时识别「平时成绩」与「考试成绩」列，可安全填表。")
    task = build_task(records, args.url, page_size=args.page_size)
    if args.dry_run:
        print("\n--- 任务文案（dry-run）---\n")
        print(task)
        return
    print("\n--- 任务文案（前 600 字）---\n")
    print(task[:600] + "..." if len(task) > 600 else task)
    print("\n（实际调用 browser-use 时传入完整 task 与 url。）")


if __name__ == "__main__":
    main()
