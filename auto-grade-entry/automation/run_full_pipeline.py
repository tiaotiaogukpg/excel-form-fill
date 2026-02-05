import argparse
import asyncio
import json
from pathlib import Path

from extract_excel import read_excel_grades
from run_batch_playwright import run as run_batch


def write_grades_json(out_path: Path, grades_rows, meta) -> None:
    out_path.parent.mkdir(parents=True, exist_ok=True)
    payload = {
        "meta": meta,
        "grades": [g.__dict__ for g in grades_rows],
    }
    out_path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")


def main() -> int:
    ap = argparse.ArgumentParser(description="完整闭环：Excel → JSON → 浏览器批量分页录入 → 一次提交")
    ap.add_argument("--excel", required=True, help="Excel 路径，例如 data.xlsx")
    ap.add_argument("--url", required=True, help="成绩录入网页 URL，例如 http://localhost:5173")
    ap.add_argument("--sheet", default=None, help="工作表名称（不填则读取第一个）")
    ap.add_argument("--out", default="automation/grades.json", help="输出 JSON 路径（默认 automation/grades.json）")
    ap.add_argument("--default-class", default=None, help="当 Excel 没有班级列时使用")
    ap.add_argument("--default-course", default=None, help="当 Excel 没有课程列时使用")
    ap.add_argument("--page-size", type=int, default=10, help="每页条数（需与网页选项一致）")
    ap.add_argument("--headless", action="store_true", help="无头模式运行（默认有头，方便观察）")
    args = ap.parse_args()

    excel_path = Path(args.excel).expanduser().resolve()
    out_path = Path(args.out).expanduser().resolve()

    grades_rows, meta = read_excel_grades(
        excel_path=excel_path,
        sheet=args.sheet,
        default_class=args.default_class,
        default_course=args.default_course,
    )
    write_grades_json(out_path, grades_rows, meta)
    print(f"已生成：{out_path}")

    return asyncio.run(run_batch(args.url, out_path, args.page_size, args.headless))


if __name__ == "__main__":
    raise SystemExit(main())

