import argparse
import asyncio
import json
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

from playwright.async_api import async_playwright


def load_grades_json(path: Path) -> List[Dict[str, Any]]:
    payload = json.loads(path.read_text(encoding="utf-8"))
    if isinstance(payload, list):
        return payload
    if isinstance(payload, dict) and "grades" in payload and isinstance(payload["grades"], list):
        return payload["grades"]
    raise ValueError("grades.json 格式不支持：需要是 list 或包含 grades 字段的 dict。")


def build_grade_map(grades: List[Dict[str, Any]]) -> Dict[str, Dict[str, Any]]:
    mp: Dict[str, Dict[str, Any]] = {}
    for g in grades:
        name = str(g.get("name", "")).strip()
        if not name:
            continue
        mp[name] = g
    return mp


def unique_value(grades: List[Dict[str, Any]], key: str) -> Optional[str]:
    vals = {str(g.get(key)).strip() for g in grades if g.get(key) not in (None, "", "nan")}
    if len(vals) == 1:
        return next(iter(vals))
    return None


async def get_state(page) -> Dict[str, Any]:
    return await page.evaluate("window.__AUTO_GRADE_ENTRY__.getState()")


async def get_visible_rows(page) -> List[Dict[str, str]]:
    # 返回当前页可见学生：[{id,name}]
    return await page.evaluate(
        """
() => {
  const st = window.__AUTO_GRADE_ENTRY__.getState();
  const start = (st.pageIndex - 1) * st.pageSize;
  const end = start + st.pageSize;
  return st.rows.slice(start, end).map(r => ({ id: r.id, name: r.name }));
}
"""
    )


async def set_scores(page, row_id: str, usual: Any, exam: Any) -> bool:
    return await page.evaluate(
        "(args) => window.__AUTO_GRADE_ENTRY__.setRowScores(args.rowId, args.usual, args.exam)",
        {"rowId": row_id, "usual": usual, "exam": exam},
    )


async def go_to_page(page, p: int) -> None:
    await page.evaluate("(p) => window.__AUTO_GRADE_ENTRY__.goToPage(p)", p)
    await page.wait_for_function("(p) => window.__AUTO_GRADE_ENTRY__.getState().pageIndex === p", p)


async def submit_page(page) -> None:
    # 用 hook 提交，不依赖按钮是否可见
    await page.evaluate("() => window.__AUTO_GRADE_ENTRY__.submitPage()")


async def run(url: str, grades_path: Path, page_size: int, headless: bool) -> int:
    grades = load_grades_json(grades_path)
    grade_map = build_grade_map(grades)
    if not grade_map:
        raise ValueError("grades.json 中没有有效的 name 记录。")

    class_unique = unique_value(grades, "class_name")
    course_unique = unique_value(grades, "course")

    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=headless)
        page = await browser.new_page()

        await page.goto(url, wait_until="domcontentloaded")
        await page.wait_for_function("() => !!window.__AUTO_GRADE_ENTRY__")

        # 统一设置：清空搜索、设置每页条数
        await page.fill("#searchInput", "")
        await page.select_option("#pageSizeSelect", str(page_size))

        # 如果 Excel 里的班级/课程都是同一个，则直接设置（该网页是“全班/全课程”录入模型）
        if class_unique:
            try:
                await page.select_option("#classSelect", class_unique)
            except Exception:
                pass
        if course_unique:
            try:
                await page.select_option("#courseSelect", course_unique)
            except Exception:
                pass

        st = await get_state(page)
        total = len(st.get("rows", []))
        page_size_effective = int(st.get("pageSize", page_size))
        total_pages = max(1, (total + page_size_effective - 1) // page_size_effective)

        filled = 0
        missing: List[str] = []

        for pi in range(1, total_pages + 1):
            await go_to_page(page, pi)
            visible = await get_visible_rows(page)

            # 整页批量写入（循环调用 hook，成本极低）
            for r in visible:
                name = r["name"]
                rid = r["id"]
                g = grade_map.get(name)
                if not g:
                    missing.append(name)
                    continue
                ok = await set_scores(page, rid, g.get("usual"), g.get("exam"))
                if ok:
                    filled += 1

            # 一次提交
            await submit_page(page)

        # 校验：是否还有 dirty
        st2 = await get_state(page)
        dirty_left = [r["name"] for r in st2.get("rows", []) if r.get("dirty")]

        await browser.close()

    print(f"完成：已填 {filled} 人（按姓名匹配）")
    if missing:
        uniq = sorted(set(missing))
        print(f"未匹配到成绩（网页名单里有，但 grades.json 没有）：{uniq[:10]}{'...' if len(uniq) > 10 else ''}")
    if dirty_left:
        print(f"仍有未提交修改（请检查）：{dirty_left[:10]}{'...' if len(dirty_left) > 10 else ''}")
        return 2
    return 0


def main() -> int:
    ap = argparse.ArgumentParser(description="优化版：Playwright 整页批量输入 + 分页循环 + 一次提交")
    ap.add_argument("--url", required=True, help="成绩录入网页 URL，例如 http://localhost:5173")
    ap.add_argument("--grades", required=True, help="grades.json 路径（extract_excel.py 输出）")
    ap.add_argument("--page-size", type=int, default=10, help="每页条数（需与网页选项一致）")
    ap.add_argument("--headless", action="store_true", help="无头模式运行（默认有头，方便观察）")
    args = ap.parse_args()

    return asyncio.run(run(args.url, Path(args.grades).resolve(), args.page_size, args.headless))


if __name__ == "__main__":
    raise SystemExit(main())

