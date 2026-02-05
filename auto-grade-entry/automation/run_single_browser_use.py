import argparse
import asyncio
import json
import os
from pathlib import Path
from typing import Any, Dict, List, Optional

from dotenv import load_dotenv

from browser_use import Agent
from browser_use.llm import ChatDeepSeek


def load_grades_for_name(grades_path: Path, name: str) -> Optional[Dict[str, Any]]:
    payload = json.loads(grades_path.read_text(encoding="utf-8"))
    grades: List[Dict[str, Any]]
    if isinstance(payload, list):
        grades = payload
    elif isinstance(payload, dict) and isinstance(payload.get("grades"), list):
        grades = payload["grades"]
    else:
        return None

    for g in grades:
        if str(g.get("name", "")).strip() == name:
            return g
    return None


def build_task(url: str, name: str, usual: int, exam: int) -> str:
    return f"""
你在一个本地网页里录入学生成绩。请严格按步骤操作，尽量少走弯路：

1) 打开网址：{url}
2) 在“搜索姓名”的输入框（id 是 #searchInput）输入：{name}，让表格只显示该学生。
3) 在该学生这一行，依次完成：
   - 点击“平时成绩”单元格 → 弹窗输入框出现 → 输入 {usual} → 按回车保存
   - 点击“考试成绩”单元格 → 输入 {exam} → 按回车保存
4) 检查“最终成绩”是否等于 平时+考试 = {usual + exam}
5) 点击右下角按钮“提交本页”
6) 最后确认该学生行的“状态”为“已提交”。

注意：
- 每次输入成绩都要按回车保存（ESC 是取消）
- 如果看不到“提交本页”按钮，说明本页没有未提交修改；请确认成绩确实被改动后再提交
"""


async def main_async(url: str, name: str, usual: int, exam: int) -> None:
    load_dotenv()
    api_key = os.getenv("DEEPSEEK_API_KEY")
    if not api_key:
        raise RuntimeError("未设置 DEEPSEEK_API_KEY。请在 .env 中配置，或设置环境变量。")

    base_url = os.getenv("DEEPSEEK_BASE_URL", "https://api.deepseek.com/v1")

    llm = ChatDeepSeek(
        base_url=base_url,
        model="deepseek-chat",
        api_key=api_key,
    )

    agent = Agent(
        task=build_task(url=url, name=name, usual=usual, exam=exam),
        llm=llm,
        use_vision=False,
        max_history_items=8,  # 重点：限制历史上下文，减少 token
        flash_mode=True,  # 重点：跳过冗余思考/评估，进一步降 token
        max_actions_per_step=6,
    )

    await agent.run(max_steps=25)


def main() -> int:
    ap = argparse.ArgumentParser(description="初版：browser-use + DeepSeek 跑通单个学生录入+提交")
    ap.add_argument("--url", required=True, help="成绩录入网页 URL，例如 http://localhost:5173")
    ap.add_argument("--name", required=True, help="学生姓名（与网页名单匹配）")
    ap.add_argument("--usual", type=int, default=None, help="平时成绩（0-100）。不填则尝试从 --grades 里按姓名读取")
    ap.add_argument("--exam", type=int, default=None, help="考试成绩（0-100）。不填则尝试从 --grades 里按姓名读取")
    ap.add_argument("--grades", default=None, help="grades.json 路径（extract_excel.py 输出），用于自动取数")
    args = ap.parse_args()

    usual = args.usual
    exam = args.exam
    if (usual is None or exam is None) and args.grades:
        g = load_grades_for_name(Path(args.grades).resolve(), args.name)
        if g:
            usual = usual if usual is not None else g.get("usual")
            exam = exam if exam is not None else g.get("exam")

    if usual is None:
        usual = 60
    if exam is None:
        exam = 40

    asyncio.run(main_async(args.url, args.name, int(usual), int(exam)))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

