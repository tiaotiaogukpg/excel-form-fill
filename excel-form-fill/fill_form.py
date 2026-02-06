"""
Phase 2：任务下发 + 强校验 + browser-use 执行

- 从 Phase 1 拿到 (records, meta)，强校验通过后生成任务文本。
- 使用 DeepSeek 模型 + browser-use Agent 在浏览器中打开录入页并按要求填写成绩。
"""
import argparse
import asyncio
import json
import os
import sys
import time
from pathlib import Path

try:
    from dotenv import load_dotenv
    load_dotenv()
except ImportError:
    pass  # 未安装 python-dotenv 时跳过，依赖已设置的环境变量

try:
    from excel_reader import (
        read_excel_to_records,
        records_to_task_text,
        validate_records_for_fill,
    )
except ModuleNotFoundError as e:
    if "pandas" in str(e).lower() or (getattr(e, "name", None) == "pandas"):
        print("未找到 pandas，请使用本项目虚拟环境并安装依赖：", file=sys.stderr)
        print("  .venv\\Scripts\\python.exe -m pip install -r requirements.txt", file=sys.stderr)
        print("  然后运行：.venv\\Scripts\\python.exe fill_form.py ...", file=sys.stderr)
    raise


def build_task(
    records: list,
    url: str,
    page_size: int | None = None,
    excel_path: Path | None = None,
) -> str:
    """构造填表任务文案：比对优先，数值已是 int，字段与网页语义对应。"""
    total = len(records)
    page_size = page_size or total
    login_user = os.getenv("GRADE_ENTRY_USER", "").strip()
    login_password = os.getenv("GRADE_ENTRY_PASSWORD", "").strip()
    login_hint = ""
    if login_user and login_password:
        login_hint = f"若出现登录页，使用用户名 **{login_user}**、密码 **{login_password}** 登录。\n\n"
    excel_hint = ""
    if excel_path is not None:
        path_abs = excel_path.resolve()
        if path_abs.exists():
            excel_hint = (
                f"若页面有「导入 Excel」「导入表格」「文件上传」等入口，**优先上传本机 Excel 文件**（用 upload_file），路径：**{path_abs}**\n"
                "上传后若有需要再逐行核对或补填。\n\n"
            )
    # 强制第一步为导航，避免模型只“思考”不输出 navigate 动作
    intro = (
        f"第一步（必须）：立即访问成绩录入系统首页。\n"
        f"URL: {url}\n"
        "在页面加载完成前不要执行其他操作。\n\n"
        f"{login_hint}"
        f"**本批共 {total} 条**（下表即这 {total} 条），只核对/填写系统中已存在的、能与下表姓名匹配到的行；人数以本表为准，不要根据系统人数推断「缺失」或「多出」。\n"
        "**不要点击「加载示例」「加载示例数据」等按钮**：直接使用当前页面进行成绩录入。\n\n"
        "**本系统支持「先上传文件」**：若页面上有 Excel/CSV 导入或文件上传入口，**先直接上传文件**，无需先选择班级、课程；班级和课程会在导入后由系统填充或再选。\n\n"
        "**文件兼容性（必须先探测页面能力）**：\n"
        "1. 检测页面是否存在文件上传入口。\n"
        "2. 判断页面提示 / 上传控件的 accept 属性 / 文案中支持的文件类型（xlsx 优先于 csv）。\n"
        "3. 若支持 xlsx → 优先使用已提供的 xlsx 文件上传（见下方路径）；若仅支持 csv → 再考虑生成并上传 csv。\n"
        "4. 若无文件导入入口 → 才进入手动逐条填写流程。\n"
        "禁止在未检测页面支持格式前，强制固定使用某一种文件格式。\n\n"
        f"{excel_hint}"
        "若页面有「导入表格」「导入 Excel」「批量导入」等入口且上方已给出 Excel 路径，优先上传该 xlsx；否则将下方数据表导入或逐行填写。\n"
        "**手动填写时**：按姓名匹配到行，读取「平时成绩」「考试成绩」当前值并与下表比对；空→填目标值，一致→跳过，不一致→清空后填入目标值；本页有修改则点击一次【提交本页】。\n\n"
        "**表格遍历优先级（必须遵守，数据由分页/状态控制，非“滚动即加载”）**：\n"
        "1. 若页面存在分页控件（页码/下一页/第X页共Y页）→ **使用分页控件遍历所有页**，禁止仅依赖 scroll。\n"
        "2. 若存在「每页显示数量」选项（10/20/全部）→ 可调整页大小，但仍需确认是否存在分页或虚拟渲染，不得假定「全部」即 DOM 含所有行。\n"
        f"3. 若页面明确显示「共 X 条」「已导入 X 条」「导入完成…已填入 X 人」→ 若 X == {total}（目标条数），**以该提示作为成功判定**，无需逐条滚动核对。\n"
        "4. **仅当**以下三者同时满足时才允许使用 scroll：页面无分页控件、无总数提示、且数据明显随 scroll 增量加载。\n\n"
        "**核对完成的统一判定标准**：\n"
        "禁止以「我已经 scroll 过」「我大概看到了很多行」作为完成依据。\n"
        "正确完成判定（**任一满足**即可）：\n"
        f"- 系统提示：成功导入 X 条记录，且 X == {total}（目标学生数量）；或\n"
        "- 系统提示：全部提交成功 / 无错误；或\n"
        "- 已遍历所有分页（第 1 页…第 N 页），且每页数据均存在。\n"
        "未满足上述任一条件 → **不允许**标记任务成功（不得调用 done）。\n"
        "若在**连续两个 Step** 中 Memory 均表明「已满足完成条件」，**不得继续执行浏览操作**，必须立即调用 done 结束任务。\n\n"
    )
    if page_size < total:
        intro += f"本批共 {total} 条，按每 {page_size} 条一页处理；翻页后继续当前游标。\n\n"
    intro += "若出现「预览/确认/二次确认」页面：不再比对成绩，只点击【确认/提交】，然后继续流程。\n\n"
    body = records_to_task_text(records)
    return intro + "数据：\n" + body


def _get_llm():
    """从 .env 读取 DEEPSEEK_API_KEY / DEEPSEEK_BASE_URL，构造 ChatDeepSeek。"""
    try:
        from dotenv import load_dotenv
        load_dotenv()
    except ImportError:
        pass
    api_key = os.getenv("DEEPSEEK_API_KEY")
    base_url = os.getenv("DEEPSEEK_BASE_URL", "https://api.deepseek.com/v1")
    if not api_key:
        raise SystemExit(
            "未配置 DEEPSEEK_API_KEY。请在项目根目录 .env 中设置：\n"
            "  DEEPSEEK_API_KEY=sk-xxx\n"
            "  DEEPSEEK_BASE_URL=https://api.deepseek.com/v1  # 可选，默认即此"
        )
    try:
        from browser_use.llm import ChatDeepSeek
    except ModuleNotFoundError as e:
        _hint_browser_use_install()
        raise
    return ChatDeepSeek(api_key=api_key, base_url=base_url.rstrip("/"), model="deepseek-chat")


def _get_playwright_chromium_path() -> str | None:
    """若已安装 Playwright Chromium，返回其可执行路径，用于改善 CDP 兼容性；否则返回 None。"""
    try:
        from playwright.sync_api import sync_playwright
        with sync_playwright() as p:
            path = getattr(p.chromium, "executable_path", None)
            if path and os.path.isfile(path):
                return path
    except Exception:
        pass
    return None


def _patch_browser_session_connect():
    """对 BrowserSession.connect 打补丁：请求 /json/version 时按 deadline 等待+重试，等浏览器完全起来再连 CDP。"""
    import time

    import httpx
    from urllib.parse import urlparse, urlunparse

    from browser_use.browser.session import BrowserSession

    _original_connect = BrowserSession.connect
    _cdp_wait_seconds = 30
    _retry_sleep = 1.0

    def _extract_ws_url(version_info) -> str | None:
        if version_info.status_code != 200:
            return None
        if not (version_info.content and version_info.content.strip()):
            return None
        try:
            data = version_info.json()
            return data.get("webSocketDebuggerUrl") or None
        except Exception:
            return None

    async def _patched_connect(self, cdp_url: str | None = None):
        self.browser_profile.cdp_url = cdp_url or self.cdp_url
        if not self.cdp_url:
            raise RuntimeError("Cannot setup CDP connection without CDP URL")
        if self._cdp_client_root is not None:
            self.logger.warning(
                "⚠️ connect() called but CDP client already exists! Cleaning up old connection before creating new one."
            )
            try:
                await self._cdp_client_root.stop()
            except Exception as e:
                self.logger.debug(f"Error stopping old CDP client: {e}")
            self._cdp_client_root = None

        if not self.cdp_url.startswith("ws"):
            parsed_url = urlparse(self.cdp_url)
            path = parsed_url.path.rstrip("/")
            if not path.endswith("/json/version"):
                path = path + "/json/version"
            url = urlunparse(
                (parsed_url.scheme, parsed_url.netloc, path, parsed_url.params, parsed_url.query, parsed_url.fragment)
            )
            deadline = time.time() + _cdp_wait_seconds
            last_err = None
            ws_url = None

            # 请求 CDP 的 localhost 时必须绕过系统代理，否则会经代理挂起并超时
            while time.time() < deadline:
                try:
                    async with httpx.AsyncClient(trust_env=False) as client:
                        headers = self.browser_profile.headers or {}
                        version_info = await client.get(url, headers=headers)
                        ws_url = _extract_ws_url(version_info)
                        if ws_url:
                            self.browser_profile.cdp_url = ws_url
                            break
                        last_err = ValueError(
                            f"/json/version status={version_info.status_code} or invalid/empty body"
                        )
                except Exception as e:
                    last_err = e
                await asyncio.sleep(_retry_sleep)

            if not ws_url:
                raise RuntimeError(
                    f"CDP not ready after {_cdp_wait_seconds}s, last error: {last_err}"
                )

        return await _original_connect(self, cdp_url=self.browser_profile.cdp_url)

    BrowserSession.connect = _patched_connect


async def _run_browser_only(headless: bool) -> None:
    """仅验证浏览器/CDP 链路：start() 成功即说明浏览器已起来且 CDP 可用，不跑 Agent。"""
    _patch_browser_session_connect()
    from browser_use import Browser

    browser_kw: dict = {"headless": headless}
    exe = os.getenv("BROWSER_EXECUTABLE_PATH") or _get_playwright_chromium_path()
    if exe and os.path.isfile(exe):
        browser_kw["executable_path"] = exe
    browser = Browser(**browser_kw)
    print(">>> browser about to start", flush=True)
    await browser.start()
    print(">>> 浏览器启动成功，CDP 链路正常。", flush=True)
    await browser.kill()
    print(">>> 自检结束（browser 已关闭）", flush=True)


def _hint_browser_use_install() -> None:
    """browser-use / DeepSeek 未安装时打印安装说明。"""
    print(
        "未找到 browser-use 或依赖。请在本项目虚拟环境中安装：\n"
        "  .venv\\Scripts\\python.exe -m pip install -r requirements.txt\n"
        "（需 Python 3.11+，且 .env 中配置 DEEPSEEK_API_KEY）",
        file=sys.stderr,
    )


async def _run_agent(
    task: str,
    max_steps: int,
    headless: bool,
    excel_path: Path | None = None,
) -> None:
    """使用 browser-use Agent + DeepSeek 执行任务。"""
    _patch_browser_session_connect()
    try:
        from browser_use import Agent, Browser
    except ModuleNotFoundError as e:
        _hint_browser_use_install()
        raise

    llm = _get_llm()
    browser_kw: dict = {"headless": headless}
    exe = os.getenv("BROWSER_EXECUTABLE_PATH") or _get_playwright_chromium_path()
    if exe and os.path.isfile(exe):
        browser_kw["executable_path"] = exe
    browser = Browser(**browser_kw)
    agent_kw: dict = {
        "task": task,
        "llm": llm,
        "browser": browser,
        "directly_open_url": True,
    }
    if excel_path is not None:
        path_abs = excel_path.resolve()
        if path_abs.exists():
            agent_kw["available_file_paths"] = [str(path_abs)]
    agent = Agent(**agent_kw)
    print(">>> agent about to run", flush=True)
    result = await agent.run(max_steps=max_steps)
    print(">>> agent finished", flush=True)
    if result and result.final_result():
        print("\n✅ Agent 完成。最终结果：", result.final_result())
    else:
        print("\n⚠️ Agent 已结束（可能未标记 done 或达到 max_steps）。")


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Excel → 强校验 → 任务生成 → DeepSeek + browser-use 自动成绩录入"
    )
    parser.add_argument("-e", "--excel", type=Path, default=None, help="Excel 文件路径（--browser-only 时可选）")
    parser.add_argument("-u", "--url", default=None, help="录入页 URL（--browser-only 时可选）")
    parser.add_argument("--sheet", default=None, help="工作表名或索引")
    parser.add_argument("--header-row", type=int, default=None, help="表头所在行号（0-based）。不传则自动在前几行中查找「平时成绩」「考试成绩」")
    parser.add_argument("--double-column", action="store_true", help="双列布局：每行拆成左、右两条记录，不丢右栏数据")
    parser.add_argument("--page-size", type=int, default=None, help="每批条数")
    parser.add_argument("--max-rows", type=int, default=None, help="最多处理行数")
    parser.add_argument("--max-steps", type=int, default=80, help="Agent 最大步数（默认 80）")
    parser.add_argument("--headless", action="store_true", help="无头模式运行浏览器（不显示窗口）")
    parser.add_argument("--dry-run", action="store_true", help="只做读取+校验+打印任务，不调 Agent")
    parser.add_argument(
        "--browser-only",
        action="store_true",
        help="仅验证浏览器/CDP 链路：启动 Browser 并 start()，不跑 Agent。用于排查「浏览器起来了但 agent 没上岗」问题。",
    )
    args = parser.parse_args()

    if args.browser_only:
        asyncio.run(_run_browser_only(headless=args.headless))
        return

    if not args.excel or not args.url:
        parser.error("填表模式需要 -e/--excel 与 -u/--url（仅 --browser-only 时可省略）")

    records, meta = read_excel_to_records(
        args.excel,
        sheet=args.sheet,
        header_row=args.header_row,
        double_column=True if args.double_column else None,
    )
    if args.max_rows:
        records = records[: args.max_rows]

    ok, msg = validate_records_for_fill(records, meta)
    if not ok:
        print("【强校验未通过】禁止自动填表。", file=sys.stderr)
        print(msg, file=sys.stderr)
        sys.exit(1)

    print(f"已读取 {len(records)} 条记录，已同时识别「平时成绩」与「考试成绩」列，可安全填表。")
    task = build_task(
        records,
        args.url,
        page_size=args.page_size,
        excel_path=args.excel,
    )

    if args.dry_run:
        print("\n--- 任务文案（dry-run）---\n")
        print(task)
        return

    print("\n--- 启动 browser-use Agent（DeepSeek）---\n")
    max_connect_retries = 3
    retry_delay_sec = 4
    last_error = None
    for attempt in range(max_connect_retries):
        try:
            asyncio.run(
                _run_agent(
                    task,
                    max_steps=args.max_steps,
                    headless=args.headless,
                    excel_path=args.excel,
                )
            )
            return
        except json.JSONDecodeError as e:
            last_error = e
            if attempt < max_connect_retries - 1:
                print(
                    f"【浏览器 CDP 连接失败（第 {attempt + 1}/{max_connect_retries} 次）】正在 {retry_delay_sec} 秒后重试…",
                    file=sys.stderr,
                )
                time.sleep(retry_delay_sec)
            else:
                print(
                    "【浏览器连接失败】browser-use 在连接浏览器 CDP 时解析版本信息失败（非 JSON 或空响应）。\n"
                    "建议：1) 安装 Playwright Chromium 后重试：uvx playwright install chromium 或 python -m playwright install chromium\n"
                    "      2) 确认已安装 Chrome/Chromium/Edge 且可正常启动；3) 检查本机防火墙/安全软件是否拦截 localhost。\n"
                    "详见：https://github.com/browser-use/browser-use/issues",
                    file=sys.stderr,
                )
                sys.exit(1)
        except Exception as e:
            last_error = e
            if "CDP" in str(e) or "connect" in str(e).lower() or "webSocket" in str(e):
                if attempt < max_connect_retries - 1:
                    print(
                        f"【浏览器连接异常（第 {attempt + 1}/{max_connect_retries} 次）】{e}\n{retry_delay_sec} 秒后重试…",
                        file=sys.stderr,
                    )
                    time.sleep(retry_delay_sec)
                    continue
            raise
    if last_error:
        raise last_error


if __name__ == "__main__":
    main()
