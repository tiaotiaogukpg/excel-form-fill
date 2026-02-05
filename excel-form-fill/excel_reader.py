"""
Phase 1：Excel → 结构化数据（唯一正确来源）

原则：只做读取，不做猜测。
- 只从**明确语义**的「平时成绩」「考试成绩」列取值。
- 绝不把单一「成绩」列自动复制成 平时+考试。
- 返回列识别元数据，供上层做强校验（未同时识别两列 → 禁止填表，输出诊断）。
"""
from pathlib import Path
from typing import Any, List, Optional, TypedDict

import pandas as pd


USUAL_SCORE_KEY = "平时成绩"
EXAM_SCORE_KEY = "考试成绩"


class ReadMeta(TypedDict):
    """读取后的列识别元数据，用于强校验。"""
    has_usual_column: bool
    has_exam_column: bool
    raw_columns: List[str]


# 只认「平时成绩」语义列，不认通用「成绩」
USUAL_HEADER_ALIASES = ["平时成绩", "平时", "平时分"]
# 只认「考试成绩」语义列
EXAM_HEADER_ALIASES = ["考试成绩", "考试", "期末成绩"]


def _norm(s: str) -> str:
    return str(s).strip().replace(" ", "").replace("\u3000", "")


def _make_column_names_unique(columns: List[Any]) -> List[str]:
    """重复列名 → 原名、原名.1、原名.2 ..."""
    seen: dict[str, int] = {}
    out: List[str] = []
    for c in columns:
        name = _norm(str(c)) or "Unnamed"
        if name in seen:
            seen[name] += 1
            out.append(f"{name}.{seen[name]}")
        else:
            seen[name] = 0
            out.append(name)
    return out


def _pick_column(columns: List[str], aliases: List[str]) -> Optional[str]:
    """按语义取第一个匹配列名；完全匹配优先。"""
    norm_to_orig = {_norm(c): c for c in columns}
    for a in aliases:
        key = _norm(a)
        if key in norm_to_orig:
            return norm_to_orig[key]
    for col in columns:
        n = _norm(col)
        for a in aliases:
            if _norm(a) in n or n in _norm(a):
                return col
    return None


def _to_int_score(v: Any) -> Optional[int]:
    if v is None:
        return None
    if isinstance(v, float) and pd.isna(v):
        return None
    if isinstance(v, str) and not v.strip():
        return None
    try:
        n = int(round(float(v)))
    except Exception:
        return None
    return max(0, min(100, n))


def read_excel_to_records(
    excel_path: Path,
    sheet: Optional[str | int] = None,
    double_column: Optional[bool] = None,
) -> tuple[List[dict[str, Any]], ReadMeta]:
    """
    将 Excel 解析为「表头→行数据」的字典列表。

    - 只从**明确**的「平时成绩」「考试成绩」列取值，绝不猜、不复制。
    - 未识别到的列为 None。
    - 返回 (records, meta)，meta 含 has_usual_column / has_exam_column，供强校验使用。
    """
    excel_path = Path(excel_path)
    sheet_name = sheet if sheet is not None else 0
    df = pd.read_excel(excel_path, sheet_name=sheet_name, header=0)
    raw_columns = list(df.columns)
    columns = _make_column_names_unique(raw_columns)
    df.columns = columns

    usual_col = _pick_column(columns, USUAL_HEADER_ALIASES)
    exam_col = _pick_column(columns, EXAM_HEADER_ALIASES)

    meta: ReadMeta = {
        "has_usual_column": usual_col is not None,
        "has_exam_column": exam_col is not None,
        "raw_columns": raw_columns,
    }

    records: List[dict[str, Any]] = []
    for idx, row in df.iterrows():
        rec: dict[str, Any] = {}
        for orig, col in zip(raw_columns, columns):
            val = row[col]
            if pd.isna(val):
                rec[orig] = ""
            elif isinstance(val, float):
                rec[orig] = int(val) if val == int(val) else val
            else:
                rec[orig] = str(val).strip()

        # 只从明确列取值，不猜、不复制
        rec[USUAL_SCORE_KEY] = _to_int_score(row[usual_col]) if usual_col and usual_col in row.index else None
        rec[EXAM_SCORE_KEY] = _to_int_score(row[exam_col]) if exam_col and exam_col in row.index else None
        records.append(rec)
    return records, meta


def validate_records_for_fill(records: List[dict[str, Any]], meta: ReadMeta) -> tuple[bool, str]:
    """
    填表前强校验：未同时识别到「平时成绩」与「考试成绩」列 → 禁止自动填表，返回诊断信息。
    这是防数据事故的生命线。
    """
    if not meta["has_usual_column"] or not meta["has_exam_column"]:
        return False, (
            "未同时识别到「平时成绩」与「考试成绩」列，禁止自动填表。\n"
            f"当前识别：平时成绩列={meta['has_usual_column']}，考试成绩列={meta['has_exam_column']}。\n"
            f"Excel 表头：{meta['raw_columns']}\n"
            "请确保表头中同时存在「平时成绩」与「考试成绩」（或别名：平时/考试/期末成绩等）。"
        )
    if not records:
        return False, "Excel 无数据行，无需填表。"
    return True, ""


def records_to_task_text(records: List[dict[str, Any]]) -> str:
    """Phase 2 用：转成 Agent 能看懂的任务文本，格式：姓名 | 平时成绩(目标值) | 考试成绩(目标值)。"""
    lines = ["姓名 | 平时成绩(目标值) | 考试成绩(目标值)", "---"]
    for r in records:
        name = r.get("姓名") or r.get("学生姓名") or ""
        usual = r.get(USUAL_SCORE_KEY)
        exam = r.get(EXAM_SCORE_KEY)
        usual_s = str(usual) if usual is not None else ""
        exam_s = str(exam) if exam is not None else ""
        lines.append(f"{name} | {usual_s} | {exam_s}")
    return "\n".join(lines)
