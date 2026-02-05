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
    header_row: int


# 只认「平时成绩」语义列，不认通用「成绩」
USUAL_HEADER_ALIASES = ["平时成绩", "平时", "平时分"]
# 只认「考试成绩」语义列
EXAM_HEADER_ALIASES = ["考试成绩", "考试", "期末成绩"]


def _norm(s: str) -> str:
    """规范化表头/列名：去首尾空白、空格、全角空格、换行。"""
    return str(s).strip().replace(" ", "").replace("\u3000", "").replace("\n", "").replace("\r", "")


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


def _pick_column(columns: List[str], aliases: List[str], exclude: Optional[str] = None) -> Optional[str]:
    """
    按语义取第一个匹配列名（可排除某列）；完全匹配优先。
    只认「列名包含别名」或「列名等于别名」，不认「列名仅为别名子串」：
    例如单列「成绩」不匹配「平时成绩」「考试成绩」，避免误绑。
    """
    norm_to_orig = {_norm(c): c for c in columns}
    for a in aliases:
        key = _norm(a)
        if key in norm_to_orig:
            cand = norm_to_orig[key]
            if exclude is not None and cand == exclude:
                continue
            return cand
    for col in columns:
        if exclude is not None and col == exclude:
            continue
        n = _norm(col)
        for a in aliases:
            # 仅当列名包含完整别名时匹配，不采用「列名 in 别名」以免「成绩」匹配「平时成绩」
            if _norm(a) in n:
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


def _is_likely_student_record(rec: dict[str, Any]) -> bool:
    """
    判断该行是否像「学生成绩记录」：姓名像人名且至少有一项成绩。
    用于过滤成绩报告单中的标题、说明、班级、任课教师、空行、权重行（如 60%/40%）等非数据行。
    """
    name = _get_name_from_record(rec)
    if not name or len(name) < 2 or len(name) > 10:
        return False
    # 排除表头下常见的「权重行」（30/70、40/60、50/50 等，和为 100 的整数百分比）
    usual, exam = rec.get(USUAL_SCORE_KEY), rec.get(EXAM_SCORE_KEY)
    weight_vals = (30, 40, 50, 60, 70)
    if usual in weight_vals and exam in weight_vals and usual + exam == 100:
        return False
    # 排除明显非人名的整段关键词（避免误杀含单字的人名）
    skip_phrases = (
        "班级", "课程", "教师", "成绩", "总评", "考查", "科目", "任课", "体质", "检测", "平时", "情况",
        "生成", "其中", "占", "由", "与", "健康", "体育", "数字", "：", ":", "。", ".",
    )
    for s in skip_phrases:
        if s in name:
            return False
    # 姓名不应全是数字
    if name.replace(" ", "").isdigit():
        return False
    # 至少有一项成绩为 0–100 的整数
    if usual is None and exam is None:
        return False
    if usual is not None and (usual < 0 or usual > 100):
        return False
    if exam is not None and (exam < 0 or exam > 100):
        return False
    return True


def _cell_str(val: Any) -> str:
    """单元格转为字符串，空/NaN 视为空串。"""
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return ""
    return str(val).strip()


def _row_to_col_names(row: pd.Series) -> List[str]:
    """把一行单元格转为列名列表（去空、规范化前）。"""
    return [_cell_str(row.iloc[j]) for j in range(len(row))]


def _find_header_row(
    excel_path: Path, sheet_name: str | int, max_rows: int = 15
) -> tuple[int, Optional[List[str]], Optional[int]]:
    """
    不限定成绩在第几行：从表顶逐行扫描，在首次出现「平时」+「考试」语义时开始识别表头。
    支持：单行表头（该行即含平时/考试）、多行表头（该行与上方若干行合并）。
    返回 (header_row, combined_columns, score_row)。
    combined_columns 非空时数据从 score_row+1 行起；否则数据从 header_row+1 行起。
    """
    df_raw = pd.read_excel(excel_path, sheet_name=sheet_name, header=None)
    nrows = min(max_rows, len(df_raw))
    for score_row in range(nrows):
        row_score = df_raw.iloc[score_row]
        raw_names = _row_to_col_names(row_score)
        cols = _make_column_names_unique(raw_names)
        usual = _pick_column(cols, USUAL_HEADER_ALIASES)
        exam = _pick_column(cols, EXAM_HEADER_ALIASES, exclude=usual)
        if usual is not None and exam is not None:
            return (score_row, None, None)
        for main_row in range(max(0, score_row - 3), score_row):
            row_main = df_raw.iloc[main_row]
            ncols = max(len(row_main), len(row_score))
            combined = []
            for j in range(ncols):
                cur = _cell_str(row_main.iloc[j]) if j < len(row_main) else ""
                sub_val = _cell_str(row_score.iloc[j]) if j < len(row_score) else ""
                sub_norm = _norm(sub_val)
                if sub_norm and (
                    any(_norm(a) in sub_norm for a in USUAL_HEADER_ALIASES + EXAM_HEADER_ALIASES)
                    or "姓名" in sub_norm
                ):
                    combined.append(sub_val if sub_val else cur)
                else:
                    combined.append(cur if cur else sub_val)
            cols = _make_column_names_unique(combined)
            usual = _pick_column(cols, USUAL_HEADER_ALIASES)
            exam = _pick_column(cols, EXAM_HEADER_ALIASES, exclude=usual)
            if usual is not None and exam is not None:
                return (main_row, cols, score_row)
    return (0, None, None)


def _is_xuhao_column(c: Any) -> bool:
    """列名是否为「序号」或「序号.1」「序号.2」等（去重后的右栏序号列）。"""
    n = _norm(str(c))
    return n == "序号" or n.startswith("序号.")

def _column_matches_usual(c: Any) -> bool:
    """列名是否匹配「平时」语义。"""
    n = _norm(str(c))
    return any(_norm(a) in n for a in USUAL_HEADER_ALIASES)


def _column_matches_exam(c: Any) -> bool:
    """列名是否匹配「考试」语义。"""
    n = _norm(str(c))
    return any(_norm(a) in n for a in EXAM_HEADER_ALIASES)


def _find_double_column_split(raw_columns: List[Any], columns: List[str]) -> Optional[int]:
    """
    检测是否为双列布局，返回右栏起始列索引（0-based）。
    - 优先：第二个「序号」列（含 序号.1）。
    - 其次：第二个「学生姓名」或「姓名」列。
    - 再次：首个带 .1 的列，右栏起点前推 2 列（且右块含平时+考试）。
    - 再再次：存在「第二个平时列」时，以该列索引减 2 为右栏起点（兼容无 .1 的列名如 Unnamed）。
    - 最后：空列。
    """
    xuhao_indices = [i for i, c in enumerate(raw_columns) if _is_xuhao_column(c)]
    if len(xuhao_indices) >= 2:
        return xuhao_indices[1]
    name_indices = [i for i, c in enumerate(raw_columns) if "姓名" in _norm(str(c))]
    if len(name_indices) >= 2:
        j = name_indices[1]
        if j > 0 and _is_xuhao_column(raw_columns[j - 1]):
            return j - 1
        return j
    first_dot1 = next((i for i, c in enumerate(raw_columns) if ".1" in _norm(str(c))), None)
    if first_dot1 is not None and first_dot1 >= 2:
        split_candidate = first_dot1 - 2
        cols_right = columns[split_candidate:]
        u = _pick_column(cols_right, USUAL_HEADER_ALIASES)
        e = _pick_column(cols_right, EXAM_HEADER_ALIASES, exclude=u)
        if u is not None and e is not None:
            return split_candidate
    usual_indices = [i for i, c in enumerate(columns) if _column_matches_usual(c)]
    if len(usual_indices) >= 2:
        split_candidate = max(0, usual_indices[1] - 2)
        cols_right = columns[split_candidate:]
        u = _pick_column(cols_right, USUAL_HEADER_ALIASES)
        e = _pick_column(cols_right, EXAM_HEADER_ALIASES, exclude=u)
        if u is not None and e is not None:
            return split_candidate
    for i, c in enumerate(raw_columns):
        if i > 0 and (_norm(str(c)) in ("", "Unnamed") or str(c).strip() == ""):
            return i
    return None


def _row_to_record(
    row: pd.Series,
    raw_columns: List[Any],
    columns: List[str],
    usual_col: Optional[str],
    exam_col: Optional[str],
) -> dict[str, Any]:
    """从一行中按给定列名取左块或右块，生成一条记录（含 平时成绩/考试成绩）。"""
    rec: dict[str, Any] = {}
    for orig, col in zip(raw_columns, columns):
        if col not in row.index:
            continue
        val = row[col]
        if pd.isna(val):
            rec[orig] = ""
        elif isinstance(val, float):
            rec[orig] = int(val) if val == int(val) else val
        else:
            rec[orig] = str(val).strip()
    rec[USUAL_SCORE_KEY] = _to_int_score(row[usual_col]) if usual_col and usual_col in row.index else None
    rec[EXAM_SCORE_KEY] = _to_int_score(row[exam_col]) if exam_col and exam_col in row.index else None
    return rec


def read_excel_to_records(
    excel_path: Path,
    sheet: Optional[str | int] = None,
    double_column: Optional[bool] = None,
    header_row: Optional[int] = None,
    filter_non_data_rows: bool = True,
) -> tuple[List[dict[str, Any]], ReadMeta]:
    """
    将 Excel 解析为「表头→行数据」的字典列表。

    - header_row: 表头所在行号（0-based）。None 时自动在前几行中查找同时含「平时成绩」「考试成绩」的行。
    - double_column: True=强制双列；False=单表；None=自动检测（表头出现两处「姓名」或中间有空列则按双列处理）。
    - filter_non_data_rows: 为 True 时只保留「像学生记录」的行。
    - 双列时：每行拆成左、右两条记录，平时/考试成绩分别在各自块内取对应列，不丢右栏数据。
    """
    excel_path = Path(excel_path)
    sheet_name = sheet if sheet is not None else 0
    if header_row is None:
        header_row, combined_columns, sub_header_row = _find_header_row(excel_path, sheet_name)
    else:
        combined_columns = None
        sub_header_row = None
    if combined_columns is not None and sub_header_row is not None:
        df_raw = pd.read_excel(excel_path, sheet_name=sheet_name, header=None)
        data_start = sub_header_row + 1
        df = df_raw.iloc[data_start:].reset_index(drop=True)
        n = df.shape[1]
        raw_columns = (
            list(combined_columns[:n])
            if len(combined_columns) >= n
            else list(combined_columns) + [f"Unnamed:{i}" for i in range(len(combined_columns), n)]
        )
        columns = _make_column_names_unique(raw_columns)
        df.columns = columns
    else:
        df = pd.read_excel(excel_path, sheet_name=sheet_name, header=header_row)
        raw_columns = list(df.columns)
        columns = _make_column_names_unique([str(c) for c in raw_columns])
        df.columns = columns

    split_at = _find_double_column_split(raw_columns, columns) if double_column is not False else None
    if split_at is None and double_column is True and len(raw_columns) >= 16:
        split_at = len(raw_columns) // 2

    records: List[dict[str, Any]] = []
    has_usual = False
    has_exam = False

    if split_at is not None and split_at > 0 and split_at < len(raw_columns):
        # 双列布局：左块 0..split_at-1，右块 split_at..end
        raw_left = raw_columns[:split_at]
        raw_right = raw_columns[split_at:]
        cols_left = columns[:split_at]
        cols_right = columns[split_at:]
        usual_left = _pick_column(cols_left, USUAL_HEADER_ALIASES)
        exam_left = _pick_column(cols_left, EXAM_HEADER_ALIASES, exclude=usual_left)
        usual_right = _pick_column(cols_right, USUAL_HEADER_ALIASES)
        exam_right = _pick_column(cols_right, EXAM_HEADER_ALIASES, exclude=usual_right)
        if usual_right is None and len(cols_right) > 2:
            usual_right = cols_right[2]
        if exam_right is None and len(cols_right) > 6:
            exam_right = cols_right[6]
        has_usual = (usual_left is not None or usual_right is not None)
        has_exam = (exam_left is not None or exam_right is not None)

        for idx, row in df.iterrows():
            rec_left = _row_to_record(row, raw_left, cols_left, usual_left, exam_left)
            rec_right = _row_to_record(row, raw_right, cols_right, usual_right, exam_right)
            records.append(rec_left)
            records.append(rec_right)
    else:
        # 单表：确保平时列 ≠ 考试列，避免同一列被当两列用
        usual_col = _pick_column(columns, USUAL_HEADER_ALIASES)
        exam_col = _pick_column(columns, EXAM_HEADER_ALIASES, exclude=usual_col)
        has_usual = usual_col is not None
        has_exam = exam_col is not None
        for idx, row in df.iterrows():
            rec = _row_to_record(row, raw_columns, columns, usual_col, exam_col)
            records.append(rec)

    meta: ReadMeta = {
        "has_usual_column": has_usual,
        "has_exam_column": has_exam,
        "raw_columns": raw_columns,
        "header_row": header_row,
    }
    if filter_non_data_rows:
        n_before = len(records)
        sample = records[0] if records else None
        records = [r for r in records if _is_likely_student_record(r)]
        if n_before > 0 and len(records) == 0 and sample is not None:
            meta["filtered_from"] = n_before
            meta["sample_record"] = sample
    records.sort(key=_get_xuhao_from_record)
    return records, meta


def validate_records_for_fill(records: List[dict[str, Any]], meta: ReadMeta) -> tuple[bool, str]:
    """
    填表前强校验：未同时识别到「平时成绩」与「考试成绩」列 → 禁止自动填表，返回诊断信息。
    这是防数据事故的生命线。
    """
    if not meta["has_usual_column"] or not meta["has_exam_column"]:
        raw = meta["raw_columns"]
        norm_list = [_norm(str(c)) for c in raw]
        has_only_score = "成绩" in norm_list or "总分" in norm_list or "总评" in norm_list
        hint = ""
        if has_only_score and not (meta["has_usual_column"] or meta["has_exam_column"]):
            hint = (
                "\n本表仅有「成绩」类列，未区分平时/考试。"
                "请将表头改为「平时成绩」「考试成绩」两列，或使用含该两列的 Excel 再填表。"
            )
        return False, (
            "未同时识别到「平时成绩」与「考试成绩」列，禁止自动填表。\n"
            f"当前识别：平时成绩列={meta['has_usual_column']}，考试成绩列={meta['has_exam_column']}。\n"
            f"表头行（第 {meta.get('header_row', 0) + 1} 行）：{raw}\n"
            "请确保表头中同时存在「平时成绩」与「考试成绩」（或别名：平时/考试/期末成绩等）。"
            + hint
        )
    if not records:
        msg = "Excel 无数据行，无需填表。"
        n_raw = meta.get("filtered_from")
        sample = meta.get("sample_record")
        if n_raw is not None and n_raw > 0 and sample is not None:
            name = _get_name_from_record(sample)
            usual = sample.get(USUAL_SCORE_KEY)
            exam = sample.get(EXAM_SCORE_KEY)
            msg += (
                f"\n诊断：过滤前共 {n_raw} 行，过滤后 0 行。"
                f" 首行示例：姓名=%r, 平时成绩=%s, 考试成绩=%s。"
                " 若姓名为空请检查表头是否含「姓名」或「学生」列；若为权重行(如30/70)会被自动过滤。"
            ) % (name or "(未识别)", usual, exam)
        return False, msg
    return True, ""


def _get_xuhao_from_record(r: dict[str, Any]) -> int:
    """从记录中取序号（用于排序）；无序号或非数字时返回 999999 以便排到末尾。"""
    for k, v in r.items():
        if k in (USUAL_SCORE_KEY, EXAM_SCORE_KEY):
            continue
        if not _is_xuhao_column(k):
            continue
        if v is None or (isinstance(v, float) and pd.isna(v)) or str(v).strip() == "":
            continue
        try:
            return int(round(float(v)))
        except (ValueError, TypeError):
            continue
    first_key = next((k for k in r if k not in (USUAL_SCORE_KEY, EXAM_SCORE_KEY)), None)
    if first_key is not None:
        v = r[first_key]
        if v is not None and v != "" and not (isinstance(v, float) and pd.isna(v)):
            try:
                return int(round(float(v)))
            except (ValueError, TypeError):
                pass
    return 999999


def _get_name_from_record(r: dict[str, Any]) -> str:
    """
    从记录中按语义取姓名；表头含「姓名」即认（支持 学生\\n姓名、姓名（必填） 等）。
    若无「姓名」列但某列含「学生」且该格值像人名，也认作姓名列。
    兜底：首个非成绩列且值像人名（2–10 字、非关键词、非纯数字）也认。
    """
    skip_phrases = ("班级", "课程", "教师", "成绩", "总评", "考查", "科目", "任课", "体质", "检测", "序号")
    skip_norm = ("序号", "平时成绩", "考试成绩", "总评", "备注")

    def _looks_like_name(val: str) -> bool:
        if not val or len(val) < 2 or len(val) > 10:
            return False
        if any(s in val for s in skip_phrases):
            return False
        if val.replace(" ", "").isdigit():
            return False
        return True

    for k, v in r.items():
        if k in (USUAL_SCORE_KEY, EXAM_SCORE_KEY):
            continue
        if v in (None, ""):
            continue
        nk = _norm(str(k))
        if "姓名" in nk:
            return str(v).strip()
    for k, v in r.items():
        if k in (USUAL_SCORE_KEY, EXAM_SCORE_KEY):
            continue
        val = str(v).strip() if v else ""
        nk = _norm(str(k))
        if "学生" in nk and _looks_like_name(val):
            return val
    for k, v in r.items():
        if k in (USUAL_SCORE_KEY, EXAM_SCORE_KEY):
            continue
        if _norm(str(k)) in skip_norm:
            continue
        val = str(v).strip() if v else ""
        if _looks_like_name(val):
            return val
    return ""


def records_to_task_text(records: List[dict[str, Any]]) -> str:
    """Phase 2 用：转成 Agent 能看懂的任务文本，格式：姓名 | 平时成绩(目标值) | 考试成绩(目标值)。"""
    lines = ["姓名 | 平时成绩(目标值) | 考试成绩(目标值)", "---"]
    for r in records:
        name = _get_name_from_record(r)
        usual = r.get(USUAL_SCORE_KEY)
        exam = r.get(EXAM_SCORE_KEY)
        usual_s = str(usual) if usual is not None else ""
        exam_s = str(exam) if exam is not None else ""
        lines.append(f"{name} | {usual_s} | {exam_s}")
    return "\n".join(lines)
