"""
从 Excel 成绩单导出 grades.json（报表/规范表均兼容）。

报表解析设计原则（不信版面，只信语义）：
- 不信版面，只信语义：不依赖列号、空列、合并单元格等版面信息。
- 不按块切表，按行读人：按行扫描，同一行左栏/右栏都可能是学生，统一归一。
- 左右两栏统一归一化：多个「学生姓名」列都识别，全部归入同一课程、同一学生列表。
"""
import argparse
import json
from dataclasses import asdict, dataclass
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

import pandas as pd


@dataclass
class GradeRow:
    name: str
    class_name: Optional[str]
    course: Optional[str]
    usual: Optional[int]
    exam: Optional[int]
    final: Optional[int]
    source_row: int


def _norm_col(s: str) -> str:
    return str(s).strip().lower().replace(" ", "")


def _pick_col(columns: List[str], aliases: List[str]) -> Optional[str]:
    norm_map = {_norm_col(c): c for c in columns}
    for a in aliases:
        if _norm_col(a) in norm_map:
            return norm_map[_norm_col(a)]
    return None


def _to_int(v: Any) -> Optional[int]:
    if v is None:
        return None
    if isinstance(v, float) and pd.isna(v):
        return None
    if isinstance(v, str) and v.strip() == "":
        return None
    try:
        n = int(round(float(v)))
    except Exception:
        return None
    return max(0, min(100, n))


def _cell_text(v: Any) -> str:
    if v is None:
        return ""
    if isinstance(v, float) and pd.isna(v):
        return ""
    return str(v)


def _find_cell_contains(df: "pd.DataFrame", keyword: str, max_rows: int = 40) -> Optional[Tuple[int, int]]:
    rmax = min(max_rows, len(df))
    for r in range(rmax):
        for c in range(df.shape[1]):
            if keyword in _cell_text(df.iat[r, c]):
                return r, c
    return None


def _find_row_contains(df: "pd.DataFrame", keyword: str, r_from: int, r_to: int) -> Optional[int]:
    for r in range(max(0, r_from), min(len(df) - 1, r_to) + 1):
        row = [_cell_text(df.iat[r, c]) for c in range(df.shape[1])]
        if any(keyword in x for x in row):
            return r
    return None


def _is_assessment_subject_sheet(sheet_name: Optional[str]) -> bool:
    """工作表名是否为「考查科目」/「考察科目」或含该字样。"""
    s = (sheet_name or "").strip()
    if not s:
        return False
    return (
        s == "考查科目"
        or s == "考察科目"
        or ("考查" in s and "科目" in s)
        or ("考察" in s and "科目" in s)
    )


def _is_likely_course_name(v: str) -> bool:
    """像「语文」「数学」等课程名，排除整段标题、纯标签（班级/课程/科目/任课教师等）、班级名（xxx班）及「课程：xxx」整格。"""
    s = (v or "").strip()
    if not s or len(s) > 20:
        return False
    if "报告单" in s or "分析表" in s or ("成绩" in s and "表" in s):
        return False
    if s in ("班级", "课程", "科目") or s.rstrip("：: ") in ("班级", "课程", "科目"):
        return False
    if s.startswith("课程") or s.startswith("科目") or "任课教师" in s or s.endswith("老师"):
        return False
    if s.endswith("班") or s.endswith("：") or s.endswith(":"):
        return False
    return True


def _extract_meta_value(df: "pd.DataFrame", keyword: str, max_rows: int = 8) -> Optional[str]:
    """
    从报表顶部区域提取类似 “班级：xxx / 课程：xxx” 的值（先行后列，第一个命中即返回）。
    """
    for r in range(min(max_rows, len(df))):
        for c in range(df.shape[1]):
            s = _cell_text(df.iat[r, c]).strip()
            if not s or keyword not in s:
                continue
            for sep in ["：", ":", " "]:
                if sep in s:
                    left, right = s.split(sep, 1)
                    if keyword in left:
                        v = (right.strip() or None)
                        if v:
                            return v
                    break
            else:
                v = s.replace(keyword, "").strip("：: ").strip() or None
                if v:
                    return v
    return None


def _extract_course_from_title_area(
    df: "pd.DataFrame",
    resolved_sheet: str,
    default_course: Optional[str],
    max_title_rows: int = 8,
) -> Optional[str]:
    """
    从工作表标题区读取科目（课程名）。部分表头只写「科目」、部分写「课程」，两种都认。
    - 考查/考察表：从「含班级」的那一行及其后 1～2 行（标题窗）内收集「课程：value」「科目：value」与单独一格；课程可能在班级下一行（如 K3）。收集后优先「课程」再「科目」再单独格，规则同前。
    - 非考查表：先在全标题区找「课程：xxx」，若无再找「科目：xxx」，第一个命中即返回。
    """
    is_assessment = _is_assessment_subject_sheet(resolved_sheet)

    # 考查表：从含「班级」行及其后 1～2 行（标题窗）内取课程，兼容「课程: 语文」在 K3、班级在 B2 的布局
    if is_assessment:
        class_row: Optional[int] = None
        for r in range(min(max_title_rows, len(df))):
            for c in range(df.shape[1]):
                if "班级" in _cell_text(df.iat[r, c]):
                    class_row = r
                    break
            if class_row is not None:
                break
        if class_row is not None:
            title_window_end = min(class_row + 3, max_title_rows, len(df))  # 班级行 + 其后 2 行
            from_course: List[Tuple[str, int]] = []
            from_subject: List[Tuple[str, int]] = []
            standalone: List[Tuple[str, int]] = []
            for row_idx in range(class_row, title_window_end):
                row = [ _cell_text(df.iat[row_idx, c]).strip() for c in range(df.shape[1]) ]
                for s in row:
                    if not s:
                        continue
                    v: Optional[str] = None
                    used_label: Optional[str] = None
                    for label in ("课程", "科目"):
                        if label not in s:
                            continue
                        for sep in ["：", ":", " "]:
                            if sep in s:
                                parts = s.split(sep, 1)
                                if label in (parts[0] or "").strip():
                                    v = (parts[1] or "").strip()
                                    used_label = label
                                break
                        if v is not None:
                            break
                    if v is not None and _is_likely_course_name(v):
                        if used_label == "课程":
                            from_course.append((v, len(v)))
                        else:
                            from_subject.append((v, len(v)))
                    elif _is_likely_course_name(s):
                        standalone.append((s, len(s)))
            # 优先课程；无课程时再科目。课程/科目与单独格取最短（语文压过信息技术）；仅多个科目时：最短≤3字取最短（语文），否则取最长（思想道德与法治压过信息技术）
            if from_course:
                from_course.sort(key=lambda x: x[1])
                best_course = from_course[0]
            else:
                best_course = None
            if from_subject:
                best_subject_short = min(from_subject, key=lambda x: x[1])
                best_subject_long = max(from_subject, key=lambda x: x[1])
                # 仅科目无单独格时：若最短≤3字（如语文）取最短，否则取最长（如思想道德与法治 vs 信息技术）
                best_subject_only = (
                    best_subject_short[0]
                    if best_subject_short[1] <= 3
                    else best_subject_long[0]
                )
            else:
                best_subject_short = None
                best_subject_only = None
            if standalone:
                standalone.sort(key=lambda x: x[1])
                best_standalone = standalone[0]
            else:
                best_standalone = None
            if best_course is not None and best_standalone is not None:
                return best_course[0] if best_course[1] <= best_standalone[1] else best_standalone[0]
            if best_course is not None:
                return best_course[0]
            if best_subject_short is not None and best_standalone is not None:
                return best_subject_short[0] if best_subject_short[1] <= best_standalone[1] else best_standalone[0]
            if best_subject_only is not None:
                return best_subject_only
            if best_standalone is not None:
                return best_standalone[0]
        for r in range(min(max_title_rows, len(df))):
            for c in range(df.shape[1]):
                s = _cell_text(df.iat[r, c]).strip()
                if s and _is_likely_course_name(s):
                    return s
        return None

    # 非考查表：标题区先「课程」后「科目」
    v = _extract_meta_value(df, "课程", max_rows=max_title_rows)
    if v is not None:
        return v
    return _extract_meta_value(df, "科目", max_rows=max_title_rows)


# ---- 报表解析原则：不信版面，只信语义；不按块切表，按行读人；左右两栏统一归一化 ----

# 表头语义：凡出现「姓名」的列都视为“学生姓名列”，可能有多个（左栏 B、右栏 M 等）
NAME_HEADER_VALUES = ("学生姓名", "姓名")

def _is_header_or_empty_name(s: str) -> bool:
    """是否为表头或空，这类单元格不当作学生姓名。"""
    t = (s or "").strip()
    if not t:
        return True
    if t in NAME_HEADER_VALUES:
        return True
    return False


def read_excel_grades_report(
    excel_path: Path,
    sheet: Optional[str],
    default_class: Optional[str],
    default_course: Optional[str],
) -> Tuple[List[GradeRow], Dict[str, Any]]:
    """
    兼容“课程成绩报告单”一类报表格式：有标题行/合并单元格/多行表头、左右双栏。

    目标结构（单行驱动，课程只建一次）：
      读取 Excel → 识别课程（一次）→ 按行扫描 → 每一行识别 0/1/2 个学生 → 全部塞进同一个 Course
    绝对不做：发现新成绩表头 → new Course()
    """
    sheet_name = sheet if sheet is not None else 0
    resolved_sheet = sheet if sheet is not None else pd.ExcelFile(excel_path).sheet_names[0]
    df = pd.read_excel(excel_path, sheet_name=sheet_name, header=None)

    # ---------- Step 1：提前定位“姓名列”（可能有多个，如左栏 2、右栏 14），只做一次 ----------
    name_row: Optional[int] = None
    name_cols: List[int] = []
    rmax = min(40, len(df))
    for r in range(rmax):
        cols = []
        for c in range(df.shape[1]):
            if "姓名" in _cell_text(df.iat[r, c]):
                cols.append(c)
        if cols:
            name_row = r
            name_cols = cols
            break

    if name_row is None or not name_cols:
        raise ValueError("Excel 未找到“姓名”表头。该文件可能不是可解析的表格（例如扫描件/图片）。")

    score_row = name_row + 1
    if score_row >= len(df):
        score_row = name_row
    if _find_row_contains(df, "平时", score_row, score_row) is None:
        rr = _find_row_contains(df, "平时", name_row, min(name_row + 5, len(df) - 1))
        if rr is not None:
            score_row = rr

    def score_cell(c: int) -> str:
        return _cell_text(df.iat[score_row, c]).strip()

    weight_row = score_row + 1 if score_row + 1 < len(df) else score_row

    def _to_float(v: Any) -> Optional[float]:
        if v is None:
            return None
        if isinstance(v, float) and pd.isna(v):
            return None
        try:
            return float(v)
        except Exception:
            return None

    # ---------- 课程只读一次：表头「课程：xxx」或外部参数，全表只建一个 Course ----------
    class_name = _extract_meta_value(df, "班级") or default_class
    course = _extract_course_from_title_area(df, resolved_sheet, default_course) or default_course

    data_start = min(len(df), score_row + 2)

    def _is_likely_name(s: str) -> bool:
        t = s.strip()
        if not t:
            return False
        if len(t) > 10:
            return False
        for ch in ["，", "。", "：", ":", ";", "；", "、", "\n", "\t"]:
            if ch in t:
                return False
        if "说明" in t or "统计" in t or "成绩" in t:
            return False
        return True

    # ---------- 为每个姓名列建立「相对列偏移」：姓名列 → 平时列、考试列（不写死列号，兼容中间空列） ----------
    groups: List[Dict[str, Any]] = []
    column_groups: List[Tuple[int, int, int, Optional[float], Optional[float]]] = []
    for name_col in name_cols:
        usual_col: Optional[int] = None
        for c in range(name_col + 1, df.shape[1]):
            if "平时" in score_cell(c):
                usual_col = c
                break
        if usual_col is None:
            continue
        exam_col: Optional[int] = None
        for c in range(usual_col + 1, df.shape[1]):
            v = df.iat[weight_row, c]
            try:
                f = float(v)
            except Exception:
                continue
            if abs(f - 0.4) < 1e-6:
                exam_col = c
                break
        if exam_col is None:
            for c in range(usual_col + 1, df.shape[1]):
                s = score_cell(c)
                if ("考试" in s) or ("期末" in s):
                    exam_col = c
                    break
        if exam_col is None:
            continue
        w_usual = _to_float(df.iat[weight_row, usual_col])
        w_exam = _to_float(df.iat[weight_row, exam_col])
        has_weights = (
            w_usual is not None
            and w_exam is not None
            and 0 < w_usual <= 1
            and 0 < w_exam <= 1
            and abs((w_usual + w_exam) - 1) < 0.02
        )
        groups.append({
            "name_col": name_col,
            "usual_col": usual_col,
            "exam_col": exam_col,
            "weights": {"usual": w_usual, "exam": w_exam} if has_weights else None,
        })
        column_groups.append((
            name_col,
            usual_col,
            exam_col,
            w_usual if has_weights else None,
            w_exam if has_weights else None,
        ))

    # ---------- Step 2：逐行扫描（单行驱动）。一行可解析 0/1/2 个学生，全部塞进同一个 Course ----------
    # for row in dataRows:
    #     for each nameCol in nameCols:
    #         name = cell(row, nameCol)
    #         if 合法学生姓名(name): addStudent(course, parseStudent(row, nameCol))
    by_name: Dict[str, GradeRow] = {}
    for r in range(data_start, len(df)):
        for name_col, usual_col, exam_col, w_usual, w_exam in column_groups:
            name = _cell_text(df.iat[r, name_col]).strip()
            # 防炸 1：排除表头（学生姓名、空、null）
            if _is_header_or_empty_name(name):
                continue
            if not _is_likely_name(name):
                continue
            raw_usual = _to_float(df.iat[r, usual_col])
            raw_exam = _to_float(df.iat[r, exam_col])
            # 防炸 2：排除成绩全空
            if raw_usual is None and raw_exam is None:
                continue
            if w_usual is not None and w_exam is not None:
                usual = _to_int((raw_usual or 0) * w_usual)
                exam = _to_int((raw_exam or 0) * w_exam)
            else:
                usual = _to_int(raw_usual)
                exam = _to_int(raw_exam)
            final = None
            if usual is not None or exam is not None:
                final = (usual or 0) + (exam or 0)
            # 防炸 3：同一姓名保留最后一次出现（by_name 键为姓名，不重复建多条）
            by_name[name] = GradeRow(
                name=name,
                class_name=class_name,
                course=course,
                usual=usual,
                exam=exam,
                final=final,
                source_row=int(r) + 1,
            )

    rows = list(by_name.values())

    meta = {
        "excel": str(excel_path),
        "sheet": resolved_sheet,
        "mode": "report",
        "detected": {
            "name_row": name_row,
            "name_cols": name_cols,
            "score_row": score_row,
            "groups": groups,
        },
        "class_name": class_name,
        "course": course,
        "count": len(rows),
    }
    return rows, meta


def read_excel_grades(
    excel_path: Path,
    sheet: Optional[str],
    default_class: Optional[str],
    default_course: Optional[str],
) -> Tuple[List[GradeRow], Dict[str, Any]]:
    sheet_name = sheet if sheet is not None else 0
    resolved_sheet = sheet if sheet is not None else pd.ExcelFile(excel_path).sheet_names[0]
    df = pd.read_excel(excel_path, sheet_name=sheet_name)
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]

    cols = list(df.columns)
    col_name = _pick_col(cols, ["姓名", "name", "student", "学生姓名"])
    col_class = _pick_col(cols, ["班级", "class", "classname", "行政班"])
    col_course = _pick_col(cols, ["课程", "course", "科目", "学科"])
    col_usual = _pick_col(cols, ["平时成绩", "平时", "usual", "平时分", "过程性评价"])
    col_exam = _pick_col(cols, ["考试成绩", "考试", "exam", "期末", "期末成绩"])

    # 如果不是“规范表格”（例如报表格式，列名全是 Unnamed），走报表解析
    if not col_name:
        return read_excel_grades_report(
            excel_path=excel_path,
            sheet=resolved_sheet,
            default_class=default_class,
            default_course=default_course,
        )

    rows: List[GradeRow] = []
    for i, r in df.iterrows():
        name = str(r.get(col_name, "")).strip()
        if not name or name.lower() == "nan":
            continue

        class_name = (
            str(r.get(col_class)).strip() if col_class and not pd.isna(r.get(col_class)) else None
        )
        course = (
            str(r.get(col_course)).strip() if col_course and not pd.isna(r.get(col_course)) else None
        )
        usual = _to_int(r.get(col_usual)) if col_usual else None
        exam = _to_int(r.get(col_exam)) if col_exam else None

        if not class_name:
            class_name = default_class
        if not course:
            course = default_course

        final = None
        if usual is not None or exam is not None:
            final = (usual or 0) + (exam or 0)

        rows.append(
            GradeRow(
                name=name,
                class_name=class_name,
                course=course,
                usual=usual,
                exam=exam,
                final=final,
                source_row=int(i) + 2,  # +2：Excel 通常 1 行表头
            )
        )

    meta = {
        "excel": str(excel_path),
        "sheet": resolved_sheet,
        "mode": "table",
        "detected_columns": {
            "name": col_name,
            "class": col_class,
            "course": col_course,
            "usual": col_usual,
            "exam": col_exam,
        },
        "count": len(rows),
    }
    return rows, meta


def main() -> int:
    p = argparse.ArgumentParser(description="从 Excel 成绩单导出 grades.json（给自动化脚本使用）")
    p.add_argument("--excel", required=True, help="Excel 路径，例如 data.xlsx")
    p.add_argument("--sheet", default=None, help="工作表名称（不填则读取第一个）")
    p.add_argument("--out", required=True, help="输出 JSON 路径，例如 automation/grades.json")
    p.add_argument("--default-class", default=None, help="当 Excel 没有班级列时使用")
    p.add_argument("--default-course", default=None, help="当 Excel 没有课程列时使用")
    args = p.parse_args()

    excel_path = Path(args.excel).expanduser().resolve()
    out_path = Path(args.out).expanduser().resolve()
    out_path.parent.mkdir(parents=True, exist_ok=True)

    rows, meta = read_excel_grades(
        excel_path=excel_path,
        sheet=args.sheet,
        default_class=args.default_class,
        default_course=args.default_course,
    )

    payload = {
        "meta": meta,
        "grades": [asdict(x) for x in rows],
    }

    out_path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"已导出：{out_path}（{len(rows)} 条）")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

