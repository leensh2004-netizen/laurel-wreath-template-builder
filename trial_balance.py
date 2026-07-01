from __future__ import annotations

import io
import re
from dataclasses import dataclass, field
from pathlib import Path
from typing import Dict, Iterable, List, Optional, Tuple

import pandas as pd


@dataclass
class NoteMap:
    section: str
    table_index: int
    row_index: int
    label: str
    groups: List[str] = field(default_factory=list)
    total_of: Optional[List[str]] = None
    display: str = "normal"  # normal, deduction
    default_include: bool = True


# How the company's sample trial balance is arranged:
# - Excel column "المجموعة" is the source for matching note rows.
# - Current year amount normally comes from "النهائي".
# - Previous year/comparative amount is the numeric column immediately BEFORE "المجموعة".
# - Liabilities/revenues often appear as negative in the trial balance, but the notes usually display positive values.
# - Deductions such as sales discounts are displayed in parentheses.
# - Matching is intentionally exact after Arabic normalization. This prevents accidental matches such as
#   "ذمم موظفين" being mixed with "ذمم موظفين - دائن".
NOTE_MAP: List[NoteMap] = [
    # (4) Cash and cash equivalents - table 4
    NoteMap("(4) النقد وما في حكمه", 4, 3, "نقد في الصندوق", ["نقد في الصندوق"]),
    NoteMap("(4) النقد وما في حكمه", 4, 4, "نقد لدى البنوك", ["نقد لدى البنوك"]),
    NoteMap("(4) النقد وما في حكمه", 4, 5, "سلفة نثرية", ["سلف النثرية", "سلفة نثرية"]),
    NoteMap("(4) النقد وما في حكمه", 4, 6, "المجموع", total_of=["نقد في الصندوق", "نقد لدى البنوك", "سلفة نثرية"]),

    # (5) Trade receivables - table 5
    NoteMap("(5) الذمم المدينة", 5, 3, "ذمم شركات التأمين", ["ذمم شركات التامين", "ذمم شركات التأمين"]),
    NoteMap("(5) الذمم المدينة", 5, 4, "ذمم عملاء", [
        "ذمم عملاء", "ذمم  عملاء", "ذمم اولياء امور", "ذمم طلاب محولة الى شركة مكين", "ذمم عملاء محولة الى المحامي"
    ]),
    NoteMap("(5) الذمم المدينة", 5, 5, "المجموع", total_of=["ذمم شركات التأمين", "ذمم عملاء"]),

    # (6) Inventory - table 6
    NoteMap("(6) المخزون", 6, 3, "ادوية طبية ومستلزمات تجميلية", ["بضاعة جاهزة", "مخزون", "ادوية طبية ومستلزمات تجميلية"]),
    NoteMap("(6) المخزون", 6, 4, "المجموع", total_of=["ادوية طبية ومستلزمات تجميلية"]),

    # (7) Other debit balances - table 7
    NoteMap("(7) أرصدة مدينة أخرى", 7, 3, "مصاريف ايجار مدفوعة مقدما", ["مصاريف مدفوعة مقدما"]),
    NoteMap("(7) أرصدة مدينة أخرى", 7, 4, "تأمينات مستردة", ["تأمينات مستردة", "تامينات مستردة"]),
    NoteMap("(7) أرصدة مدينة أخرى", 7, 5, "ذمم الموظفين", ["ذمم موظفين"]),
    NoteMap("(7) أرصدة مدينة أخرى", 7, 6, "المجموع", total_of=["مصاريف ايجار مدفوعة مقدما", "تأمينات مستردة", "ذمم الموظفين"]),

    # (10) Payables - table 16
    NoteMap("(10) الذمم الدائنة", 16, 3, "مستودعات طبية", ["ذمم موردين", "ذمم موردين سلع", "ذمم دائنة موردين", "مستودعات طبية"]),
    NoteMap("(10) الذمم الدائنة", 16, 4, "موردين آخرين - خدمات", ["موردين آخرين - خدمات"]),
    NoteMap("(10) الذمم الدائنة", 16, 5, "المجموع", total_of=["مستودعات طبية", "موردين آخرين - خدمات"]),

    # (11) Bank overdraft - table 17
    NoteMap("(11) بنك دائن", 17, 3, "البنك العربي – بطاقة ائتمانية", ["بنك دائن", "بطاقة ائتمانية"]),
    NoteMap("(11) بنك دائن", 17, 4, "المجموع", total_of=["البنك العربي – بطاقة ائتمانية"]),

    # (12) Accrued expenses - table 18
    NoteMap("(12) المصاريف المستحقة", 18, 3, "اتعاب تدقيق مستحقة", ["مصاريف مستحقة - اتعاب مهنية", "اتعاب تدقيق مستحقة"]),
    NoteMap("(12) المصاريف المستحقة", 18, 4, "مصاريف اخرى مستحقة", ["مصاريف مستحقة - رواتب", "مصاريف اخرى مستحقة"]),
    NoteMap("(12) المصاريف المستحقة", 18, 5, "المجموع", total_of=["اتعاب تدقيق مستحقة", "مصاريف اخرى مستحقة"]),

    # (13) Related party payable - table 19
    # This note often has one row per named shareholder/partner. The trial balance may contain
    # many account-name rows that cannot safely be mapped to the fixed partner rows automatically.
    # Therefore it is shown in the editable table but not applied by default.
    NoteMap("(13) مطلوب لأطراف ذات علاقة", 19, 3, "مطلوب من أطراف ذات علاقة", [
        "مطلوب  من اطراف ذات علاقه", "مطلوب من اطراف ذات علاقه", "مطلوب لاطراف ذات علاقة", "ذمم المساهمين"
    ], default_include=False),
    NoteMap("(13) مطلوب لأطراف ذات علاقة", 19, 6, "المجموع", total_of=["مطلوب من أطراف ذات علاقة"], default_include=False),

    # (16) Other credit balances - table 22
    NoteMap("(16) أرصدة دائنة أخرى", 22, 3, "ذمم موظفين", ["ذمم موظفين - دائن"]),
    NoteMap("(16) أرصدة دائنة أخرى", 22, 4, "امانات ضريبة الدخل %5", ["امانات ضريبة الدخل", "امانات ضريبة الدخل - دائن"]),
    NoteMap("(16) أرصدة دائنة أخرى", 22, 5, "امانات ضريبة المبيعات", ["امانات ضريبة المبيعات"]),
    NoteMap("(16) أرصدة دائنة أخرى", 22, 6, "ضريبة دخل موظفين", ["ضريبة دخل موظفين"]),
    NoteMap("(16) أرصدة دائنة أخرى", 22, 7, "امانات الضمان الاجتماعي", ["امانات الضمان الاجتماعي"]),
    NoteMap("(16) أرصدة دائنة أخرى", 22, 8, "المجموع", total_of=["ذمم موظفين", "امانات ضريبة الدخل %5", "امانات ضريبة المبيعات", "ضريبة دخل موظفين", "امانات الضمان الاجتماعي"]),

    # (20) Sales - table 24
    NoteMap("(20) صافي المبيعات", 24, 3, "المبيعات", [
        "مبيعات كتب مدرسية", "مبيعات زي مدرسي", "مبيعات مقصف", "رسوم الاقساط", "رسوم التسجيل",
        "ايراد المواصلات", "رسوم التوجيهي الحكومية", "ايراد دخوليات مسبح", "ايرادات تأجير",
        "ايرادات نوادي", "ايراد اصدار شهادات", "ايراد دورات سباحة", "ايرادات نشاطات اخرى", "ايرادات الماراثون",
        "المبيعات"
    ]),
    NoteMap("(20) صافي المبيعات", 24, 4, "خصم المبيعات", ["خصم المبيعات"], display="deduction"),
    NoteMap("(20) صافي المبيعات", 24, 5, "الخصم التعاقدي", ["الخصم التعاقدي"], display="deduction"),
    NoteMap("(20) صافي المبيعات", 24, 6, "مردودات المبيعات", ["مردودات المبيعات"], display="deduction"),
    NoteMap("(20) صافي المبيعات", 24, 7, "صافي المبيعات", total_of=["المبيعات", "خصم المبيعات", "الخصم التعاقدي", "مردودات المبيعات"]),

    # (21) Administrative and general expenses - table 25
    NoteMap("(21) المصاريف الإدارية والعمومية", 25, 3, "رواتب واجور", ["الرواتب و الاجور", "مصروف الرواتب - تشغيلية"]),
    NoteMap("(21) المصاريف الإدارية والعمومية", 25, 4, "مساهمة الشركة في الضمان الاجتماعي", ["مصروف الضمان الاجتماعي", "الضمان الاجتماعي - تكاليف"]),
    NoteMap("(21) المصاريف الإدارية والعمومية", 25, 5, "مصاريف العمل الاضافي", ["مصاريف العمل الاضافي"]),
    NoteMap("(21) المصاريف الإدارية والعمومية", 25, 6, "مكافات", ["مكافئة الموظفين", "مكافات"]),
    NoteMap("(21) المصاريف الإدارية والعمومية", 25, 7, "بدل اجازات سنوية", ["بدل اجازات سنوية"]),
    NoteMap("(21) المصاريف الإدارية والعمومية", 25, 8, "دورات تدريبية", ["مصاريف التدريب", "دورات تدريبية"]),
    NoteMap("(21) المصاريف الإدارية والعمومية", 25, 9, "مصاريف ايجارات", ["ايجارات", "مصاريف ايجارات"]),
    NoteMap("(21) المصاريف الإدارية والعمومية", 25, 10, "مصاريف بريد وهاتف", ["مصاريف بريد وهاتف"]),
    NoteMap("(21) المصاريف الإدارية والعمومية", 25, 11, "مصاريف كهربا ومياه", ["مصاريف كهرباء و تدفئة ومياه", "مصاريف كهربا ومياه"]),
    NoteMap("(21) المصاريف الإدارية والعمومية", 25, 12, "مصاريف دعاية واعلان", ["مصاريف دعاية و اعلان", "مصاريف دعاية واعلان"]),
    NoteMap("(21) المصاريف الإدارية والعمومية", 25, 13, "مصاريف سيارات", ["تكاليف الباصات", "مصاريف سيارات"]),
    NoteMap("(21) المصاريف الإدارية والعمومية", 25, 14, "مصاريف صيانة واصلاح", ["مصاريف صيانة و أصلاح", "مصاريف صيانة واصلاح"]),
    NoteMap("(21) المصاريف الإدارية والعمومية", 25, 15, "مصاريف حكومية", ["مصاريف حكومية"]),
    NoteMap("(21) المصاريف الإدارية والعمومية", 25, 16, "بدل رخصة مزاولة مهنة", ["بدل رخصة مزاولة مهنة"]),
    NoteMap("(21) المصاريف الإدارية والعمومية", 25, 17, "مصاريف مخصص قضايا", ["مصاريف قضايا", "مصاريف رسوم قضايا"]),
    NoteMap("(21) المصاريف الإدارية والعمومية", 25, 18, "مصاريف تنقلات", ["مصاريف سفر و تنقلات", "مصاريف تنقلات"]),
    NoteMap("(21) المصاريف الإدارية والعمومية", 25, 19, "مصاريف بنكية", ["عمولات مدفوعة 1", "مصاريف بنكية"]),
    NoteMap("(21) المصاريف الإدارية والعمومية", 25, 20, "مصاريف الاستهلاك", ["مصروف الاستهلاك"]),
    NoteMap("(21) المصاريف الإدارية والعمومية", 25, 21, "مصاريف اتعاب قانونية", ["مصروف اتعاب تحصيل", "مصاريف اتعاب قانونية"]),
    NoteMap("(21) المصاريف الإدارية والعمومية", 25, 22, "مصاريف اتعاب تدقيق", ["مصاريف اتعاب تدقيق", "اتعاب تدقيق"]),
    NoteMap("(21) المصاريف الإدارية والعمومية", 25, 23, "تبرعات", ["تبرعات"]),
    NoteMap("(21) المصاريف الإدارية والعمومية", 25, 24, "مصاريف قرطاسية ومطبوعات", ["مصاريف قرطاسية و مطبوعات", "مصاريف قرطاسية ومطبوعات"]),
    NoteMap("(21) المصاريف الإدارية والعمومية", 25, 25, "مصاريف ضيافة ونظافة", ["ضيافة ونظافة", "مصاريف ضيافة ونظافة"]),
    NoteMap("(21) المصاريف الإدارية والعمومية", 25, 26, "مصاريف غرامات", ["مصاريف غرامات"]),
    NoteMap("(21) المصاريف الإدارية والعمومية", 25, 27, "اتعاب استشارات", ["اتعاب مهنية", "اتعاب استشارات"]),
    NoteMap("(21) المصاريف الإدارية والعمومية", 25, 28, "مصاريف انظمة وحماية", ["مصاريف انظمة وحماية"]),
    NoteMap("(21) المصاريف الإدارية والعمومية", 25, 29, "رسوم واشتراكات", ["رسوم واشتراكات"]),
    NoteMap("(21) المصاريف الإدارية والعمومية", 25, 30, "مصاريف ضريبة معارف", ["مصاريف ضريبة معارف"]),
    NoteMap("(21) المصاريف الإدارية والعمومية", 25, 31, "مصاريف محروقات", ["مصاريف محروقات"]),
    NoteMap("(21) المصاريف الإدارية والعمومية", 25, 32, "مصاريف ندوات ومؤتمرات", ["مصاريف ندوات ومؤتمرات"]),
    NoteMap("(21) المصاريف الإدارية والعمومية", 25, 33, "مصاريف تصميم مواقع", ["مصاريف مواقع الكترونية", "مصاريف تصميم مواقع"]),
    NoteMap("(21) المصاريف الإدارية والعمومية", 25, 34, "مصاريف اخرى", ["مصاريف اخرى", "مصاريف تشغيلية", "مواد مستهلكة", "مصاريف (ايرادات) اخرى"]),
    NoteMap("(21) المصاريف الإدارية والعمومية", 25, 35, "المجموع", total_of=[
        "رواتب واجور", "مساهمة الشركة في الضمان الاجتماعي", "مصاريف العمل الاضافي", "مكافات", "بدل اجازات سنوية",
        "دورات تدريبية", "مصاريف ايجارات", "مصاريف بريد وهاتف", "مصاريف كهربا ومياه", "مصاريف دعاية واعلان",
        "مصاريف سيارات", "مصاريف صيانة واصلاح", "مصاريف حكومية", "بدل رخصة مزاولة مهنة", "مصاريف مخصص قضايا",
        "مصاريف تنقلات", "مصاريف بنكية", "مصاريف الاستهلاك", "مصاريف اتعاب قانونية", "مصاريف اتعاب تدقيق",
        "تبرعات", "مصاريف قرطاسية ومطبوعات", "مصاريف ضيافة ونظافة", "مصاريف غرامات", "اتعاب استشارات",
        "مصاريف انظمة وحماية", "رسوم واشتراكات", "مصاريف ضريبة معارف", "مصاريف محروقات", "مصاريف ندوات ومؤتمرات",
        "مصاريف تصميم مواقع", "مصاريف اخرى"
    ]),
]


def _normalize_arabic(text: object) -> str:
    text = "" if text is None else str(text)
    text = text.replace("\u00a0", " ")
    text = re.sub(r"[إأآا]", "ا", text)
    text = text.replace("ى", "ي").replace("ة", "ه")
    # Normalize punctuation used in exports.
    text = text.replace("–", "-").replace("—", "-")
    text = re.sub(r"\s+", " ", text)
    return text.strip()


def _to_number(value: object) -> float:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return 0.0
    if isinstance(value, (int, float)):
        return float(value)
    text = str(value).strip().replace(",", "")
    if text in {"", "-", "nan", "None"}:
        return 0.0
    neg = text.startswith("(") and text.endswith(")")
    text = text.strip("()")
    try:
        num = float(text)
        return -num if neg else num
    except ValueError:
        return 0.0


def _format_number(num: float, blank_zero: bool = False) -> str:
    if blank_zero and abs(num) < 0.0001:
        return "-"
    if abs(num - round(num)) < 0.005:
        return f"{round(num):,}"
    return f"{num:,.3f}".rstrip("0").rstrip(".")


def _format_amount(value: object, blank_zero: bool = False, display: str = "normal") -> str:
    num = _to_number(value)
    if blank_zero and abs(num) < 0.0001:
        return "-"
    shown = abs(num)
    formatted = _format_number(shown, blank_zero=False)
    if display == "deduction" and shown > 0:
        return f"({formatted})"
    return formatted


def _find_header_row(df: pd.DataFrame) -> int:
    for idx, row in df.iterrows():
        vals = [_normalize_arabic(v) for v in row.tolist()]
        joined = " ".join(vals)
        if "المجموعه" in joined and ("النهائي" in joined or "الرصيد" in joined) and "الحساب" in joined:
            return int(idx)
    return 1


def _find_col(header_values: List[str], contains: Iterable[str]) -> Optional[int]:
    targets = [_normalize_arabic(x) for x in contains]
    for i, val in enumerate(header_values):
        nval = _normalize_arabic(val)
        if any(t in nval for t in targets):
            return i
    return None


def _find_previous_year_col(header: List[str], group_col: int) -> Optional[int]:
    """The company's exported TB stores the previous/comparative amount just before المجموعة."""
    for col in range(group_col - 1, -1, -1):
        # Pick the nearest column to the left that is not part of the known right-side headers.
        if _normalize_arabic(header[col]) not in {"", "المجموعه", "النهائي", "قيود التعديل", "الرصيد", "اسم الحساب", "رقم الحساب"}:
            return col
    # In the sample, the header cell above the comparative column is blank, so use group_col - 1.
    return group_col - 1 if group_col > 0 else None


def read_trial_balance(uploaded_file_or_path, current_amount_source: str = "final") -> pd.DataFrame:
    """Read common Arabic trial-balance exports (.xls or .xlsx) into a normalized DataFrame.

    current_amount_source:
        - "final": use column النهائي for the current year.
        - "balance": use column الرصيد for the current year.

    Previous year is taken from the comparative numeric column immediately before "المجموعة".
    """
    if isinstance(uploaded_file_or_path, (str, Path)):
        raw = uploaded_file_or_path
    else:
        raw = io.BytesIO(uploaded_file_or_path.read())

    df = pd.read_excel(raw, sheet_name=0, header=None)
    header_idx = _find_header_row(df)
    header = [_normalize_arabic(v) for v in df.iloc[header_idx].tolist()]

    group_col = _find_col(header, ["المجموعه", "المجموعة"])
    final_col = _find_col(header, ["النهائي"])
    balance_col = _find_col(header, ["الرصيد"])
    account_name_col = _find_col(header, ["اسم الحساب", "إسم الحساب"])
    account_no_col = _find_col(header, ["رقم الحساب"])

    if group_col is None:
        raise ValueError("Could not find the 'المجموعة' column in the uploaded trial balance.")
    if final_col is None and balance_col is None:
        raise ValueError("Could not find either 'النهائي' or 'الرصيد' columns in the uploaded trial balance.")

    current_col = final_col if current_amount_source == "final" and final_col is not None else balance_col
    previous_col = _find_previous_year_col(header, group_col)

    data = df.iloc[header_idx + 1:].copy()
    out = pd.DataFrame()
    out["group"] = data.iloc[:, group_col].map(_normalize_arabic)
    out["current"] = data.iloc[:, current_col].map(_to_number)
    out["previous"] = data.iloc[:, previous_col].map(_to_number) if previous_col is not None else 0.0
    out["account_name"] = data.iloc[:, account_name_col].astype(str) if account_name_col is not None else ""
    out["account_no"] = data.iloc[:, account_no_col].astype(str) if account_no_col is not None else ""
    out = out[out["group"].astype(str).str.len() > 0]
    return out


def group_totals(tb: pd.DataFrame) -> Dict[str, Tuple[float, float]]:
    grouped = tb.groupby("group", dropna=True)[["current", "previous"]].sum().reset_index()
    return {row["group"]: (float(row["current"]), float(row["previous"])) for _, row in grouped.iterrows()}


def _sum_groups(totals: Dict[str, Tuple[float, float]], groups: List[str]) -> Tuple[float, float, List[str]]:
    current = 0.0
    previous = 0.0
    matched_names: List[str] = []
    targets = {_normalize_arabic(g) for g in groups if _normalize_arabic(g)}
    for group_name, (cur, prev) in totals.items():
        ng = _normalize_arabic(group_name)
        if ng in targets:
            current += cur
            previous += prev
            matched_names.append(group_name)
    return current, previous, matched_names


def build_note_rows(tb: Optional[pd.DataFrame] = None, round_to_whole_dinars: bool = True) -> pd.DataFrame:
    totals = group_totals(tb) if tb is not None and not tb.empty else {}
    rows = []
    values_by_label: Dict[str, Tuple[float, float]] = {}
    matched_by_label: Dict[str, List[str]] = {}

    for m in NOTE_MAP:
        if m.total_of:
            cur = sum(values_by_label.get(label, (0.0, 0.0))[0] for label in m.total_of)
            prev = sum(values_by_label.get(label, (0.0, 0.0))[1] for label in m.total_of)
            matched = []
            for label in m.total_of:
                matched.extend(matched_by_label.get(label, []))
        else:
            cur, prev, matched = _sum_groups(totals, m.groups)

        values_by_label[m.label] = (cur, prev)
        matched_by_label[m.label] = matched
        cur_to_show = round(cur) if round_to_whole_dinars else cur
        prev_to_show = round(prev) if round_to_whole_dinars else prev
        rows.append({
            "include": m.default_include,
            "section": m.section,
            "table_index": m.table_index,
            "row_index": m.row_index,
            "label": m.label,
            "current_year": _format_amount(cur_to_show, blank_zero=True, display=m.display),
            "previous_year": _format_amount(prev_to_show, blank_zero=True, display=m.display),
            "matched_groups": ", ".join(dict.fromkeys(matched)) if matched else "",
        })
    return pd.DataFrame(rows)



# -----------------------------
# V10 movement-table helpers
# -----------------------------

@dataclass
class MovementCell:
    table_index: int
    row_index: int
    col_index: int
    value: str
    label: str = ""
    source: str = ""

# Word table column positions for the PPE movement tables.
# The template uses spacer columns, so useful amount columns are 2,4,6,...,16.
PPE_AMOUNT_COLS = {
    "اجهزة حاسوب": 2,
    "اثاث ومفروشات": 4,
    "ديكورات": 6,
    "اجهزة اتصالات ومعدات": 8,
    "سيارات": 10,
    "تحسينات مباني": 12,
    "مكيفات": 14,
    "المجموع": 16,
}

# Conservative group mapping for PPE cost and accumulated depreciation.
# If a category is not clearly present in the uploaded trial balance, V10 leaves
# that category unchanged instead of guessing and overwriting the Word template.
PPE_GROUPS = {
    "اجهزة حاسوب": {
        "cost": ["اجهزه حاسوب", "اجهزة حاسوب"],
        "dep": ["الاستهلاك المتراكم - اجهزه الحاسوب", "الاستهلاك المتراكم - اجهزة الحاسوب"],
    },
    "اثاث ومفروشات": {
        "cost": ["اثاث و مفروشات", "اثاث ومفروشات"],
        "dep": ["الاستهلاك المتراكم - اثاث و مفروشات", "الاستهلاك المتراكم - اثاث ومفروشات"],
    },
    "ديكورات": {
        "cost": ["ديكورات"],
        "dep": ["الاستهلاك المتراكم - ديكورات"],
    },
    "اجهزة اتصالات ومعدات": {
        "cost": ["اجهزه اتصالات و معدات", "اجهزة اتصالات ومعدات"],
        "dep": ["الاستهلاك المتراكم - اجهزه اتصالات و معدات", "الاستهلاك المتراكم - اجهزة اتصالات ومعدات"],
    },
    "سيارات": {
        "cost": ["سيارات"],
        "dep": ["الاستهلاك المتراكم - سيارات"],
    },
    "تحسينات مباني": {
        "cost": ["مباني", "تحسينات مباني", "تحسينات مباني - ديكورات"],
        "dep": ["الاستهلاك المتراكم - مباني", "الاستهلاك المتراكم - تحسينات مباني"],
    },
    "مكيفات": {
        "cost": ["مكيفات"],
        "dep": ["الاستهلاك المتراكم - مكيفات"],
    },
}


def _lookup_group_total(totals: Dict[str, Tuple[float, float]], names: List[str]) -> Tuple[float, float, List[str]]:
    return _sum_groups(totals, names)


def build_ppe_movement_cells(tb: Optional[pd.DataFrame], round_to_whole_dinars: bool = True) -> pd.DataFrame:
    """Build editable 2024 PPE movement-table cell suggestions.

    These are based on trial-balance ending balances:
    - Opening cost = previous/comparative cost balance
    - Additions = current cost - previous cost, if positive
    - Closing cost = current cost
    - Opening accumulated depreciation = abs(previous accumulated depreciation)
    - Depreciation for the year = abs(current accumulated dep.) - abs(previous accumulated dep.), if positive
    - Closing accumulated depreciation = abs(current accumulated depreciation)
    - Net book value = current cost - abs(current accumulated depreciation)

    V10 intentionally does not overwrite categories that are not clearly matched.
    """
    if tb is None or tb.empty:
        return pd.DataFrame(columns=["include", "table_index", "row_index", "col_index", "label", "value", "source"])

    totals = group_totals(tb)
    rows = []
    totals_acc = {
        "cost_opening": 0.0,
        "cost_additions": 0.0,
        "cost_closing": 0.0,
        "dep_opening": 0.0,
        "dep_year": 0.0,
        "dep_closing": 0.0,
        "nbv": 0.0,
    }

    def fmt(v):
        v = round(v) if round_to_whole_dinars else v
        return _format_amount(v, blank_zero=True)

    for label, maps in PPE_GROUPS.items():
        cost_cur, cost_prev, cost_matches = _lookup_group_total(totals, maps["cost"])
        dep_cur, dep_prev, dep_matches = _lookup_group_total(totals, maps["dep"])
        if not cost_matches and not dep_matches:
            continue
        col = PPE_AMOUNT_COLS[label]
        cost_opening = abs(cost_prev)
        cost_closing = abs(cost_cur)
        cost_additions = max(cost_closing - cost_opening, 0.0)
        dep_opening = abs(dep_prev)
        dep_closing = abs(dep_cur)
        dep_year = max(dep_closing - dep_opening, 0.0)
        nbv = max(cost_closing - dep_closing, 0.0)

        source = ", ".join(dict.fromkeys(cost_matches + dep_matches))
        cell_specs = [
            (8, 2, col, f"2024 الكلفة - {label} - رصيد 1 كانون الثاني", cost_opening),
            (8, 3, col, f"2024 الكلفة - {label} - إضافات", cost_additions),
            (8, 4, col, f"2024 الكلفة - {label} - رصيد 31 كانون الأول", cost_closing),
            (9, 3, col, f"2024 الاستهلاك المتراكم - {label} - رصيد 1 كانون الثاني", dep_opening),
            (9, 4, col, f"2024 الاستهلاك المتراكم - {label} - استهلاك السنة", dep_year),
            (9, 5, col, f"2024 الاستهلاك المتراكم - {label} - رصيد 31 كانون الأول", dep_closing),
            (9, 6, col, f"2024 صافي القيمة الدفترية - {label}", nbv),
        ]
        for table_i, row_i, col_i, desc, val in cell_specs:
            rows.append({
                "include": True,
                "table_index": table_i,
                "row_index": row_i,
                "col_index": col_i,
                "label": desc,
                "value": fmt(val),
                "source": source,
            })
        totals_acc["cost_opening"] += cost_opening
        totals_acc["cost_additions"] += cost_additions
        totals_acc["cost_closing"] += cost_closing
        totals_acc["dep_opening"] += dep_opening
        totals_acc["dep_year"] += dep_year
        totals_acc["dep_closing"] += dep_closing
        totals_acc["nbv"] += nbv

    if rows:
        total_col = PPE_AMOUNT_COLS["المجموع"]
        total_specs = [
            (8, 2, total_col, "2024 الكلفة - المجموع - رصيد 1 كانون الثاني", totals_acc["cost_opening"]),
            (8, 3, total_col, "2024 الكلفة - المجموع - إضافات", totals_acc["cost_additions"]),
            (8, 4, total_col, "2024 الكلفة - المجموع - رصيد 31 كانون الأول", totals_acc["cost_closing"]),
            (9, 3, total_col, "2024 الاستهلاك المتراكم - المجموع - رصيد 1 كانون الثاني", totals_acc["dep_opening"]),
            (9, 4, total_col, "2024 الاستهلاك المتراكم - المجموع - استهلاك السنة", totals_acc["dep_year"]),
            (9, 5, total_col, "2024 الاستهلاك المتراكم - المجموع - رصيد 31 كانون الأول", totals_acc["dep_closing"]),
            (9, 6, total_col, "2024 صافي القيمة الدفترية - المجموع", totals_acc["nbv"]),
        ]
        for table_i, row_i, col_i, desc, val in total_specs:
            rows.append({
                "include": True,
                "table_index": table_i,
                "row_index": row_i,
                "col_index": col_i,
                "label": desc,
                "value": fmt(val),
                "source": "calculated total",
            })
    return pd.DataFrame(rows)


def rows_to_cell_values(df: pd.DataFrame) -> Dict[str, str]:
    values: Dict[str, str] = {}
    if df is None or df.empty:
        return values
    for _, row in df.iterrows():
        include = row.get("include", True)
        if isinstance(include, str):
            include = include.strip().lower() not in {"false", "0", "no", "لا"}
        if not include:
            continue
        try:
            key = f"{int(row['table_index'])}:{int(row['row_index'])}:{int(row['col_index'])}"
        except Exception:
            continue
        values[key] = str(row.get("value", "") or "")
    return values


def rows_to_financial_values(df: pd.DataFrame) -> Dict[str, Dict[str, str]]:
    values: Dict[str, Dict[str, str]] = {}
    if df is None or df.empty:
        return values
    for _, row in df.iterrows():
        include = row.get("include", True)
        if isinstance(include, str):
            include = include.strip().lower() not in {"false", "0", "no", "لا"}
        if not include:
            continue
        key = f"{int(row['table_index'])}:{int(row['row_index'])}"
        values[key] = {
            "current": str(row.get("current_year", "") or ""),
            "previous": str(row.get("previous_year", "") or ""),
            "label": str(row.get("label", "") or ""),
            "section": str(row.get("section", "") or ""),
        }
    return values
