from __future__ import annotations

import re
from dataclasses import dataclass
from typing import Dict, List, Optional, Tuple

import pandas as pd
from openpyxl import load_workbook


@dataclass
class MappingItem:
    group_name: str
    account_name: str
    sheet_name: str
    cell: str


@dataclass
class MatchedAccount:
    group_name: str
    mapping_account_name: str
    trial_balance_account_name: str
    current_amount: float
    previous_amount: float
    match_status: str


def clean_text(value) -> str:
    text = str(value or "").strip()
    text = re.sub(r"\s+", " ", text)

    replacements = {
        "أ": "ا",
        "إ": "ا",
        "آ": "ا",
        "ة": "ه",
        "ى": "ي",
        "ـ": "",
    }

    for old, new in replacements.items():
        text = text.replace(old, new)

    return text.strip()


def clean_number(value) -> float:
    if value is None:
        return 0.0

    text = str(value).strip()

    if not text or text.lower() == "nan":
        return 0.0

    # Arabic/English accounting negative format: (123)
    negative = text.startswith("(") and text.endswith(")")

    text = (
        text.replace(",", "")
        .replace("(", "")
        .replace(")", "")
        .replace("د.أ", "")
        .replace("دينار", "")
        .strip()
    )

    try:
        number = float(text)
    except ValueError:
        return 0.0

    return -number if negative else number


def is_grey_or_filled_cell(cell) -> bool:
    """
    Detects parent/category cells in the mapping Excel.
    Usually these are grey/dark filled cells.
    """
    fill = cell.fill

    if not fill or fill.fill_type is None:
        return False

    color = fill.fgColor.rgb or fill.fgColor.indexed

    if not color:
        return False

    # If it has any solid fill, treat it as a parent/category cell.
    # This is safer than trying to guess exact grey color.
    return True


def read_mapping_excel(mapping_file) -> List[MappingItem]:
    """
    Reads the mapping Excel correctly:
    - Column A = code like A1, A1.1, L3.1
    - Column B = Arabic account/group name
    - Rows without dot like A1, A2, L3 are parent groups
    - Rows with dot like A1.1, A1.2 are detailed accounts
    """
    wb = load_workbook(mapping_file, data_only=True)
    ws = wb.active

    items: List[MappingItem] = []
    current_group: Optional[str] = None

    for row in range(2, ws.max_row + 1):
        code = str(ws.cell(row=row, column=1).value or "").strip()
        desc = str(ws.cell(row=row, column=2).value or "").strip()

        if not code or not desc:
            continue

        if code in ["A", "L"]:
            current_group = None
            continue

        if desc.replace("*", "").strip() == "":
            continue

        # Parent group rows: A1, A2, A6, L1, L3...
        if "." not in code:
            current_group = desc
            continue

        # Detail rows: A1.1, A1.2, L3.1...
        if current_group:
            items.append(
                MappingItem(
                    group_name=current_group,
                    account_name=desc,
                    sheet_name=ws.title,
                    cell=ws.cell(row=row, column=2).coordinate,
                )
            )

    return items

def read_trial_balance(trial_balance_file) -> pd.DataFrame:
    """
    Reads the Laurel trial balance format.

    Expected columns:
    - المجموعة = account group used for matching
    - النهائي = current year amount
    - column before المجموعة = previous year amount
    - إسم الحساب = detailed account name
    - رقم الحساب = account number, ignored as amount
    """
    raw = pd.read_excel(trial_balance_file, header=None)

    header_row_index = None
    group_col = None
    current_col = None
    detail_col = None
    account_no_col = None

    for r in range(min(20, len(raw))):
        row_values = [clean_text(v) for v in raw.iloc[r].values]

        for c, value in enumerate(row_values):
            if "المجموعه" in value or "المجموعة" in str(raw.iloc[r, c]):
                header_row_index = r
                group_col = c

            if "النهائي" in value:
                current_col = c

            if "اسم الحساب" in value or "إسم الحساب" in str(raw.iloc[r, c]):
                detail_col = c

            if "رقم الحساب" in value:
                account_no_col = c

    if header_row_index is None or group_col is None or current_col is None:
        return pd.DataFrame(
            columns=[
                "account_name",
                "account_name_clean",
                "detail_account_name",
                "current_amount",
                "previous_amount",
            ]
        )

    previous_col = group_col - 1

    rows = []

    for r in range(header_row_index + 1, len(raw)):
        account_group = str(raw.iloc[r, group_col] or "").strip()

        if not account_group or account_group.lower() == "nan":
            continue

        detail_name = ""
        if detail_col is not None:
            detail_name = str(raw.iloc[r, detail_col] or "").strip()

        current_amount = clean_number(raw.iloc[r, current_col])
        previous_amount = clean_number(raw.iloc[r, previous_col])

        rows.append(
            {
                "account_name": account_group,
                "account_name_clean": clean_text(account_group),
                "detail_account_name": detail_name,
                "current_amount": current_amount,
                "previous_amount": previous_amount,
            }
        )

    return pd.DataFrame(rows)


def find_best_match(
    account_name: str,
    trial_balance_df: pd.DataFrame,
    used_indexes: Optional[set] = None,
) -> Optional[pd.Series]:
    used_indexes = used_indexes or set()
    target = clean_text(account_name)

    if not target or trial_balance_df.empty:
        return None

    # Add clean detail column if missing
    if "detail_account_name_clean" not in trial_balance_df.columns:
        trial_balance_df["detail_account_name_clean"] = trial_balance_df[
            "detail_account_name"
        ].apply(clean_text)

    # 1. Exact match with detailed account name
    exact_detail = trial_balance_df[
        (~trial_balance_df.index.isin(used_indexes))
        & (trial_balance_df["detail_account_name_clean"] == target)
    ]

    if not exact_detail.empty:
        return exact_detail.iloc[0]

    # 2. Exact match with group/account name
    exact_group = trial_balance_df[
        (~trial_balance_df.index.isin(used_indexes))
        & (trial_balance_df["account_name_clean"] == target)
    ]

    if not exact_group.empty:
        return exact_group.iloc[0]

    # 3. Safe contains match on detailed account name only
    for idx, row in trial_balance_df.iterrows():
        if idx in used_indexes:
            continue

        detail = str(row.get("detail_account_name_clean", ""))

        if len(target) >= 6 and (target in detail or detail in target):
            return row

    # 4. Strong word overlap on detailed account name
    target_words = {w for w in target.split() if len(w) > 2}

    best_row = None
    best_score = 0.0

    for idx, row in trial_balance_df.iterrows():
        if idx in used_indexes:
            continue

        detail_words = {
            w for w in str(row.get("detail_account_name_clean", "")).split()
            if len(w) > 2
        }

        if not target_words or not detail_words:
            continue

        overlap = len(target_words & detail_words)
        score = overlap / max(len(target_words), len(detail_words))

        if score > best_score:
            best_score = score
            best_row = row

    if best_score >= 0.70:
        return best_row

    return None


def build_mapping_preview(mapping_file, trial_balance_file) -> Tuple[pd.DataFrame, pd.DataFrame]:
    mapping_items = read_mapping_excel(mapping_file)
    trial_balance_df = read_trial_balance(trial_balance_file)

    matched_rows: List[Dict[str, object]] = []
    used_trial_indexes = set()

    for item in mapping_items:
        match = find_best_match(
            item.account_name,
            trial_balance_df,
            used_indexes=used_trial_indexes,
        )

        if match is None:
            matched_rows.append(
                {
                    "Group": item.group_name,
                    "Mapping account": item.account_name,
                    "Matched trial balance account": "",
                    "Current amount": 0.0,
                    "Previous amount": 0.0,
                    "Status": "Not matched",
                    "Mapping cell": item.cell,
                }
            )
        else:
            used_trial_indexes.add(match.name)

            matched_rows.append(
                {
                    "Group": item.group_name,
                    "Mapping account": item.account_name,
                    "Matched trial balance account": match.get(
                        "detail_account_name",
                        match["account_name"],
                    ),
                    "Current amount": match["current_amount"],
                    "Previous amount": match["previous_amount"],
                    "Status": "Matched",
                    "Mapping cell": item.cell,
                }
            )

    preview_df = pd.DataFrame(matched_rows)

    if preview_df.empty:
        summary_df = pd.DataFrame(
            columns=["Group", "Current total", "Previous total", "Matched accounts"]
        )
    else:
        summary_df = (
            preview_df[preview_df["Status"] == "Matched"]
            .groupby("Group", as_index=False)
            .agg(
                {
                    "Current amount": "sum",
                    "Previous amount": "sum",
                    "Mapping account": "count",
                }
            )
            .rename(
                columns={
                    "Current amount": "Current total",
                    "Previous amount": "Previous total",
                    "Mapping account": "Matched accounts",
                }
            )
        )

    return preview_df, summary_df
