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
    Reads the grey-cell mapping Excel.
    Grey/filled cells are treated as parent groups.
    Non-filled cells below each group are treated as mapped accounts.
    """
    wb = load_workbook(mapping_file, data_only=True)
    items: List[MappingItem] = []

    for ws in wb.worksheets:
        max_row = ws.max_row
        max_col = ws.max_column

        for col in range(1, max_col + 1):
            current_group: Optional[str] = None

            for row in range(1, max_row + 1):
                cell = ws.cell(row=row, column=col)
                value = str(cell.value or "").strip()

                if not value:
                    continue

                if is_grey_or_filled_cell(cell):
                    current_group = value
                    continue

                if current_group:
                    items.append(
                        MappingItem(
                            group_name=current_group,
                            account_name=value,
                            sheet_name=ws.title,
                            cell=cell.coordinate,
                        )
                    )

    return items


def read_trial_balance(trial_balance_file) -> pd.DataFrame:
    """
    Reads trial balance and tries to detect:
    - account name column
    - current year amount column
    - previous year amount column
    """
    raw = pd.read_excel(trial_balance_file, header=None)

    # Drop fully empty rows/columns
    raw = raw.dropna(how="all").dropna(axis=1, how="all")
    raw = raw.reset_index(drop=True)

    rows = []

    for _, row in raw.iterrows():
        values = list(row.values)

        text_values = [
            str(v).strip()
            for v in values
            if str(v).strip() and str(v).strip().lower() != "nan"
        ]

        if not text_values:
            continue

        # Account name = longest text value in row
        possible_names = [
            v for v in text_values
            if not re.fullmatch(r"[-+]?\(?[\d,]+(\.\d+)?\)?", v)
        ]

        if not possible_names:
            continue

        account_name = max(possible_names, key=len)

        numeric_values = [
            clean_number(v)
            for v in values
            if clean_number(v) != 0
        ]

        current_amount = numeric_values[0] if len(numeric_values) >= 1 else 0.0
        previous_amount = numeric_values[1] if len(numeric_values) >= 2 else 0.0

        rows.append(
            {
                "account_name": account_name,
                "account_name_clean": clean_text(account_name),
                "current_amount": current_amount,
                "previous_amount": previous_amount,
            }
        )

    return pd.DataFrame(rows)


def find_best_match(account_name: str, trial_balance_df: pd.DataFrame) -> Optional[pd.Series]:
    target = clean_text(account_name)

    if not target:
        return None

    # 1. Exact normalized match
    exact = trial_balance_df[trial_balance_df["account_name_clean"] == target]
    if not exact.empty:
        return exact.iloc[0]

    # 2. Contains match
    contains = trial_balance_df[
        trial_balance_df["account_name_clean"].apply(
            lambda x: target in x or x in target
        )
    ]

    if not contains.empty:
        return contains.iloc[0]

    # 3. Word overlap match
    target_words = set(target.split())

    best_row = None
    best_score = 0

    for _, row in trial_balance_df.iterrows():
        words = set(str(row["account_name_clean"]).split())
        if not words:
            continue

        score = len(target_words & words)

        if score > best_score:
            best_score = score
            best_row = row

    if best_score >= 2:
        return best_row

    return None


def build_mapping_preview(mapping_file, trial_balance_file) -> Tuple[pd.DataFrame, pd.DataFrame]:
    mapping_items = read_mapping_excel(mapping_file)
    trial_balance_df = read_trial_balance(trial_balance_file)

    matched_rows: List[Dict[str, object]] = []

    for item in mapping_items:
        match = find_best_match(item.account_name, trial_balance_df)

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
            matched_rows.append(
                {
                    "Group": item.group_name,
                    "Mapping account": item.account_name,
                    "Matched trial balance account": match["account_name"],
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
