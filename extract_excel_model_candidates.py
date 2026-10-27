from __future__ import annotations

import math
import re
from dataclasses import dataclass
from datetime import date
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Sequence, Tuple

import xlwings as xw
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

# ------------------------------- INPUTS ------------------------------------ #
# Update these folders before running.
input_dir = "./input"
output_dir = "./output"


EMPIRICAL_HEADERS = [
    "model",
    "ticker",
    "model_period",
    "model_date",
    "method",
    "parameter_name",
    "parameter_value",
    "num_quarters_used",
    "last_quarter_used",
    "forecast_value",
    "actual_value",
    "forecast_max",
    "forecast_min",
    "range_width",
    "avg_penetration_pct",
    "quarterly_sales",
    "reported_sales",
    "growth_rate_pct",
    "sales_captured_in_db_pct",
    "source_file",
]

REGRESSION_HEADERS = [
    "model",
    "ticker",
    "model_period",
    "model_date",
    "method",
    "parameter_name",
    "parameter_value",
    "num_quarters_used",
    "forecast_value",
    "actual_value",
    "forecast_max",
    "forecast_min",
    "range_width",
    "intercept",
    "slope",
    "source_file",
]


MONTH_LOOKUP = {
    "jan": 1,
    "feb": 2,
    "mar": 3,
    "apr": 4,
    "may": 5,
    "jun": 6,
    "jul": 7,
    "aug": 8,
    "sep": 9,
    "oct": 10,
    "nov": 11,
    "dec": 12,
}

PERIOD_DAY_LOOKUP = {
    "early": 5,
    "mid": 15,
    "late": 25,
}


@dataclass
class SheetSnapshot:
    values: List[List[Any]]
    start_row: int
    start_col: int
    end_row: int
    end_col: int

    def get(self, row: int, col: int) -> Any:
        if row < self.start_row or row > self.end_row:
            return None
        if col < self.start_col or col > self.end_col:
            return None
        return self.values[row - self.start_row][col - self.start_col]


def normalize_text(value: Any) -> str:
    if value is None:
        return ""
    text = str(value).strip().lower()
    text = re.sub(r"[\s_/\\-]+", " ", text)
    text = re.sub(r"[^a-z0-9% ]+", "", text)
    return text.strip()


def normalize_for_key(value: Any) -> str:
    return normalize_text(value).replace("%", "pct")


def as_2d(values: Any) -> List[List[Any]]:
    if values is None:
        return []
    if isinstance(values, list):
        if not values:
            return []
        if isinstance(values[0], list):
            return values
        return [values]
    return [[values]]


def to_float(value: Any) -> Optional[float]:
    if value is None:
        return None
    if isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        if isinstance(value, float) and math.isnan(value):
            return None
        return float(value)
    text = str(value).strip()
    if not text:
        return None
    pct = "%" in text
    cleaned = text.replace("%", "").replace(",", "")
    try:
        parsed = float(cleaned)
    except ValueError:
        return None
    if pct:
        return parsed / 100.0
    return parsed


def ensure_pct_decimal(value: Any) -> Optional[float]:
    numeric = to_float(value)
    if numeric is None:
        return None
    if abs(numeric) > 1.5 and abs(numeric) <= 100:
        return numeric / 100.0
    return numeric


def snapshot_sheet(sheet: xw.Sheet) -> Optional[SheetSnapshot]:
    used = sheet.used_range
    values = as_2d(used.value)
    if not values:
        return None
    row_count = len(values)
    col_count = max(len(row) for row in values) if row_count else 0
    normalized: List[List[Any]] = []
    for row in values:
        padded = list(row) + [None] * (col_count - len(row))
        normalized.append(padded)
    start_row = used.row
    start_col = used.column
    end_row = start_row + row_count - 1
    end_col = start_col + col_count - 1
    return SheetSnapshot(
        values=normalized,
        start_row=start_row,
        start_col=start_col,
        end_row=end_row,
        end_col=end_col,
    )


def iter_snapshot_cells(snapshot: SheetSnapshot) -> Iterable[Tuple[int, int, Any]]:
    for r_offset, row in enumerate(snapshot.values):
        row_no = snapshot.start_row + r_offset
        for c_offset, value in enumerate(row):
            col_no = snapshot.start_col + c_offset
            yield row_no, col_no, value


def find_max_anchor(snapshot: SheetSnapshot) -> Optional[Tuple[int, int]]:
    candidates: List[Tuple[int, int, int]] = []
    for row, col, value in iter_snapshot_cells(snapshot):
        if normalize_text(value) != "max":
            continue
        score = 0
        for delta in range(-4, 5):
            if delta == 0:
                continue
            if normalize_text(snapshot.get(row, col + delta)) == "min":
                score += 5
        row_text = [
            normalize_text(snapshot.get(row, check_col))
            for check_col in range(snapshot.start_col, snapshot.end_col + 1)
        ]
        if any("quarter" in t and "num" in t for t in row_text):
            score += 3
        if any("forecast" in t or "fcst" in t for t in row_text):
            score += 2
        candidates.append((score, row, col))

    if not candidates:
        return None
    candidates.sort(key=lambda item: (-item[0], item[1], item[2]))
    _, row, col = candidates[0]
    return row, col


def find_best_header_row(snapshot: SheetSnapshot, anchor_row: int) -> int:
    best_row = anchor_row
    best_score = -1
    for row in range(max(snapshot.start_row, anchor_row - 2), min(snapshot.end_row, anchor_row + 2) + 1):
        score = 0
        for col in range(snapshot.start_col, snapshot.end_col + 1):
            text = normalize_text(snapshot.get(row, col))
            if not text:
                continue
            if any(char.isalpha() for char in text):
                score += 1
                if any(word in text for word in ("quarter", "sales", "forecast", "penetration", "intercept", "slope")):
                    score += 2
        if score > best_score:
            best_score = score
            best_row = row
    return best_row


def find_col_by_keyword_groups(
    snapshot: SheetSnapshot,
    rows_to_search: Sequence[int],
    keyword_groups: Sequence[Sequence[str]],
) -> Optional[int]:
    for row in rows_to_search:
        if row < snapshot.start_row or row > snapshot.end_row:
            continue
        for col in range(snapshot.start_col, snapshot.end_col + 1):
            text = normalize_for_key(snapshot.get(row, col))
            if not text:
                continue
            for keyword_group in keyword_groups:
                if all(keyword in text for keyword in keyword_group):
                    return col
    return None


def find_min_col_near_anchor(snapshot: SheetSnapshot, anchor_row: int, anchor_col: int) -> Optional[int]:
    for delta in range(-4, 5):
        if delta == 0:
            continue
        test_col = anchor_col + delta
        if normalize_text(snapshot.get(anchor_row, test_col)) == "min":
            return test_col
    return None


def find_label_value(snapshot: SheetSnapshot, keyword_groups: Sequence[Sequence[str]]) -> Any:
    for row, col, value in iter_snapshot_cells(snapshot):
        label = normalize_for_key(value)
        if not label:
            continue
        for keyword_group in keyword_groups:
            if all(keyword in label for keyword in keyword_group):
                right_value = snapshot.get(row, col + 1)
                if right_value not in (None, ""):
                    return right_value
                down_value = snapshot.get(row + 1, col)
                if down_value not in (None, ""):
                    return down_value
    return None


def collect_numeric_from_col(
    snapshot: SheetSnapshot,
    col: int,
    min_row: Optional[int] = None,
    max_row: Optional[int] = None,
    pct_mode: bool = False,
) -> List[Tuple[int, float]]:
    min_row = snapshot.start_row if min_row is None else max(min_row, snapshot.start_row)
    max_row = snapshot.end_row if max_row is None else min(max_row, snapshot.end_row)
    values: List[Tuple[int, float]] = []
    for row in range(min_row, max_row + 1):
        raw = snapshot.get(row, col)
        numeric = ensure_pct_decimal(raw) if pct_mode else to_float(raw)
        if numeric is None:
            continue
        values.append((row, numeric))
    return values


def choose_penetration_series(
    snapshot: SheetSnapshot,
    anchor_col: int,
    header_rows: Sequence[int],
) -> List[float]:
    candidate_cols: List[int] = []
    for header_row in header_rows:
        for col in range(snapshot.start_col, snapshot.end_col + 1):
            text = normalize_for_key(snapshot.get(header_row, col))
            if not text:
                continue
            if "penetration" in text or ("captured" in text and "db" in text):
                candidate_cols.append(col)
    candidate_cols.extend([anchor_col - 7, anchor_col - 8, anchor_col - 6])

    best_values: List[float] = []
    seen_cols = set()
    for col in candidate_cols:
        if col in seen_cols or col < snapshot.start_col or col > snapshot.end_col:
            continue
        seen_cols.add(col)
        series = collect_numeric_from_col(snapshot, col, pct_mode=True)
        pct_values = [value for _, value in series if 0 < value <= 1.5]
        if len(pct_values) > len(best_values):
            best_values = pct_values
    return best_values


def last_numeric_from_col(snapshot: SheetSnapshot, col: Optional[int], default: Any = None) -> Any:
    if col is None:
        return default
    for row in range(snapshot.end_row, snapshot.start_row - 1, -1):
        numeric = to_float(snapshot.get(row, col))
        if numeric is not None:
            return numeric
    return default


def set_formula2(cell: xw.Range, formula_r1c1: str) -> None:
    try:
        cell.formula2 = formula_r1c1
        return
    except Exception:
        pass
    try:
        cell.api.Formula2R1C1 = formula_r1c1
        return
    except Exception:
        pass
    cell.formula = formula_r1c1


def parse_file_metadata(file_path: Path) -> Dict[str, str]:
    stem = file_path.stem
    parts = [part.strip() for part in stem.split(" - ")]

    ticker = "UNKNOWN"
    period_token = None

    if len(parts) >= 2 and parts[1]:
        ticker = re.sub(r"[^A-Za-z0-9]", "", parts[1].upper()) or "UNKNOWN"

    if len(parts) >= 3:
        period_token = parts[2].split("_")[0].strip()

    if not period_token:
        fallback_match = re.search(r"(Early|Mid|Late)([A-Za-z]{3,9})(\d{4})", stem, flags=re.IGNORECASE)
        if fallback_match:
            period_token = fallback_match.group(0)

    if not period_token:
        model_period = "Unknown_0000"
        model_date = ""
        return {
            "model": f"{ticker}_{model_period}",
            "ticker": ticker,
            "model_period": model_period,
            "model_date": model_date,
        }

    period_match = re.match(r"^(Early|Mid|Late)([A-Za-z]{3,9})(\d{4})$", period_token, flags=re.IGNORECASE)
    if not period_match:
        model_period = "Unknown_0000"
        model_date = ""
        return {
            "model": f"{ticker}_{model_period}",
            "ticker": ticker,
            "model_period": model_period,
            "model_date": model_date,
        }

    period_phase = period_match.group(1).capitalize()
    month_text = period_match.group(2)
    year = int(period_match.group(3))
    month_key = month_text[:3].lower()
    month = MONTH_LOOKUP.get(month_key)
    if month is None:
        model_period = f"{period_phase}{month_text}_{year}"
        model_date = ""
        return {
            "model": f"{ticker}_{model_period}",
            "ticker": ticker,
            "model_period": model_period,
            "model_date": model_date,
        }

    month_label = month_text[:3].title()
    model_period = f"{period_phase}{month_label}_{year}"
    day = PERIOD_DAY_LOOKUP[period_phase.lower()]
    model_date = date(year, month, day).isoformat()

    return {
        "model": f"{ticker}_{model_period}",
        "ticker": ticker,
        "model_period": model_period,
        "model_date": model_date,
    }


def safe_close_source_workbook(workbook: xw.Book) -> None:
    close_attempts = [
        lambda: workbook.close(save=False),
        lambda: workbook.close(False),
        lambda: workbook.api.Close(False),
        lambda: workbook.api.Close(SaveChanges=False),
    ]
    for attempt in close_attempts:
        try:
            attempt()
            return
        except Exception:
            continue


def process_empirical_model(
    workbook: xw.Book,
    metadata: Dict[str, str],
    source_file: str,
) -> List[Dict[str, Any]]:
    if "Empirical Model" not in [sheet.name for sheet in workbook.sheets]:
        print(f"  Skipped empirical extraction: missing sheet 'Empirical Model'")
        return []

    sheet = workbook.sheets["Empirical Model"]
    snapshot = snapshot_sheet(sheet)
    if snapshot is None:
        print("  Skipped empirical extraction: sheet is empty")
        return []

    anchor = find_max_anchor(snapshot)
    if not anchor:
        print("  Skipped empirical extraction: could not find 'max' anchor")
        return []

    anchor_row, anchor_col = anchor
    header_row = find_best_header_row(snapshot, anchor_row)
    rows_to_search = [header_row, anchor_row, anchor_row - 1, anchor_row + 1]

    num_quarters_col = find_col_by_keyword_groups(
        snapshot,
        rows_to_search,
        keyword_groups=[("num", "quarter"), ("quarters", "used")],
    )
    last_quarter_col = find_col_by_keyword_groups(
        snapshot,
        rows_to_search,
        keyword_groups=[("last", "quarter")],
    )
    avg_penetration_col = find_col_by_keyword_groups(
        snapshot,
        rows_to_search,
        keyword_groups=[("avg", "penetration"), ("average", "penetration")],
    )
    quarterly_sales_col = find_col_by_keyword_groups(
        snapshot,
        rows_to_search,
        keyword_groups=[("quarterly", "sales")],
    )
    reported_sales_col = find_col_by_keyword_groups(
        snapshot,
        rows_to_search,
        keyword_groups=[("reported", "sales"), ("actual", "sales")],
    )
    growth_rate_col = find_col_by_keyword_groups(
        snapshot,
        rows_to_search,
        keyword_groups=[("growth", "rate"), ("growth",)],
    )
    sales_captured_col = find_col_by_keyword_groups(
        snapshot,
        rows_to_search,
        keyword_groups=[("captured", "db"), ("sales", "captured")],
    )
    forecast_value_col = find_col_by_keyword_groups(
        snapshot,
        rows_to_search,
        keyword_groups=[("estimated", "total", "sold"), ("forecast", "value"), ("tot", "fcst")],
    )
    min_col = find_min_col_near_anchor(snapshot, anchor_row, anchor_col)
    if min_col is None:
        min_col = find_col_by_keyword_groups(snapshot, rows_to_search, keyword_groups=[("min",)])

    table_rows: List[int] = []
    blank_streak = 0
    for row in range(anchor_row + 1, min(snapshot.end_row, anchor_row + 120) + 1):
        num_quarters_value = to_float(snapshot.get(row, num_quarters_col)) if num_quarters_col else None
        forecast_value = to_float(snapshot.get(row, forecast_value_col)) if forecast_value_col else None
        max_value = to_float(snapshot.get(row, anchor_col))
        min_value = to_float(snapshot.get(row, min_col)) if min_col else None
        avg_pen = ensure_pct_decimal(snapshot.get(row, avg_penetration_col)) if avg_penetration_col else None
        if any(value is not None for value in (num_quarters_value, forecast_value, max_value, min_value, avg_pen)):
            table_rows.append(row)
            blank_streak = 0
            if len(table_rows) >= 10:
                break
        elif table_rows:
            blank_streak += 1
            if blank_streak >= 3:
                break

    penetration_series = choose_penetration_series(snapshot, anchor_col, rows_to_search)

    quarterly_sales_scalar = (
        last_numeric_from_col(snapshot, quarterly_sales_col)
        if quarterly_sales_col is not None
        else to_float(find_label_value(snapshot, keyword_groups=[("quarterly", "sales")]))
    )
    if quarterly_sales_scalar is None:
        quarterly_sales_scalar = to_float(find_label_value(snapshot, keyword_groups=[("sales", "quarter")]))

    reported_sales_scalar = (
        last_numeric_from_col(snapshot, reported_sales_col)
        if reported_sales_col is not None
        else to_float(find_label_value(snapshot, keyword_groups=[("reported", "sales"), ("actual", "sales")]))
    )
    growth_rate_scalar = (
        last_numeric_from_col(snapshot, growth_rate_col)
        if growth_rate_col is not None
        else ensure_pct_decimal(find_label_value(snapshot, keyword_groups=[("growth", "rate")]))
    )
    sales_captured_scalar = (
        last_numeric_from_col(snapshot, sales_captured_col)
        if sales_captured_col is not None
        else ensure_pct_decimal(find_label_value(snapshot, keyword_groups=[("captured", "db"), ("sales", "captured")]))
    )

    formula_results: Dict[int, Dict[str, Optional[float]]] = {}
    max_n = min(10, len(penetration_series))
    if max_n > 0:
        scratch_col = snapshot.end_col + 2
        history_start_row = snapshot.end_row + 2
        history_values = [[value] for value in penetration_series]
        sheet.range((history_start_row, scratch_col), (history_start_row + len(history_values) - 1, scratch_col)).value = history_values
        history_end_row = history_start_row + len(history_values) - 1

        calc_rows: Dict[int, int] = {}
        for n_quarters in range(1, max_n + 1):
            calc_row = history_end_row + n_quarters
            calc_rows[n_quarters] = calc_row
            first_row = history_end_row - n_quarters + 1

            set_formula2(
                sheet.range((calc_row, scratch_col)),
                f"=AVERAGE(R{first_row}C{scratch_col}:R{history_end_row}C{scratch_col})",
            )
            set_formula2(
                sheet.range((calc_row, scratch_col + 5)),
                f"=MIN(R{first_row}C{scratch_col}:R{history_end_row}C{scratch_col})",
            )
            set_formula2(
                sheet.range((calc_row, scratch_col + 6)),
                f"=MAX(R{first_row}C{scratch_col}:R{history_end_row}C{scratch_col})",
            )

            if quarterly_sales_scalar is not None:
                sheet.range((calc_row, scratch_col + 1)).value = quarterly_sales_scalar
                set_formula2(
                    sheet.range((calc_row, scratch_col + 2)),
                    f'=IF(R{calc_row}C{scratch_col}=0,"",R{calc_row}C{scratch_col + 1}/R{calc_row}C{scratch_col})',
                )
                set_formula2(
                    sheet.range((calc_row, scratch_col + 3)),
                    f'=IF(R{calc_row}C{scratch_col + 5}=0,"",R{calc_row}C{scratch_col + 1}/R{calc_row}C{scratch_col + 5})',
                )
                set_formula2(
                    sheet.range((calc_row, scratch_col + 4)),
                    f'=IF(R{calc_row}C{scratch_col + 6}=0,"",R{calc_row}C{scratch_col + 1}/R{calc_row}C{scratch_col + 6})',
                )

        workbook.app.calculate()

        for n_quarters, calc_row in calc_rows.items():
            formula_results[n_quarters] = {
                "avg_penetration": ensure_pct_decimal(sheet.range((calc_row, scratch_col)).value),
                "forecast_value": to_float(sheet.range((calc_row, scratch_col + 2)).value),
                "forecast_max": to_float(sheet.range((calc_row, scratch_col + 3)).value),
                "forecast_min": to_float(sheet.range((calc_row, scratch_col + 4)).value),
            }

    empirical_rows: List[Dict[str, Any]] = []
    if table_rows:
        n_limit = min(10, len(table_rows))
        for idx in range(n_limit):
            row_no = table_rows[idx]
            default_n = idx + 1
            num_quarters_used = int(round(to_float(snapshot.get(row_no, num_quarters_col)) or default_n)) if num_quarters_col else default_n

            table_avg_pen = ensure_pct_decimal(snapshot.get(row_no, avg_penetration_col)) if avg_penetration_col else None
            table_forecast = to_float(snapshot.get(row_no, forecast_value_col)) if forecast_value_col else None
            table_max = to_float(snapshot.get(row_no, anchor_col))
            table_min = to_float(snapshot.get(row_no, min_col)) if min_col else None

            formula_for_n = formula_results.get(num_quarters_used) or formula_results.get(default_n) or {}
            avg_pen = table_avg_pen if table_avg_pen is not None else formula_for_n.get("avg_penetration")
            forecast_value = table_forecast if table_forecast is not None else formula_for_n.get("forecast_value")
            forecast_max = table_max if table_max is not None else formula_for_n.get("forecast_max")
            forecast_min = table_min if table_min is not None else formula_for_n.get("forecast_min")

            quarterly_sales = (
                to_float(snapshot.get(row_no, quarterly_sales_col))
                if quarterly_sales_col
                else quarterly_sales_scalar
            )
            if quarterly_sales is None:
                quarterly_sales = quarterly_sales_scalar

            reported_sales = (
                to_float(snapshot.get(row_no, reported_sales_col))
                if reported_sales_col
                else reported_sales_scalar
            )
            if reported_sales is None:
                reported_sales = reported_sales_scalar

            growth_rate = (
                ensure_pct_decimal(snapshot.get(row_no, growth_rate_col))
                if growth_rate_col
                else growth_rate_scalar
            )
            if growth_rate is None:
                growth_rate = growth_rate_scalar

            sales_captured = (
                ensure_pct_decimal(snapshot.get(row_no, sales_captured_col))
                if sales_captured_col
                else sales_captured_scalar
            )
            if sales_captured is None:
                sales_captured = sales_captured_scalar

            last_quarter_used = snapshot.get(row_no, last_quarter_col) if last_quarter_col else None
            range_width = (
                forecast_max - forecast_min
                if forecast_max is not None and forecast_min is not None
                else None
            )

            empirical_rows.append(
                {
                    "model": metadata["model"],
                    "ticker": metadata["ticker"],
                    "model_period": metadata["model_period"],
                    "model_date": metadata["model_date"],
                    "method": "empirical",
                    "parameter_name": "avg_penetration_pct",
                    "parameter_value": avg_pen,
                    "num_quarters_used": num_quarters_used,
                    "last_quarter_used": last_quarter_used,
                    "forecast_value": forecast_value,
                    "actual_value": reported_sales,
                    "forecast_max": forecast_max,
                    "forecast_min": forecast_min,
                    "range_width": range_width,
                    "avg_penetration_pct": avg_pen,
                    "quarterly_sales": quarterly_sales,
                    "reported_sales": reported_sales,
                    "growth_rate_pct": growth_rate,
                    "sales_captured_in_db_pct": sales_captured,
                    "source_file": source_file,
                }
            )
    else:
        if max_n == 0:
            return empirical_rows

        for n_quarters in range(1, max_n + 1):
            formula_for_n = formula_results.get(n_quarters, {})
            avg_pen = formula_for_n.get("avg_penetration")
            forecast_value = formula_for_n.get("forecast_value")
            forecast_max = formula_for_n.get("forecast_max")
            forecast_min = formula_for_n.get("forecast_min")
            range_width = (
                forecast_max - forecast_min
                if forecast_max is not None and forecast_min is not None
                else None
            )
            empirical_rows.append(
                {
                    "model": metadata["model"],
                    "ticker": metadata["ticker"],
                    "model_period": metadata["model_period"],
                    "model_date": metadata["model_date"],
                    "method": "empirical",
                    "parameter_name": "avg_penetration_pct",
                    "parameter_value": avg_pen,
                    "num_quarters_used": n_quarters,
                    "last_quarter_used": None,
                    "forecast_value": forecast_value,
                    "actual_value": reported_sales_scalar,
                    "forecast_max": forecast_max,
                    "forecast_min": forecast_min,
                    "range_width": range_width,
                    "avg_penetration_pct": avg_pen,
                    "quarterly_sales": quarterly_sales_scalar,
                    "reported_sales": reported_sales_scalar,
                    "growth_rate_pct": growth_rate_scalar,
                    "sales_captured_in_db_pct": sales_captured_scalar,
                    "source_file": source_file,
                }
            )

    return empirical_rows


def get_longest_contiguous_block(pairs: Sequence[Tuple[int, float, float]]) -> List[Tuple[int, float, float]]:
    if not pairs:
        return []
    sorted_pairs = sorted(pairs, key=lambda item: item[0])
    blocks: List[List[Tuple[int, float, float]]] = []
    current_block: List[Tuple[int, float, float]] = [sorted_pairs[0]]
    for pair in sorted_pairs[1:]:
        if pair[0] - current_block[-1][0] <= 1:
            current_block.append(pair)
        else:
            blocks.append(current_block)
            current_block = [pair]
    blocks.append(current_block)
    blocks.sort(key=lambda block: (len(block), block[-1][0]), reverse=True)
    return blocks[0]


def process_regression_model(
    workbook: xw.Book,
    metadata: Dict[str, str],
    source_file: str,
) -> List[Dict[str, Any]]:
    if "Regression Model" not in [sheet.name for sheet in workbook.sheets]:
        print("  Skipped regression extraction: missing sheet 'Regression Model'")
        return []

    sheet = workbook.sheets["Regression Model"]
    snapshot = snapshot_sheet(sheet)
    if snapshot is None:
        print("  Skipped regression extraction: sheet is empty")
        return []

    anchor = find_max_anchor(snapshot)
    if not anchor:
        print("  Skipped regression extraction: could not find 'max' anchor")
        return []

    anchor_row, anchor_col = anchor
    y_col = anchor_col - 7
    x_col = anchor_col - 11

    xy_pairs: List[Tuple[int, float, float]] = []
    for row in range(snapshot.start_row, snapshot.end_row + 1):
        x_val = to_float(snapshot.get(row, x_col))
        y_val = to_float(snapshot.get(row, y_col))
        if x_val is None or y_val is None:
            continue
        xy_pairs.append((row, x_val, y_val))

    contiguous_block = get_longest_contiguous_block(xy_pairs)
    if len(contiguous_block) < 2:
        print("  Skipped regression extraction: insufficient x/y history for regression")
        return []

    history = contiguous_block[-10:]
    max_n = min(10, len(history))
    if max_n == 0:
        return []

    header_row = find_best_header_row(snapshot, anchor_row)
    rows_to_search = [header_row, anchor_row, anchor_row - 1, anchor_row + 1]
    num_quarters_col = find_col_by_keyword_groups(
        snapshot,
        rows_to_search,
        keyword_groups=[("num", "quarter"), ("quarters", "used")],
    )
    forecast_col = find_col_by_keyword_groups(
        snapshot,
        rows_to_search,
        keyword_groups=[
            ("tot", "fcst"),
            ("forecast", "without", "sa"),
            ("forecast", "w/o", "sa"),
            ("forecast", "value"),
        ],
    )
    actual_col = find_col_by_keyword_groups(
        snapshot,
        rows_to_search,
        keyword_groups=[("actual", "sales"), ("reported", "sales"), ("actual",)],
    )
    min_col = find_min_col_near_anchor(snapshot, anchor_row, anchor_col)
    if min_col is None:
        min_col = find_col_by_keyword_groups(snapshot, rows_to_search, keyword_groups=[("min",)])

    row_for_num_quarters: Dict[int, int] = {}
    if num_quarters_col:
        blank_streak = 0
        for row in range(anchor_row + 1, min(snapshot.end_row, anchor_row + 150) + 1):
            row_value = to_float(snapshot.get(row, num_quarters_col))
            if row_value is None:
                blank_streak += 1
                if blank_streak >= 4 and row_for_num_quarters:
                    break
                continue
            blank_streak = 0
            row_for_num_quarters[int(round(row_value))] = row

    scratch_col = snapshot.end_col + 2
    scratch_start = snapshot.end_row + 2
    calc_cells: Dict[int, Dict[str, xw.Range]] = {}

    next_block_row = scratch_start
    for n_quarters in range(2, max_n + 1):
        points = history[-n_quarters:]
        x_vals = [[point[1]] for point in points]
        y_vals = [[point[2]] for point in points]
        data_start = next_block_row
        data_end = data_start + n_quarters - 1

        sheet.range((data_start, scratch_col), (data_end, scratch_col)).value = x_vals
        sheet.range((data_start, scratch_col + 1), (data_end, scratch_col + 1)).value = y_vals

        intercept_cell = sheet.range((data_end + 1, scratch_col))
        slope_cell = sheet.range((data_end + 1, scratch_col + 1))
        forecast_cell = sheet.range((data_end + 1, scratch_col + 2))

        set_formula2(
            intercept_cell,
            f"=INTERCEPT(R{data_start}C{scratch_col + 1}:R{data_end}C{scratch_col + 1},R{data_start}C{scratch_col}:R{data_end}C{scratch_col})",
        )
        set_formula2(
            slope_cell,
            f"=SLOPE(R{data_start}C{scratch_col + 1}:R{data_end}C{scratch_col + 1},R{data_start}C{scratch_col}:R{data_end}C{scratch_col})",
        )
        next_x = points[-1][1] + 1
        set_formula2(
            forecast_cell,
            f"=R{data_end + 1}C{scratch_col}+R{data_end + 1}C{scratch_col + 1}*{next_x}",
        )

        calc_cells[n_quarters] = {
            "intercept": intercept_cell,
            "slope": slope_cell,
            "forecast": forecast_cell,
        }
        next_block_row = data_end + 4

    if calc_cells:
        workbook.app.calculate()

    regression_rows: List[Dict[str, Any]] = []
    previous_signature: Optional[Tuple[Any, ...]] = None
    for n_quarters in range(1, max_n + 1):
        if n_quarters == 1:
            y_last = history[-1][2]
            intercept = y_last
            slope = 0.0
            forecast_value = y_last
        else:
            intercept = to_float(calc_cells[n_quarters]["intercept"].value)
            slope = to_float(calc_cells[n_quarters]["slope"].value)
            forecast_value = to_float(calc_cells[n_quarters]["forecast"].value)

        result_row = row_for_num_quarters.get(n_quarters, anchor_row + n_quarters)
        table_forecast = to_float(snapshot.get(result_row, forecast_col)) if forecast_col else None
        table_max = to_float(snapshot.get(result_row, anchor_col))
        table_min = to_float(snapshot.get(result_row, min_col)) if min_col else None
        table_actual = to_float(snapshot.get(result_row, actual_col)) if actual_col else None

        if table_forecast is not None:
            forecast_value = table_forecast

        forecast_max = table_max if table_max is not None else forecast_value
        forecast_min = table_min if table_min is not None else forecast_value
        range_width = (
            forecast_max - forecast_min
            if forecast_max is not None and forecast_min is not None
            else None
        )

        signature = (
            round(intercept, 10) if intercept is not None else None,
            round(slope, 10) if slope is not None else None,
            round(forecast_value, 10) if forecast_value is not None else None,
            round(forecast_max, 10) if forecast_max is not None else None,
            round(forecast_min, 10) if forecast_min is not None else None,
        )
        if previous_signature is not None and signature == previous_signature:
            continue
        previous_signature = signature

        regression_rows.append(
            {
                "model": metadata["model"],
                "ticker": metadata["ticker"],
                "model_period": metadata["model_period"],
                "model_date": metadata["model_date"],
                "method": "regression",
                "parameter_name": "num_quarters_used",
                "parameter_value": n_quarters,
                "num_quarters_used": n_quarters,
                "forecast_value": forecast_value,
                "actual_value": table_actual,
                "forecast_max": forecast_max,
                "forecast_min": forecast_min,
                "range_width": range_width,
                "intercept": intercept,
                "slope": slope,
                "source_file": source_file,
            }
        )

    return regression_rows


def unique_output_path(input_path: Path, output_path: Path) -> Path:
    output_path.mkdir(parents=True, exist_ok=True)
    folder_name = input_path.name
    base_stem = f"{folder_name}_PARAM"
    candidate = output_path / f"{base_stem}.xlsx"
    if not candidate.exists():
        return candidate
    suffix = 1
    while True:
        candidate = output_path / f"{base_stem}.{suffix}.xlsx"
        if not candidate.exists():
            return candidate
        suffix += 1


def write_output_workbook(
    empirical_rows: Sequence[Dict[str, Any]],
    regression_rows: Sequence[Dict[str, Any]],
    output_file: Path,
) -> None:
    workbook = Workbook()
    empirical_sheet = workbook.active
    empirical_sheet.title = "empirical_candidates"
    regression_sheet = workbook.create_sheet(title="regression_candidates")

    def write_sheet(sheet: Any, headers: Sequence[str], rows: Sequence[Dict[str, Any]]) -> None:
        sheet.append(list(headers))
        for row in rows:
            sheet.append([row.get(header) for header in headers])

        for cell in sheet[1]:
            cell.font = Font(bold=True)

        sheet.freeze_panes = "A2"
        if sheet.max_column > 0 and sheet.max_row > 0:
            sheet.auto_filter.ref = sheet.dimensions

        for col_idx, header in enumerate(headers, start=1):
            max_len = len(header)
            for row_idx in range(2, sheet.max_row + 1):
                cell_value = sheet.cell(row=row_idx, column=col_idx).value
                if cell_value is None:
                    continue
                max_len = max(max_len, len(str(cell_value)))
            sheet.column_dimensions[get_column_letter(col_idx)].width = min(max(max_len + 2, 12), 42)

    write_sheet(empirical_sheet, EMPIRICAL_HEADERS, empirical_rows)
    write_sheet(regression_sheet, REGRESSION_HEADERS, regression_rows)
    workbook.save(output_file)


def source_workbooks(input_path: Path) -> Iterable[Path]:
    for file_path in sorted(input_path.iterdir()):
        if not file_path.is_file():
            continue
        if file_path.name.startswith("~"):
            print(f"Skipped {file_path.name}: temp file")
            continue
        if file_path.suffix.lower() != ".xlsx":
            print(f"Skipped {file_path.name}: not an .xlsx file")
            continue
        yield file_path


def main() -> None:
    input_path = Path(input_dir).expanduser().resolve()
    output_path = Path(output_dir).expanduser().resolve()

    if not input_path.exists() or not input_path.is_dir():
        raise SystemExit(f"Input folder does not exist or is not a directory: {input_path}")

    output_file = unique_output_path(input_path=input_path, output_path=output_path)
    output_file_name = output_file.name

    empirical_rows: List[Dict[str, Any]] = []
    regression_rows: List[Dict[str, Any]] = []
    processed_files = 0

    app = xw.App(visible=False, add_book=False)
    app.display_alerts = False
    app.screen_updating = False
    original_calculation = app.calculation
    app.calculation = "manual"

    try:
        for file_path in source_workbooks(input_path):
            if file_path.name == output_file_name:
                print(f"Skipped {file_path.name}: output file target")
                continue

            print(f"Processing {file_path.name}")
            workbook = None
            try:
                workbook = app.books.open(str(file_path), update_links=False)
                metadata = parse_file_metadata(file_path)
                empirical_rows.extend(
                    process_empirical_model(
                        workbook=workbook,
                        metadata=metadata,
                        source_file=file_path.name,
                    )
                )
                regression_rows.extend(
                    process_regression_model(
                        workbook=workbook,
                        metadata=metadata,
                        source_file=file_path.name,
                    )
                )
                processed_files += 1
            except Exception as exc:
                print(f"  Skipped {file_path.name}: processing error ({exc})")
            finally:
                if workbook is not None:
                    safe_close_source_workbook(workbook)
    finally:
        try:
            app.calculation = original_calculation
        except Exception:
            pass
        app.quit()

    write_output_workbook(
        empirical_rows=empirical_rows,
        regression_rows=regression_rows,
        output_file=output_file,
    )

    print(f"Output workbook: {output_file}")
    print(f"Files processed: {processed_files}")
    print(f"Empirical rows: {len(empirical_rows)}")
    print(f"Regression rows: {len(regression_rows)}")


if __name__ == "__main__":
    main()
