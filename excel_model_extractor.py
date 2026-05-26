#!/usr/bin/env python3
"""Extract empirical/regression model candidates from Excel workbooks.

This script scans all .xlsx files in input_dir, opens each workbook exactly
once, extracts data from both "Empirical Model" and "Regression Model", and
writes one consolidated output workbook with:
  - empirical_candidates
  - regression_candidates
"""

from __future__ import annotations

import math
import re
from dataclasses import dataclass
from datetime import date
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Sequence, Tuple

try:
    import xlwings as xw
except ImportError:  # pragma: no cover - environment dependent
    xw = None

try:
    from openpyxl import Workbook
    from openpyxl.styles import Font
    from openpyxl.utils import get_column_letter
    from openpyxl.worksheet.worksheet import Worksheet
except ImportError:  # pragma: no cover - environment dependent
    Workbook = None
    Font = None
    get_column_letter = None
    Worksheet = Any


# ---------------------------------------------------------------------------
# User-editable paths
# ---------------------------------------------------------------------------
input_dir = Path("input")
output_dir = Path("output")


# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------
N_QUARTERS = 10

EMPIRICAL_COLUMNS = [
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

REGRESSION_COLUMNS = [
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

MONTH_MAP = {
    "jan": 1,
    "january": 1,
    "feb": 2,
    "february": 2,
    "mar": 3,
    "march": 3,
    "apr": 4,
    "april": 4,
    "may": 5,
    "jun": 6,
    "june": 6,
    "jul": 7,
    "july": 7,
    "aug": 8,
    "august": 8,
    "sep": 9,
    "sept": 9,
    "september": 9,
    "oct": 10,
    "october": 10,
    "nov": 11,
    "november": 11,
    "dec": 12,
    "december": 12,
}

MODEL_DAY_MAP = {"early": 5, "mid": 15, "late": 25}


EMPIRICAL_ALIAS_MAP = {
    "num_quarters_used": [
        "num quarters used",
        "number of quarters used",
        "quarters used",
        "num quarters",
        "n quarters",
        "qtrs used",
    ],
    "last_quarter_used": [
        "last quarter used",
        "last qtr used",
        "latest quarter used",
    ],
    "forecast_value": [
        "estimated total sold",
        "total forecast",
        "tot fcst",
        "forecast value",
        "forecast",
    ],
    "actual_value": [
        "actual sales",
        "reported sales",
        "actual value",
        "actual",
    ],
    "forecast_max": ["max"],
    "forecast_min": ["min"],
    "avg_penetration_pct": [
        "avg penetration",
        "average penetration",
        "avg pen",
        "penetration pct",
    ],
    "quarterly_sales": ["quarterly sales", "quarter sales", "sales per quarter"],
    "reported_sales": ["reported sales", "reported"],
    "growth_rate_pct": ["growth rate", "growth pct", "growth"],
    "sales_captured_in_db_pct": [
        "sales captured in db",
        "captured in db",
        "capture pct",
    ],
}

REGRESSION_ALIAS_MAP = {
    "num_quarters_used": [
        "num quarters used",
        "number of quarters used",
        "quarters used",
        "num quarters",
        "n quarters",
        "qtrs used",
    ],
    "forecast_value": [
        "tot fcst w/o sa",
        "total forecast w/o sa",
        "forecast without sa",
        "tot fcst",
        "forecast value",
    ],
    "actual_value": ["actual", "actual value", "reported sales", "actual sales"],
    "forecast_max": ["max"],
    "forecast_min": ["min"],
}


# Anchor-relative fallback offsets (used when headers are not detected).
EMPIRICAL_FALLBACK_OFFSETS = {
    "num_quarters_used": -8,
    "last_quarter_used": -7,
    "forecast_value": -2,
    "actual_value": -1,
    "forecast_max": 0,
    "forecast_min": 1,
    "avg_penetration_pct": -5,
    "quarterly_sales": -4,
    "reported_sales": -3,
    "growth_rate_pct": -6,
    "sales_captured_in_db_pct": -9,
}

REGRESSION_FALLBACK_OFFSETS = {
    "num_quarters_used": -6,
    "forecast_value": -1,
    "actual_value": -2,
    "forecast_max": 0,
    "forecast_min": 1,
}


@dataclass
class FileLabel:
    ticker: str
    model_period: str
    model_date: str
    model: str


@dataclass
class SheetSnapshot:
    origin_row: int
    origin_col: int
    matrix: List[List[Any]]

    @property
    def row_count(self) -> int:
        return len(self.matrix)

    @property
    def col_count(self) -> int:
        return len(self.matrix[0]) if self.matrix else 0

    @property
    def end_row(self) -> int:
        return self.origin_row + self.row_count - 1

    @property
    def end_col(self) -> int:
        return self.origin_col + self.col_count - 1

    def get(self, row: int, col: int) -> Any:
        r_idx = row - self.origin_row
        c_idx = col - self.origin_col
        if r_idx < 0 or c_idx < 0:
            return None
        if r_idx >= self.row_count or c_idx >= self.col_count:
            return None
        return self.matrix[r_idx][c_idx]


def as_2d(values: Any) -> List[List[Any]]:
    if values is None:
        return []
    if isinstance(values, tuple):
        values = list(values)
    if isinstance(values, list):
        if not values:
            return []
        first = values[0]
        if isinstance(first, tuple):
            values = [list(row) for row in values]
            first = values[0] if values else []
        if isinstance(first, list):
            return values
        return [values]
    return [[values]]


def normalize_text(value: Any) -> str:
    if value is None:
        return ""
    text = str(value).strip().lower()
    if not text:
        return ""
    text = re.sub(r"[^a-z0-9]+", " ", text)
    return re.sub(r"\s+", " ", text).strip()


def coerce_float(value: Any) -> Optional[float]:
    if value is None:
        return None
    if isinstance(value, bool):
        return float(value)
    if isinstance(value, (int, float)):
        if isinstance(value, float) and math.isnan(value):
            return None
        return float(value)
    if isinstance(value, str):
        text = value.strip()
        if not text:
            return None
        pct = False
        if text.endswith("%"):
            pct = True
            text = text[:-1]
        text = text.replace(",", "").replace("$", "")
        try:
            numeric = float(text)
        except ValueError:
            return None
        return numeric / 100.0 if pct else numeric
    return None


def safe_subtract(a: Any, b: Any) -> Optional[float]:
    a_float = coerce_float(a)
    b_float = coerce_float(b)
    if a_float is None or b_float is None:
        return None
    return a_float - b_float


def safe_close_book(wb: xw.Book) -> None:
    try:
        wb.close(save=False)
        return
    except TypeError:
        pass
    except Exception:
        pass
    try:
        wb.api.Close(SaveChanges=False)
    except Exception:
        try:
            wb.close()
        except Exception:
            pass


def set_formula2(rng: xw.Range, formula: str) -> None:
    try:
        rng.formula2 = formula
    except Exception:
        rng.formula = formula


def parse_model_token(period_token: str) -> Optional[Tuple[str, int, int]]:
    token = period_token.strip()
    token = re.sub(r"_?send$", "", token, flags=re.IGNORECASE)
    token = token.replace("_", "").replace(" ", "")
    match = re.match(
        r"^(Early|Mid|Late)([A-Za-z]{3,9})(\d{4})$",
        token,
        flags=re.IGNORECASE,
    )
    if not match:
        return None
    cadence_raw, month_raw, year_raw = match.groups()
    cadence = cadence_raw.capitalize()
    month_num = MONTH_MAP.get(month_raw.lower())
    if month_num is None:
        return None
    year = int(year_raw)
    return cadence, month_num, year


def parse_file_label(file_name: str) -> Optional[FileLabel]:
    stem = Path(file_name).stem
    parts = [part.strip() for part in stem.split(" - ")]
    if len(parts) < 3:
        return None
    ticker = parts[1].upper()
    token_data = parse_model_token(parts[2])
    if token_data is None:
        return None
    cadence, month_num, year = token_data
    day = MODEL_DAY_MAP[cadence.lower()]
    month_abbr = date(year, month_num, day).strftime("%b")
    model_period = f"{cadence}{month_abbr}_{year}"
    model_date = date(year, month_num, day).isoformat()
    model = f"{ticker}_{model_period}"
    return FileLabel(ticker=ticker, model_period=model_period, model_date=model_date, model=model)


def get_output_path(in_dir: Path, out_dir: Path) -> Path:
    out_dir.mkdir(parents=True, exist_ok=True)
    base_name = f"{in_dir.name}_PARAM.xlsx"
    candidate = out_dir / base_name
    suffix = 1
    while candidate.exists():
        candidate = out_dir / f"{in_dir.name}_PARAM.{suffix}.xlsx"
        suffix += 1
    return candidate


def locate_sheet(wb: xw.Book, expected_name: str) -> Optional[xw.Sheet]:
    expected = expected_name.strip().lower()
    for sheet in wb.sheets:
        if sheet.name.strip().lower() == expected:
            return sheet
    return None


def snapshot_sheet(sheet: xw.Sheet) -> SheetSnapshot:
    used = sheet.used_range
    matrix = as_2d(used.value)
    return SheetSnapshot(origin_row=used.row, origin_col=used.column, matrix=matrix)


def find_anchor_max(snapshot: SheetSnapshot) -> Optional[Tuple[int, int]]:
    candidates: List[Tuple[int, int, int]] = []
    for r_idx, row in enumerate(snapshot.matrix):
        for c_idx, cell in enumerate(row):
            if normalize_text(cell) == "max":
                row_num = snapshot.origin_row + r_idx
                col_num = snapshot.origin_col + c_idx
                min_bonus = 0
                right_text = normalize_text(snapshot.get(row_num, col_num + 1))
                left_text = normalize_text(snapshot.get(row_num, col_num - 1))
                if right_text == "min" or left_text == "min":
                    min_bonus = -1
                candidates.append((min_bonus, row_num, col_num))
    if not candidates:
        return None
    candidates.sort(key=lambda item: (item[0], item[1], item[2]))
    _, row_num, col_num = candidates[0]
    return row_num, col_num


def detect_offsets(
    snapshot: SheetSnapshot,
    anchor_row: int,
    anchor_col: int,
    alias_map: Dict[str, List[str]],
    fallback_offsets: Dict[str, int],
) -> Dict[str, int]:
    offsets: Dict[str, int] = {}
    search_rows = [anchor_row, anchor_row - 1, anchor_row + 1]
    search_min_col = max(snapshot.origin_col, anchor_col - 40)
    search_max_col = min(snapshot.end_col, anchor_col + 40)

    normalized_aliases = {
        key: [normalize_text(alias) for alias in aliases]
        for key, aliases in alias_map.items()
    }

    for field, aliases in normalized_aliases.items():
        best_match: Optional[Tuple[int, int]] = None
        for row in search_rows:
            if row < snapshot.origin_row or row > snapshot.end_row:
                continue
            for col in range(search_min_col, search_max_col + 1):
                text = normalize_text(snapshot.get(row, col))
                if not text:
                    continue
                if any(alias in text for alias in aliases):
                    distance = abs(row - anchor_row) * 3 + abs(col - anchor_col)
                    if best_match is None or distance < best_match[0]:
                        best_match = (distance, col)
        if best_match is not None:
            offsets[field] = best_match[1] - anchor_col

    for field, offset in fallback_offsets.items():
        offsets.setdefault(field, offset)

    return offsets


def read_block(
    sheet: xw.Sheet,
    start_row: int,
    end_row: int,
    start_col: int,
    end_col: int,
) -> List[List[Any]]:
    if start_row > end_row or start_col > end_col:
        return []
    values = sheet.range((start_row, start_col), (end_row, end_col)).value
    return as_2d(values)


def block_value(
    block: List[List[Any]],
    row: int,
    col: int,
    start_row: int,
    start_col: int,
) -> Any:
    r_idx = row - start_row
    c_idx = col - start_col
    if r_idx < 0 or c_idx < 0:
        return None
    if r_idx >= len(block):
        return None
    row_values = block[r_idx]
    if c_idx >= len(row_values):
        return None
    return row_values[c_idx]


def is_numeric(value: Any) -> bool:
    return coerce_float(value) is not None


def extract_empirical_rows(
    wb: xw.Book,
    file_label: FileLabel,
    source_file: str,
) -> List[Dict[str, Any]]:
    sheet = locate_sheet(wb, "Empirical Model")
    if sheet is None:
        return []

    snapshot = snapshot_sheet(sheet)
    anchor = find_anchor_max(snapshot)
    if anchor is None:
        return []
    anchor_row, anchor_col = anchor

    offsets = detect_offsets(
        snapshot=snapshot,
        anchor_row=anchor_row,
        anchor_col=anchor_col,
        alias_map=EMPIRICAL_ALIAS_MAP,
        fallback_offsets=EMPIRICAL_FALLBACK_OFFSETS,
    )

    data_start_row = anchor_row + 1
    data_end_row = data_start_row + N_QUARTERS - 1

    cols_needed = [anchor_col + offset for offset in offsets.values()]
    if not cols_needed:
        return []
    data_start_col = min(cols_needed)
    data_end_col = max(cols_needed)
    data_block = read_block(sheet, data_start_row, data_end_row, data_start_col, data_end_col)

    # Empirical average-penetration calculation via R1C1 formula2.
    avg_pen_col = anchor_col + offsets["avg_penetration_pct"]
    scratch_col = max(snapshot.end_col + 3, data_end_col + 3)
    scratch_start_row = data_start_row

    wrote_any_formula = False
    for idx in range(N_QUARTERS):
        row = data_start_row + idx
        formula = f'=IFERROR(AVERAGE(R{data_start_row}C{avg_pen_col}:R{row}C{avg_pen_col}),"")'
        set_formula2(sheet.range((scratch_start_row + idx, scratch_col)), formula)
        wrote_any_formula = True

    avg_pen_calc_values: List[Any] = [None] * N_QUARTERS
    if wrote_any_formula:
        wb.app.calculate()
        calc_matrix = read_block(
            sheet,
            scratch_start_row,
            scratch_start_row + N_QUARTERS - 1,
            scratch_col,
            scratch_col,
        )
        for idx in range(N_QUARTERS):
            avg_pen_calc_values[idx] = block_value(
                calc_matrix,
                scratch_start_row + idx,
                scratch_col,
                scratch_start_row,
                scratch_col,
            )

    rows: List[Dict[str, Any]] = []
    for idx in range(N_QUARTERS):
        row_num = data_start_row + idx

        num_quarters = block_value(
            data_block,
            row_num,
            anchor_col + offsets["num_quarters_used"],
            data_start_row,
            data_start_col,
        )
        if not is_numeric(num_quarters):
            num_quarters = idx + 1

        last_quarter_used = block_value(
            data_block,
            row_num,
            anchor_col + offsets["last_quarter_used"],
            data_start_row,
            data_start_col,
        )
        forecast_value = block_value(
            data_block,
            row_num,
            anchor_col + offsets["forecast_value"],
            data_start_row,
            data_start_col,
        )
        actual_value = block_value(
            data_block,
            row_num,
            anchor_col + offsets["actual_value"],
            data_start_row,
            data_start_col,
        )
        forecast_max = block_value(
            data_block,
            row_num,
            anchor_col + offsets["forecast_max"],
            data_start_row,
            data_start_col,
        )
        forecast_min = block_value(
            data_block,
            row_num,
            anchor_col + offsets["forecast_min"],
            data_start_row,
            data_start_col,
        )
        range_width = safe_subtract(forecast_max, forecast_min)

        avg_pen_cell = block_value(
            data_block,
            row_num,
            anchor_col + offsets["avg_penetration_pct"],
            data_start_row,
            data_start_col,
        )
        avg_pen_formula = avg_pen_calc_values[idx] if idx < len(avg_pen_calc_values) else None
        avg_penetration_pct = avg_pen_formula if avg_pen_formula not in (None, "") else avg_pen_cell

        quarterly_sales = block_value(
            data_block,
            row_num,
            anchor_col + offsets["quarterly_sales"],
            data_start_row,
            data_start_col,
        )
        reported_sales = block_value(
            data_block,
            row_num,
            anchor_col + offsets["reported_sales"],
            data_start_row,
            data_start_col,
        )
        growth_rate_pct = block_value(
            data_block,
            row_num,
            anchor_col + offsets["growth_rate_pct"],
            data_start_row,
            data_start_col,
        )
        sales_captured_pct = block_value(
            data_block,
            row_num,
            anchor_col + offsets["sales_captured_in_db_pct"],
            data_start_row,
            data_start_col,
        )

        row_data = {
            "model": file_label.model,
            "ticker": file_label.ticker,
            "model_period": file_label.model_period,
            "model_date": file_label.model_date,
            "method": "empirical",
            "parameter_name": "avg_penetration_pct",
            "parameter_value": avg_penetration_pct,
            "num_quarters_used": num_quarters,
            "last_quarter_used": last_quarter_used,
            "forecast_value": forecast_value,
            "actual_value": actual_value,
            "forecast_max": forecast_max,
            "forecast_min": forecast_min,
            "range_width": range_width,
            "avg_penetration_pct": avg_penetration_pct,
            "quarterly_sales": quarterly_sales,
            "reported_sales": reported_sales,
            "growth_rate_pct": growth_rate_pct,
            "sales_captured_in_db_pct": sales_captured_pct,
            "source_file": source_file,
        }
        rows.append(row_data)

    return rows


def contiguous_numeric_rows(snapshot: SheetSnapshot, x_col: int, y_col: int, anchor_row: int) -> List[int]:
    rows: List[int] = []
    started = False
    for row in range(anchor_row - 1, snapshot.origin_row - 1, -1):
        x_val = snapshot.get(row, x_col)
        y_val = snapshot.get(row, y_col)
        if is_numeric(x_val) and is_numeric(y_val):
            rows.append(row)
            started = True
        elif started:
            break
    rows.reverse()
    return rows


def rounded_signature_value(value: Any) -> Any:
    numeric = coerce_float(value)
    if numeric is None:
        return value
    return round(numeric, 10)


def extract_regression_rows(
    wb: xw.Book,
    file_label: FileLabel,
    source_file: str,
) -> List[Dict[str, Any]]:
    sheet = locate_sheet(wb, "Regression Model")
    if sheet is None:
        return []

    snapshot = snapshot_sheet(sheet)
    anchor = find_anchor_max(snapshot)
    if anchor is None:
        return []
    anchor_row, anchor_col = anchor

    offsets = detect_offsets(
        snapshot=snapshot,
        anchor_row=anchor_row,
        anchor_col=anchor_col,
        alias_map=REGRESSION_ALIAS_MAP,
        fallback_offsets=REGRESSION_FALLBACK_OFFSETS,
    )

    x_col = anchor_col - 11
    y_col = anchor_col - 7
    history_rows = contiguous_numeric_rows(snapshot, x_col=x_col, y_col=y_col, anchor_row=anchor_row)

    data_start_row = anchor_row + 1
    data_end_row = data_start_row + N_QUARTERS - 1
    cols_needed = [anchor_col + offset for offset in offsets.values()]
    data_start_col = min(cols_needed)
    data_end_col = max(cols_needed)
    data_block = read_block(sheet, data_start_row, data_end_row, data_start_col, data_end_col)

    scratch_col = max(snapshot.end_col + 3, data_end_col + 3)
    scratch_start_row = data_start_row
    wrote_any_formula = False

    for idx in range(N_QUARTERS):
        n_used = idx + 1
        if len(history_rows) < n_used:
            sheet.range((scratch_start_row + idx, scratch_col)).value = None
            sheet.range((scratch_start_row + idx, scratch_col + 1)).value = None
            continue
        start_row = history_rows[len(history_rows) - n_used]
        end_row = history_rows[-1]
        intercept_formula = (
            f'=IFERROR(INTERCEPT(R{start_row}C{y_col}:R{end_row}C{y_col},'
            f'R{start_row}C{x_col}:R{end_row}C{x_col}),"")'
        )
        slope_formula = (
            f'=IFERROR(SLOPE(R{start_row}C{y_col}:R{end_row}C{y_col},'
            f'R{start_row}C{x_col}:R{end_row}C{x_col}),"")'
        )
        set_formula2(sheet.range((scratch_start_row + idx, scratch_col)), intercept_formula)
        set_formula2(sheet.range((scratch_start_row + idx, scratch_col + 1)), slope_formula)
        wrote_any_formula = True

    intercept_values: List[Any] = [None] * N_QUARTERS
    slope_values: List[Any] = [None] * N_QUARTERS
    if wrote_any_formula:
        wb.app.calculate()
        calc_block = read_block(
            sheet,
            scratch_start_row,
            scratch_start_row + N_QUARTERS - 1,
            scratch_col,
            scratch_col + 1,
        )
        for idx in range(N_QUARTERS):
            row_num = scratch_start_row + idx
            intercept_values[idx] = block_value(
                calc_block, row_num, scratch_col, scratch_start_row, scratch_col
            )
            slope_values[idx] = block_value(
                calc_block, row_num, scratch_col + 1, scratch_start_row, scratch_col
            )

    rows: List[Dict[str, Any]] = []
    previous_signature: Optional[Tuple[Any, ...]] = None

    for idx in range(N_QUARTERS):
        row_num = data_start_row + idx
        num_quarters = block_value(
            data_block,
            row_num,
            anchor_col + offsets["num_quarters_used"],
            data_start_row,
            data_start_col,
        )
        if not is_numeric(num_quarters):
            num_quarters = idx + 1

        forecast_value = block_value(
            data_block,
            row_num,
            anchor_col + offsets["forecast_value"],
            data_start_row,
            data_start_col,
        )
        actual_value = block_value(
            data_block,
            row_num,
            anchor_col + offsets["actual_value"],
            data_start_row,
            data_start_col,
        )
        forecast_max = block_value(
            data_block,
            row_num,
            anchor_col + offsets["forecast_max"],
            data_start_row,
            data_start_col,
        )
        forecast_min = block_value(
            data_block,
            row_num,
            anchor_col + offsets["forecast_min"],
            data_start_row,
            data_start_col,
        )
        range_width = safe_subtract(forecast_max, forecast_min)
        intercept = intercept_values[idx] if idx < len(intercept_values) else None
        slope = slope_values[idx] if idx < len(slope_values) else None

        signature = (
            rounded_signature_value(num_quarters),
            rounded_signature_value(forecast_value),
            rounded_signature_value(forecast_max),
            rounded_signature_value(forecast_min),
            rounded_signature_value(intercept),
            rounded_signature_value(slope),
        )
        if previous_signature is not None and signature == previous_signature:
            continue
        previous_signature = signature

        row_data = {
            "model": file_label.model,
            "ticker": file_label.ticker,
            "model_period": file_label.model_period,
            "model_date": file_label.model_date,
            "method": "regression",
            "parameter_name": "num_quarters_used",
            "parameter_value": num_quarters,
            "num_quarters_used": num_quarters,
            "forecast_value": forecast_value,
            "actual_value": actual_value,
            "forecast_max": forecast_max,
            "forecast_min": forecast_min,
            "range_width": range_width,
            "intercept": intercept,
            "slope": slope,
            "source_file": source_file,
        }
        rows.append(row_data)

    return rows


def write_sheet(ws: Worksheet, columns: Sequence[str], rows: Iterable[Dict[str, Any]]) -> None:
    ws.append(list(columns))
    for item in rows:
        ws.append([item.get(column) for column in columns])

    for cell in ws[1]:
        cell.font = Font(bold=True)

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions

    for idx, column_name in enumerate(columns, start=1):
        max_length = len(column_name)
        for row in ws.iter_rows(min_row=2, min_col=idx, max_col=idx):
            value = row[0].value
            if value is None:
                continue
            length = len(str(value))
            if length > max_length:
                max_length = length
        ws.column_dimensions[get_column_letter(idx)].width = min(max(max_length + 2, 12), 48)


def write_output_workbook(
    output_path: Path,
    empirical_rows: List[Dict[str, Any]],
    regression_rows: List[Dict[str, Any]],
) -> None:
    wb = Workbook()
    ws_empirical = wb.active
    ws_empirical.title = "empirical_candidates"
    ws_regression = wb.create_sheet("regression_candidates")

    write_sheet(ws_empirical, EMPIRICAL_COLUMNS, empirical_rows)
    write_sheet(ws_regression, REGRESSION_COLUMNS, regression_rows)

    wb.save(output_path)


def iter_input_files(in_dir: Path) -> Iterable[Path]:
    for path in sorted(in_dir.iterdir(), key=lambda p: p.name.lower()):
        if not path.is_file():
            continue
        yield path


def main() -> int:
    if xw is None:
        print("SKIP all files: xlwings is not installed")
        return 1
    if Workbook is None or Font is None or get_column_letter is None:
        print("SKIP all files: openpyxl is not installed")
        return 1

    in_dir = input_dir
    out_dir = output_dir

    if not in_dir.exists():
        print(f"SKIP all files: input directory does not exist -> {in_dir}")
        return 1

    output_path = get_output_path(in_dir, out_dir)

    empirical_rows: List[Dict[str, Any]] = []
    regression_rows: List[Dict[str, Any]] = []
    files_processed = 0

    app: Optional[xw.App] = None
    try:
        app = xw.App(visible=False, add_book=False)
        app.display_alerts = False
        app.screen_updating = False
        try:
            app.enable_events = False
        except Exception:
            pass
        try:
            app.api.Calculation = -4135  # xlCalculationManual
        except Exception:
            pass

        for file_path in iter_input_files(in_dir):
            file_name = file_path.name

            if file_name.startswith("~"):
                print(f"SKIP {file_name}: temporary file")
                continue
            if file_path.suffix.lower() != ".xlsx":
                print(f"SKIP {file_name}: not an .xlsx file")
                continue

            file_label = parse_file_label(file_name)
            if file_label is None:
                print(f"SKIP {file_name}: could not parse ticker/model period from filename")
                continue

            print(f"PROCESS {file_name}")
            wb_source: Optional[xw.Book] = None
            try:
                wb_source = app.books.open(str(file_path), update_links=False)
                empirical_rows.extend(
                    extract_empirical_rows(
                        wb=wb_source,
                        file_label=file_label,
                        source_file=file_name,
                    )
                )
                regression_rows.extend(
                    extract_regression_rows(
                        wb=wb_source,
                        file_label=file_label,
                        source_file=file_name,
                    )
                )
                files_processed += 1
            except Exception as exc:
                print(f"SKIP {file_name}: processing failed ({exc})")
            finally:
                if wb_source is not None:
                    safe_close_book(wb_source)
    finally:
        if app is not None:
            try:
                app.display_alerts = True
                app.screen_updating = True
                app.quit()
            except Exception:
                pass

    write_output_workbook(output_path, empirical_rows, regression_rows)

    print(f"OUTPUT {output_path}")
    print(f"FILES_PROCESSED {files_processed}")
    print(f"EMPIRICAL_ROWS {len(empirical_rows)}")
    print(f"REGRESSION_ROWS {len(regression_rows)}")

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
