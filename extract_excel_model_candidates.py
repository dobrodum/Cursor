#!/usr/bin/env python3
"""Extract empirical/regression model candidates from .xlsx workbooks.

This script opens each source workbook once with xlwings, processes both
"Empirical Model" and "Regression Model" sheets while the workbook is open,
and writes one consolidated output workbook with:
  - empirical_candidates
  - regression_candidates
"""

from __future__ import annotations

import re
from dataclasses import dataclass
from datetime import date
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Sequence, Tuple

from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

try:
    import xlwings as xw
except ImportError as exc:  # pragma: no cover - environment-specific dependency
    raise SystemExit(
        "xlwings is required for this script. Install with: pip install xlwings"
    ) from exc


# --- User-configurable paths ---
input_dir = Path("./input")
output_dir = Path("./output")


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

# Header matching terms are intentionally broad to support slightly different
# workbook label naming. Values are normalized before matching.
HEADER_SYNONYMS = {
    "num_quarters_used": [
        "num quarters used",
        "quarters used",
        "number of quarters",
        "num quarters",
        "n quarters",
        "n qtrs",
        "qtrs",
        "quarters",
    ],
    "last_quarter_used": ["last quarter used", "last qtr used", "last quarter"],
    "forecast_value_empirical": [
        "estimated total sold",
        "est total sold",
        "forecast value",
        "forecast",
    ],
    "forecast_value_regression": [
        "tot fcst w/o sa",
        "tot fcst without sa",
        "total forecast without sa",
        "forecast total without sa",
    ],
    "actual_value": ["actual value", "actual sales", "reported sales"],
    "forecast_max": ["max", "forecast max"],
    "forecast_min": ["min", "forecast min"],
    "avg_penetration_pct": [
        "avg penetration pct",
        "average penetration pct",
        "avg penetration",
        "average penetration",
    ],
    "quarterly_sales": ["quarterly sales", "qtr sales", "sales quarter"],
    "reported_sales": ["reported sales", "sales reported"],
    "growth_rate_pct": ["growth rate pct", "growth pct", "growth rate"],
    "sales_captured_in_db_pct": [
        "sales captured in db pct",
        "sales captured in db",
        "captured in db pct",
        "db capture pct",
    ],
    "intercept": ["intercept"],
    "slope": ["slope"],
}

# Fallback offsets relative to "max" anchor column.
# These are used only when headers cannot be matched.
EMPIRICAL_FALLBACK_OFFSETS = {
    "num_quarters_used": -11,
    "last_quarter_used": -10,
    "forecast_value": -2,
    "actual_value": -1,
    "forecast_max": 0,
    "forecast_min": 1,
    "avg_penetration_pct": -6,
    "quarterly_sales": -8,
    "reported_sales": -1,
    "growth_rate_pct": -5,
    "sales_captured_in_db_pct": -4,
}

REGRESSION_FALLBACK_OFFSETS = {
    "num_quarters_used": -12,
    "forecast_value": -2,
    "forecast_max": 0,
    "forecast_min": 1,
    "intercept": -4,
    "slope": -3,
}

MONTH_LOOKUP = {
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

DAY_BY_PERIOD_PREFIX = {"early": 5, "mid": 15, "late": 25}


@dataclass(frozen=True)
class ModelInfo:
    model: str
    ticker: str
    model_period: str
    model_date: str


@dataclass(frozen=True)
class SheetSnapshot:
    values: List[List[Any]]
    start_row: int
    start_col: int
    end_row: int
    end_col: int


def normalize_text(value: Any) -> str:
    if value is None:
        return ""
    text = str(value).strip().lower()
    text = re.sub(r"\s+", " ", text)
    text = re.sub(r"[^a-z0-9%/ ]", "", text)
    return text.strip()


def as_float(value: Any) -> Optional[float]:
    if value is None:
        return None
    if isinstance(value, (int, float)):
        return float(value)
    text = str(value).strip().replace(",", "")
    if text == "":
        return None
    try:
        return float(text)
    except ValueError:
        return None


def as_int(value: Any) -> Optional[int]:
    numeric = as_float(value)
    if numeric is None:
        return None
    return int(round(numeric))


def get_model_info_from_filename(file_name: str) -> Optional[ModelInfo]:
    # Expected example: MedMiner_Model - AORT - MidJan2026_Send.xlsx
    pattern = re.compile(
        r"^.+?\s*-\s*([A-Za-z0-9]+)\s*-\s*([A-Za-z]+)(\d{4})_Send\.xlsx$",
        re.IGNORECASE,
    )
    match = pattern.match(file_name)
    if not match:
        return None

    ticker = match.group(1).upper()
    period_token = match.group(2)
    year = int(match.group(3))

    period_match = re.match(r"^(Early|Mid|Late)([A-Za-z]+)$", period_token, re.IGNORECASE)
    if not period_match:
        return None

    period_prefix = period_match.group(1).lower()
    month_token = period_match.group(2)
    month_number = MONTH_LOOKUP.get(month_token.lower())
    if month_number is None:
        # Accept three-letter month prefix as fallback.
        month_number = MONTH_LOOKUP.get(month_token[:3].lower())
    if month_number is None:
        return None

    day = DAY_BY_PERIOD_PREFIX[period_prefix]
    model_date = date(year, month_number, day).isoformat()
    period_prefix_title = period_prefix.capitalize()
    month_abbr = date(year, month_number, 1).strftime("%b")
    model_period = f"{period_prefix_title}{month_abbr}_{year}"
    model = f"{ticker}_{model_period}"
    return ModelInfo(
        model=model,
        ticker=ticker,
        model_period=model_period,
        model_date=model_date,
    )


def build_output_path(source_input_dir: Path, source_output_dir: Path) -> Path:
    source_output_dir.mkdir(parents=True, exist_ok=True)
    stem = f"{source_input_dir.name}_PARAM"
    path = source_output_dir / f"{stem}.xlsx"
    suffix_number = 1
    while path.exists():
        path = source_output_dir / f"{stem}.{suffix_number}.xlsx"
        suffix_number += 1
    return path


def should_skip_file(file_path: Path) -> Optional[str]:
    if not file_path.is_file():
        return "not a file"
    if file_path.name.startswith("~"):
        return "temporary workbook (~ prefix)"
    if file_path.suffix.lower() != ".xlsx":
        return "not .xlsx"
    return None


def capture_sheet_snapshot(sheet: xw.Sheet) -> SheetSnapshot:
    used = sheet.used_range
    raw = used.value
    if raw is None:
        values: List[List[Any]] = []
    elif isinstance(raw, list):
        if raw and not isinstance(raw[0], list):
            values = [raw]
        else:
            values = raw
    else:
        values = [[raw]]
    row_count = len(values)
    col_count = max((len(row) for row in values), default=0)
    normalized_rows: List[List[Any]] = []
    for row in values:
        padded = list(row) + [None] * (col_count - len(row))
        normalized_rows.append(padded)
    start_row = used.row
    start_col = used.column
    end_row = start_row + max(row_count - 1, 0)
    end_col = start_col + max(col_count - 1, 0)
    return SheetSnapshot(
        values=normalized_rows,
        start_row=start_row,
        start_col=start_col,
        end_row=end_row,
        end_col=end_col,
    )


def iter_snapshot_cells(snapshot: SheetSnapshot) -> Iterable[Tuple[int, int, Any]]:
    for r_idx, row in enumerate(snapshot.values):
        abs_row = snapshot.start_row + r_idx
        for c_idx, value in enumerate(row):
            abs_col = snapshot.start_col + c_idx
            yield abs_row, abs_col, value


def find_anchor_max(snapshot: SheetSnapshot) -> Optional[Tuple[int, int]]:
    for abs_row, abs_col, value in iter_snapshot_cells(snapshot):
        if normalize_text(value) == "max":
            return abs_row, abs_col
    return None


def get_cell_from_snapshot(snapshot: SheetSnapshot, row: int, col: int) -> Any:
    if row < snapshot.start_row or col < snapshot.start_col:
        return None
    r_idx = row - snapshot.start_row
    c_idx = col - snapshot.start_col
    if r_idx < 0 or c_idx < 0:
        return None
    if r_idx >= len(snapshot.values):
        return None
    row_values = snapshot.values[r_idx]
    if c_idx >= len(row_values):
        return None
    return row_values[c_idx]


def find_header_row(snapshot: SheetSnapshot, anchor_row: int) -> int:
    candidate_rows = range(max(snapshot.start_row, anchor_row - 2), min(snapshot.end_row, anchor_row + 2) + 1)
    best_row = anchor_row
    best_score = -1
    for row in candidate_rows:
        score = 0
        for col in range(snapshot.start_col, snapshot.end_col + 1):
            value = normalize_text(get_cell_from_snapshot(snapshot, row, col))
            if not value:
                continue
            if value in ("max", "min"):
                score += 2
            for synonyms in HEADER_SYNONYMS.values():
                if any(syn in value for syn in synonyms):
                    score += 1
                    break
        if score > best_score:
            best_score = score
            best_row = row
    return best_row


def build_header_map(snapshot: SheetSnapshot, header_row: int) -> Dict[str, int]:
    mapping: Dict[str, int] = {}
    for col in range(snapshot.start_col, snapshot.end_col + 1):
        normalized = normalize_text(get_cell_from_snapshot(snapshot, header_row, col))
        if not normalized:
            continue
        for key, synonyms in HEADER_SYNONYMS.items():
            if normalized == key:
                mapping[key] = col
                continue
            if any(syn in normalized for syn in synonyms):
                mapping[key] = col
    return mapping


def choose_col(
    anchor_col: int,
    header_map: Dict[str, int],
    header_key: str,
    fallback_offsets: Dict[str, int],
    default_offset: int = 0,
) -> int:
    if header_key in header_map:
        return header_map[header_key]
    return anchor_col + fallback_offsets.get(header_key, default_offset)


def set_formula2(cell: xw.Range, formula_r1c1: str) -> None:
    try:
        cell.formula2 = formula_r1c1
    except Exception:
        # Excel engines that do not expose Range.formula2 still usually expose
        # Formula2 on the COM/automation object.
        cell.api.Formula2 = formula_r1c1


def safe_close_workbook(wb: xw.Book) -> None:
    try:
        wb.close(save=False)
        return
    except TypeError:
        pass
    except Exception:
        pass
    try:
        wb.app.display_alerts = False
    except Exception:
        pass
    try:
        wb.api.Close(False)
    except Exception:
        try:
            wb.close()
        except Exception:
            pass


def calculate_if_needed(wb: xw.Book, formulas_written: bool) -> None:
    if not formulas_written:
        return
    wb.app.calculate()


def make_base_row(model_info: ModelInfo, source_file: str, method: str) -> Dict[str, Any]:
    return {
        "model": model_info.model,
        "ticker": model_info.ticker,
        "model_period": model_info.model_period,
        "model_date": model_info.model_date,
        "method": method,
        "source_file": source_file,
    }


def pull_empirical_rows(
    wb: xw.Book,
    sheet: xw.Sheet,
    model_info: ModelInfo,
    source_file: str,
) -> List[Dict[str, Any]]:
    snapshot = capture_sheet_snapshot(sheet)
    anchor = find_anchor_max(snapshot)
    if anchor is None:
        print(f"  skipped empirical extraction: no 'max' anchor in {sheet.name}")
        return []
    anchor_row, anchor_col = anchor
    header_row = find_header_row(snapshot, anchor_row)
    header_map = build_header_map(snapshot, header_row)

    num_q_col = choose_col(anchor_col, header_map, "num_quarters_used", EMPIRICAL_FALLBACK_OFFSETS)
    last_q_col = choose_col(anchor_col, header_map, "last_quarter_used", EMPIRICAL_FALLBACK_OFFSETS)
    forecast_col = choose_col(
        anchor_col,
        header_map,
        "forecast_value_empirical",
        {"forecast_value_empirical": EMPIRICAL_FALLBACK_OFFSETS["forecast_value"]},
    )
    actual_col = choose_col(anchor_col, header_map, "actual_value", EMPIRICAL_FALLBACK_OFFSETS)
    max_col = choose_col(anchor_col, header_map, "forecast_max", EMPIRICAL_FALLBACK_OFFSETS)
    min_col = choose_col(anchor_col, header_map, "forecast_min", EMPIRICAL_FALLBACK_OFFSETS)
    avg_pen_col = choose_col(anchor_col, header_map, "avg_penetration_pct", EMPIRICAL_FALLBACK_OFFSETS)
    quarterly_sales_col = choose_col(anchor_col, header_map, "quarterly_sales", EMPIRICAL_FALLBACK_OFFSETS)
    reported_sales_col = choose_col(anchor_col, header_map, "reported_sales", EMPIRICAL_FALLBACK_OFFSETS)
    growth_col = choose_col(anchor_col, header_map, "growth_rate_pct", EMPIRICAL_FALLBACK_OFFSETS)
    capture_col = choose_col(anchor_col, header_map, "sales_captured_in_db_pct", EMPIRICAL_FALLBACK_OFFSETS)

    # Temporary formula writes to average penetration column for recalculation.
    # Uses R1C1 and formula2; workbook is closed without saving.
    formulas_written = False
    source_for_avg_col = capture_col if capture_col != avg_pen_col else max(avg_pen_col - 1, 1)
    data_start_row = header_row + 1
    data_end_row = min(snapshot.end_row, data_start_row + N_QUARTERS - 1)
    for row in range(data_start_row, data_end_row + 1):
        n_quarters_raw = sheet.cells(row, num_q_col).value
        n_quarters = as_int(n_quarters_raw)
        if not n_quarters or n_quarters <= 0:
            continue
        start_row = max(1, row - n_quarters + 1)
        formula = f"=AVERAGE(R{start_row}C{source_for_avg_col}:R{row}C{source_for_avg_col})"
        set_formula2(sheet.cells(row, avg_pen_col), formula)
        formulas_written = True
    calculate_if_needed(wb, formulas_written)

    rows: List[Dict[str, Any]] = []
    for row in range(data_start_row, data_end_row + 1):
        num_q = as_int(sheet.cells(row, num_q_col).value)
        if num_q is None:
            continue
        base = make_base_row(model_info, source_file, "empirical")
        forecast_max = as_float(sheet.cells(row, max_col).value)
        forecast_min = as_float(sheet.cells(row, min_col).value)
        avg_pen = as_float(sheet.cells(row, avg_pen_col).value)
        actual_value = as_float(sheet.cells(row, actual_col).value)
        reported_sales = as_float(sheet.cells(row, reported_sales_col).value)
        row_dict = {
            **base,
            "parameter_name": "avg_penetration_pct",
            "parameter_value": avg_pen,
            "num_quarters_used": num_q,
            "last_quarter_used": sheet.cells(row, last_q_col).value,
            "forecast_value": as_float(sheet.cells(row, forecast_col).value),
            "actual_value": actual_value,
            "forecast_max": forecast_max,
            "forecast_min": forecast_min,
            "range_width": (forecast_max - forecast_min)
            if forecast_max is not None and forecast_min is not None
            else None,
            "avg_penetration_pct": avg_pen,
            "quarterly_sales": as_float(sheet.cells(row, quarterly_sales_col).value),
            "reported_sales": reported_sales if reported_sales is not None else actual_value,
            "growth_rate_pct": as_float(sheet.cells(row, growth_col).value),
            "sales_captured_in_db_pct": as_float(sheet.cells(row, capture_col).value),
        }
        rows.append(row_dict)

    return rows


def collect_xy_pairs(sheet: xw.Sheet, x_col: int, y_col: int, end_row: int) -> List[Tuple[int, float, float]]:
    pairs: List[Tuple[int, float, float]] = []
    for row in range(1, end_row + 1):
        x_value = as_float(sheet.cells(row, x_col).value)
        y_value = as_float(sheet.cells(row, y_col).value)
        if x_value is None or y_value is None:
            continue
        pairs.append((row, x_value, y_value))
    return pairs


def same_number(a: Any, b: Any, eps: float = 1e-9) -> bool:
    a_num = as_float(a)
    b_num = as_float(b)
    if a_num is None and b_num is None:
        return True
    if a_num is None or b_num is None:
        return False
    return abs(a_num - b_num) <= eps


def is_duplicate_regression_row(new_row: Dict[str, Any], previous_row: Dict[str, Any]) -> bool:
    keys = [
        "num_quarters_used",
        "forecast_value",
        "forecast_max",
        "forecast_min",
        "intercept",
        "slope",
    ]
    return all(same_number(new_row.get(key), previous_row.get(key)) for key in keys)


def pull_regression_rows(
    wb: xw.Book,
    sheet: xw.Sheet,
    model_info: ModelInfo,
    source_file: str,
) -> List[Dict[str, Any]]:
    snapshot = capture_sheet_snapshot(sheet)
    anchor = find_anchor_max(snapshot)
    if anchor is None:
        print(f"  skipped regression extraction: no 'max' anchor in {sheet.name}")
        return []
    anchor_row, anchor_col = anchor
    header_row = find_header_row(snapshot, anchor_row)
    header_map = build_header_map(snapshot, header_row)

    # Required by instruction.
    y_col = anchor_col - 7
    x_col = anchor_col - 11

    num_q_col = choose_col(anchor_col, header_map, "num_quarters_used", REGRESSION_FALLBACK_OFFSETS)
    forecast_col = choose_col(
        anchor_col,
        header_map,
        "forecast_value_regression",
        {"forecast_value_regression": REGRESSION_FALLBACK_OFFSETS["forecast_value"]},
    )
    max_col = choose_col(anchor_col, header_map, "forecast_max", REGRESSION_FALLBACK_OFFSETS)
    min_col = choose_col(anchor_col, header_map, "forecast_min", REGRESSION_FALLBACK_OFFSETS)
    actual_col = header_map.get("actual_value")

    helper_col_start = max(snapshot.end_col + 2, anchor_col + 2)
    intercept_col = header_map.get(
        "intercept",
        anchor_col + REGRESSION_FALLBACK_OFFSETS["intercept"],
    )
    slope_col = header_map.get(
        "slope",
        anchor_col + REGRESSION_FALLBACK_OFFSETS["slope"],
    )
    if intercept_col == slope_col:
        intercept_col = helper_col_start
        slope_col = helper_col_start + 1

    pairs = collect_xy_pairs(sheet, x_col, y_col, anchor_row - 1)
    if not pairs:
        print(f"  skipped regression extraction: no numeric X/Y pairs before anchor in {sheet.name}")
        return []

    pairs = pairs[-N_QUARTERS:]
    formulas_written = False
    data_start_row = header_row + 1
    for idx, (row_idx, _, _) in enumerate(pairs, start=0):
        target_row = data_start_row + idx
        n_quarters = idx + 1
        start_row = pairs[max(0, len(pairs) - n_quarters)][0]
        end_row = pairs[-1][0]
        intercept_formula = (
            f"=INTERCEPT(R{start_row}C{y_col}:R{end_row}C{y_col},"
            f"R{start_row}C{x_col}:R{end_row}C{x_col})"
        )
        slope_formula = (
            f"=SLOPE(R{start_row}C{y_col}:R{end_row}C{y_col},"
            f"R{start_row}C{x_col}:R{end_row}C{x_col})"
        )
        set_formula2(sheet.cells(target_row, intercept_col), intercept_formula)
        set_formula2(sheet.cells(target_row, slope_col), slope_formula)
        formulas_written = True
        if sheet.cells(target_row, num_q_col).value is None:
            sheet.cells(target_row, num_q_col).value = n_quarters
    calculate_if_needed(wb, formulas_written)

    rows: List[Dict[str, Any]] = []
    max_rows = min(N_QUARTERS, len(pairs))
    for idx in range(max_rows):
        row = data_start_row + idx
        num_q = as_int(sheet.cells(row, num_q_col).value)
        if num_q is None:
            num_q = idx + 1
        forecast_max = as_float(sheet.cells(row, max_col).value)
        forecast_min = as_float(sheet.cells(row, min_col).value)
        actual_value = as_float(sheet.cells(row, actual_col).value) if actual_col else None
        base = make_base_row(model_info, source_file, "regression")
        row_dict = {
            **base,
            "parameter_name": "num_quarters_used",
            "parameter_value": num_q,
            "num_quarters_used": num_q,
            "forecast_value": as_float(sheet.cells(row, forecast_col).value),
            "actual_value": actual_value,
            "forecast_max": forecast_max,
            "forecast_min": forecast_min,
            "range_width": (forecast_max - forecast_min)
            if forecast_max is not None and forecast_min is not None
            else None,
            "intercept": as_float(sheet.cells(row, intercept_col).value),
            "slope": as_float(sheet.cells(row, slope_col).value),
        }
        if rows and idx == max_rows - 1 and is_duplicate_regression_row(row_dict, rows[-1]):
            continue
        rows.append(row_dict)
    return rows


def write_output_sheet(ws, columns: Sequence[str], rows: Sequence[Dict[str, Any]]) -> None:
    ws.append(list(columns))
    for item in rows:
        ws.append([item.get(column) for column in columns])

    for cell in ws[1]:
        cell.font = Font(bold=True)
    ws.freeze_panes = "A2"

    last_col_letter = get_column_letter(len(columns))
    last_row = max(1, ws.max_row)
    ws.auto_filter.ref = f"A1:{last_col_letter}{last_row}"

    for col_idx, column_name in enumerate(columns, start=1):
        max_len = len(column_name)
        for row_idx in range(2, ws.max_row + 1):
            cell_value = ws.cell(row=row_idx, column=col_idx).value
            if cell_value is None:
                continue
            value_len = len(str(cell_value))
            if value_len > max_len:
                max_len = value_len
        ws.column_dimensions[get_column_letter(col_idx)].width = min(max(max_len + 2, 12), 60)


def write_output_workbook(
    output_path: Path,
    empirical_rows: Sequence[Dict[str, Any]],
    regression_rows: Sequence[Dict[str, Any]],
) -> None:
    wb = Workbook()
    empirical_ws = wb.active
    empirical_ws.title = "empirical_candidates"
    regression_ws = wb.create_sheet("regression_candidates")

    write_output_sheet(empirical_ws, EMPIRICAL_COLUMNS, empirical_rows)
    write_output_sheet(regression_ws, REGRESSION_COLUMNS, regression_rows)

    wb.save(output_path)


def iter_xlsx_files(source_input_dir: Path) -> Iterable[Path]:
    for path in sorted(source_input_dir.iterdir(), key=lambda p: p.name.lower()):
        yield path


def get_sheet(wb: xw.Book, name: str) -> Optional[xw.Sheet]:
    try:
        return wb.sheets[name]
    except (KeyError, IndexError):
        return None


def main() -> None:
    if not input_dir.exists():
        raise SystemExit(f"Input directory not found: {input_dir}")

    output_path = build_output_path(input_dir, output_dir)
    empirical_rows: List[Dict[str, Any]] = []
    regression_rows: List[Dict[str, Any]] = []
    processed_files = 0

    app = xw.App(visible=False, add_book=False)
    app.display_alerts = False
    app.screen_updating = False
    try:
        try:
            # Manual calculation for speed; explicit calculate after formula updates.
            app.api.Calculation = -4135  # xlCalculationManual
        except Exception:
            pass

        for file_path in iter_xlsx_files(input_dir):
            skip_reason = should_skip_file(file_path)
            if skip_reason:
                print(f"Skipped: {file_path.name} ({skip_reason})")
                continue

            model_info = get_model_info_from_filename(file_path.name)
            if model_info is None:
                print(f"Skipped: {file_path.name} (filename format not recognized)")
                continue

            print(f"Processing: {file_path.name}")
            wb: Optional[xw.Book] = None
            try:
                wb = app.books.open(str(file_path), update_links=False)
                empirical_sheet = get_sheet(wb, "Empirical Model")
                if empirical_sheet is None:
                    print("  skipped empirical extraction: 'Empirical Model' sheet missing")
                else:
                    empirical_rows.extend(
                        pull_empirical_rows(wb, empirical_sheet, model_info, file_path.name)
                    )

                regression_sheet = get_sheet(wb, "Regression Model")
                if regression_sheet is None:
                    print("  skipped regression extraction: 'Regression Model' sheet missing")
                else:
                    regression_rows.extend(
                        pull_regression_rows(wb, regression_sheet, model_info, file_path.name)
                    )
                processed_files += 1
            except Exception as exc:
                print(f"  skipped workbook due to error: {exc}")
            finally:
                if wb is not None:
                    safe_close_workbook(wb)
    finally:
        app.quit()

    write_output_workbook(output_path, empirical_rows, regression_rows)

    print(f"Output path: {output_path}")
    print(f"Files processed: {processed_files}")
    print(f"Empirical rows: {len(empirical_rows)}")
    print(f"Regression rows: {len(regression_rows)}")


if __name__ == "__main__":
    main()
