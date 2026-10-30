#!/usr/bin/env python3
"""
Extract empirical/regression parameter candidates from all .xlsx files in a folder.

The script opens each source workbook exactly once, processes both the
"Empirical Model" and "Regression Model" sheets while it is open, then closes
the workbook without saving any source-file changes.
"""

from __future__ import annotations

import re
from dataclasses import dataclass
from datetime import date
from pathlib import Path
from typing import Any, Dict, List, Optional, Sequence, Tuple

try:
    import xlwings as xw
except ImportError as exc:  # pragma: no cover - runtime dependency guard
    raise SystemExit("Missing dependency: xlwings") from exc

try:
    from openpyxl import Workbook
    from openpyxl.styles import Font
    from openpyxl.utils import get_column_letter
except ImportError as exc:  # pragma: no cover - runtime dependency guard
    raise SystemExit("Missing dependency: openpyxl") from exc


# -------------------------- USER CONFIGURATION ---------------------------
input_dir = "./input"
output_dir = "./output"
# -------------------------------------------------------------------------

EMPIRICAL_SHEET_NAME = "Empirical Model"
REGRESSION_SHEET_NAME = "Regression Model"

EMPIRICAL_COLUMNS: Sequence[str] = (
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
)

REGRESSION_COLUMNS: Sequence[str] = (
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
)

MONTH_MAP = {
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

PERIOD_DAY_MAP = {"early": 5, "mid": 15, "late": 25}

PERIOD_RE = re.compile(
    r"^(?P<phase>Early|Mid|Late)(?P<month>[A-Za-z]{3})(?P<year>\d{4})$",
    re.IGNORECASE,
)

# These relative offsets are from the "max" anchor column on each data row.
FORECAST_VALUE_COL_OFFSET = -2
ACTUAL_VALUE_COL_OFFSET = -1
FORECAST_MAX_COL_OFFSET = 0
FORECAST_MIN_COL_OFFSET = 1
QUARTERLY_SALES_COL_OFFSET = -6
REPORTED_SALES_COL_OFFSET = -5
GROWTH_RATE_COL_OFFSET = -4
SALES_CAPTURED_COL_OFFSET = -3
LAST_QUARTER_COL_OFFSET = -7

N_QUARTERS = 10


@dataclass
class FileLabels:
    model: str
    ticker: str
    model_period: str
    model_date: str


def parse_file_labels(file_name: str) -> FileLabels:
    """
    Example source file:
    MedMiner_Model - AORT - MidJan2026_Send.xlsx

    Parsed:
      ticker       -> AORT
      model_period -> MidJan_2026
      model_date   -> 2026-01-15
      model        -> AORT_MidJan_2026
    """
    stem = Path(file_name).stem
    parts = [p.strip() for p in stem.split(" - ")]

    ticker = "UNKNOWN"
    period_token = ""

    if len(parts) >= 2 and parts[1]:
        ticker = parts[1].upper()
    if len(parts) >= 3 and parts[2]:
        period_token = parts[2].split("_")[0].strip()

    period_match = PERIOD_RE.match(period_token)
    if not period_match:
        model_period = "unknown"
        return FileLabels(
            model=f"{ticker}_{model_period}",
            ticker=ticker,
            model_period=model_period,
            model_date="",
        )

    phase = period_match.group("phase").capitalize()
    month_token = period_match.group("month").capitalize()
    month_number = MONTH_MAP.get(month_token.lower())
    year = int(period_match.group("year"))

    if month_number is None:
        model_period = "unknown"
        return FileLabels(
            model=f"{ticker}_{model_period}",
            ticker=ticker,
            model_period=model_period,
            model_date="",
        )

    day = PERIOD_DAY_MAP[phase.lower()]
    model_period = f"{phase}{month_token}_{year}"
    model_date = date(year, month_number, day).isoformat()
    model = f"{ticker}_{model_period}"
    return FileLabels(model=model, ticker=ticker, model_period=model_period, model_date=model_date)


def to_float(value: Any) -> Optional[float]:
    if value is None or value == "":
        return None
    if isinstance(value, (int, float)):
        return float(value)
    if isinstance(value, str):
        raw = value.strip()
        if not raw:
            return None
        is_pct = "%" in raw
        cleaned = raw.replace(",", "").replace("%", "")
        try:
            out = float(cleaned)
            return out / 100.0 if is_pct else out
        except ValueError:
            return None
    return None


def roundish(value: Optional[float]) -> Optional[float]:
    if value is None:
        return None
    return round(value, 12)


def calculate_range_width(max_value: Optional[float], min_value: Optional[float]) -> Optional[float]:
    if max_value is None or min_value is None:
        return None
    return max_value - min_value


def max_anchor(sheet: xw.Sheet) -> Optional[Tuple[int, int]]:
    """Find the first cell whose text is exactly 'max' (case-insensitive)."""
    used = sheet.used_range
    values = used.value
    if values is None:
        return None

    # Normalize to 2D list-like.
    if not isinstance(values, list):
        values_2d = [[values]]
    elif values and not isinstance(values[0], (list, tuple)):
        values_2d = [values]
    else:
        values_2d = values

    start_row = used.row
    start_col = used.column

    for r_idx, row_values in enumerate(values_2d):
        row_iter = row_values if isinstance(row_values, (list, tuple)) else [row_values]
        for c_idx, cell_value in enumerate(row_iter):
            if isinstance(cell_value, str) and cell_value.strip().lower() == "max":
                return start_row + r_idx, start_col + c_idx
    return None


def get_cell(sheet: xw.Sheet, row: int, col: int) -> Any:
    if row < 1 or col < 1:
        return None
    try:
        return sheet.cells(row, col).value
    except Exception:
        return None


def set_formula2(cell: xw.Range, formula: str) -> None:
    """Use .formula2 when available (preferred), with safe fallback."""
    try:
        cell.formula2 = formula
    except Exception:
        cell.formula = formula


def safe_close_workbook(wb: xw.Book) -> None:
    """
    Close source workbook without saving.
    Some Excel engines differ in close signatures, so we try safe fallbacks.
    """
    for closer in (
        lambda: wb.close(save=False),
        lambda: wb.close(False),
        lambda: wb.api.Close(SaveChanges=False),
        lambda: wb.api.Close(False),
    ):
        try:
            closer()
            return
        except Exception:
            continue


def extract_empirical_rows(
    wb: xw.Book,
    labels: FileLabels,
    source_file: str,
) -> List[Dict[str, Any]]:
    if EMPIRICAL_SHEET_NAME not in [s.name for s in wb.sheets]:
        print(f"  skipped empirical extraction: sheet '{EMPIRICAL_SHEET_NAME}' not found")
        return []

    sheet = wb.sheets[EMPIRICAL_SHEET_NAME]
    anchor = max_anchor(sheet)
    if anchor is None:
        print("  skipped empirical extraction: 'max' anchor not found")
        return []

    anchor_row, anchor_col = anchor
    used_last_col = sheet.used_range.last_cell.column
    scratch_col = max(used_last_col + 2, anchor_col + 10)
    scratch_row = anchor_row

    rows: List[Dict[str, Any]] = []
    for n_quarters in range(1, N_QUARTERS + 1):
        data_row = anchor_row + n_quarters

        # Dynamic avg penetration calculation (R1C1 / formula2).
        penetration_end_col = max(1, anchor_col + SALES_CAPTURED_COL_OFFSET)
        penetration_start_col = max(1, penetration_end_col - n_quarters + 1)
        avg_formula = (
            f"=AVERAGE(R{data_row}C{penetration_start_col}:"
            f"R{data_row}C{penetration_end_col})"
        )
        avg_cell = sheet.cells(scratch_row, scratch_col)
        set_formula2(avg_cell, avg_formula)
        wb.app.calculate()
        avg_penetration_pct = to_float(avg_cell.value)

        forecast_value = to_float(get_cell(sheet, data_row, anchor_col + FORECAST_VALUE_COL_OFFSET))
        actual_value = to_float(get_cell(sheet, data_row, anchor_col + ACTUAL_VALUE_COL_OFFSET))
        forecast_max = to_float(get_cell(sheet, data_row, anchor_col + FORECAST_MAX_COL_OFFSET))
        forecast_min = to_float(get_cell(sheet, data_row, anchor_col + FORECAST_MIN_COL_OFFSET))
        quarterly_sales = to_float(get_cell(sheet, data_row, anchor_col + QUARTERLY_SALES_COL_OFFSET))
        reported_sales = to_float(get_cell(sheet, data_row, anchor_col + REPORTED_SALES_COL_OFFSET))
        growth_rate_pct = to_float(get_cell(sheet, data_row, anchor_col + GROWTH_RATE_COL_OFFSET))
        sales_captured_pct = to_float(get_cell(sheet, data_row, anchor_col + SALES_CAPTURED_COL_OFFSET))
        last_quarter_used = get_cell(sheet, data_row, anchor_col + LAST_QUARTER_COL_OFFSET)
        range_width = calculate_range_width(forecast_max, forecast_min)

        # Skip blank rows while still preserving dense candidate rows when data exists.
        if all(
            value is None
            for value in (
                forecast_value,
                actual_value,
                forecast_max,
                forecast_min,
                avg_penetration_pct,
            )
        ):
            continue

        rows.append(
            {
                "model": labels.model,
                "ticker": labels.ticker,
                "model_period": labels.model_period,
                "model_date": labels.model_date,
                "method": "empirical",
                "parameter_name": "avg_penetration_pct",
                "parameter_value": avg_penetration_pct,
                "num_quarters_used": n_quarters,
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
        )

    return rows


def extract_regression_rows(
    wb: xw.Book,
    labels: FileLabels,
    source_file: str,
) -> List[Dict[str, Any]]:
    if REGRESSION_SHEET_NAME not in [s.name for s in wb.sheets]:
        print(f"  skipped regression extraction: sheet '{REGRESSION_SHEET_NAME}' not found")
        return []

    sheet = wb.sheets[REGRESSION_SHEET_NAME]
    anchor = max_anchor(sheet)
    if anchor is None:
        print("  skipped regression extraction: 'max' anchor not found")
        return []

    anchor_row, anchor_col = anchor
    y_col = anchor_col - 7
    x_col = anchor_col - 11
    data_start_row = anchor_row + 1

    if y_col < 1 or x_col < 1:
        print("  skipped regression extraction: invalid anchor offsets for x/y columns")
        return []

    used_last_col = sheet.used_range.last_cell.column
    scratch_intercept_col = max(used_last_col + 2, anchor_col + 10)
    scratch_slope_col = scratch_intercept_col + 1
    scratch_row = anchor_row

    rows: List[Dict[str, Any]] = []
    previous_signature: Optional[Tuple[Optional[float], ...]] = None

    for n_quarters in range(2, N_QUARTERS + 1):
        end_row = data_start_row + n_quarters - 1

        intercept_formula = (
            f"=INTERCEPT(R{data_start_row}C{y_col}:R{end_row}C{y_col},"
            f"R{data_start_row}C{x_col}:R{end_row}C{x_col})"
        )
        slope_formula = (
            f"=SLOPE(R{data_start_row}C{y_col}:R{end_row}C{y_col},"
            f"R{data_start_row}C{x_col}:R{end_row}C{x_col})"
        )

        intercept_cell = sheet.cells(scratch_row, scratch_intercept_col)
        slope_cell = sheet.cells(scratch_row, scratch_slope_col)
        set_formula2(intercept_cell, intercept_formula)
        set_formula2(slope_cell, slope_formula)
        wb.app.calculate()

        intercept = to_float(intercept_cell.value)
        slope = to_float(slope_cell.value)

        data_row = anchor_row + n_quarters
        forecast_value = to_float(get_cell(sheet, data_row, anchor_col + FORECAST_VALUE_COL_OFFSET))
        actual_value = to_float(get_cell(sheet, data_row, anchor_col + ACTUAL_VALUE_COL_OFFSET))
        forecast_max = to_float(get_cell(sheet, data_row, anchor_col + FORECAST_MAX_COL_OFFSET))
        forecast_min = to_float(get_cell(sheet, data_row, anchor_col + FORECAST_MIN_COL_OFFSET))
        range_width = calculate_range_width(forecast_max, forecast_min)

        signature = (
            roundish(intercept),
            roundish(slope),
            roundish(forecast_value),
            roundish(forecast_max),
            roundish(forecast_min),
        )

        # Prevent duplicate final-row append when values repeat.
        if signature == previous_signature:
            continue
        previous_signature = signature

        if all(value is None for value in signature):
            continue

        rows.append(
            {
                "model": labels.model,
                "ticker": labels.ticker,
                "model_period": labels.model_period,
                "model_date": labels.model_date,
                "method": "regression",
                "parameter_name": "num_quarters_used",
                "parameter_value": n_quarters,
                "num_quarters_used": n_quarters,
                "forecast_value": forecast_value,
                "actual_value": actual_value,
                "forecast_max": forecast_max,
                "forecast_min": forecast_min,
                "range_width": range_width,
                "intercept": intercept,
                "slope": slope,
                "source_file": source_file,
            }
        )

    return rows


def ensure_unique_output_path(input_path: Path, out_dir: Path) -> Path:
    base_name = f"{input_path.name}_PARAM"
    candidate = out_dir / f"{base_name}.xlsx"
    idx = 1
    while candidate.exists():
        candidate = out_dir / f"{base_name}.{idx}.xlsx"
        idx += 1
    return candidate


def write_sheet(
    workbook: Workbook,
    sheet_name: str,
    columns: Sequence[str],
    rows: List[Dict[str, Any]],
) -> None:
    ws = workbook.create_sheet(title=sheet_name)
    ws.append(list(columns))
    for row in rows:
        ws.append([row.get(col) for col in columns])

    # Header formatting / basic usability.
    for cell in ws[1]:
        cell.font = Font(bold=True)
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(ws.max_column)}{max(ws.max_row, 1)}"

    for col_idx, col_name in enumerate(columns, start=1):
        width = max(12, len(col_name) + 2)
        for row_idx in range(2, ws.max_row + 1):
            value = ws.cell(row=row_idx, column=col_idx).value
            if value is not None:
                width = max(width, min(60, len(str(value)) + 2))
        ws.column_dimensions[get_column_letter(col_idx)].width = min(width, 60)


def write_output_workbook(
    output_path: Path,
    empirical_rows: List[Dict[str, Any]],
    regression_rows: List[Dict[str, Any]],
) -> None:
    wb_out = Workbook()
    default_sheet = wb_out.active
    wb_out.remove(default_sheet)
    write_sheet(wb_out, "empirical_candidates", EMPIRICAL_COLUMNS, empirical_rows)
    write_sheet(wb_out, "regression_candidates", REGRESSION_COLUMNS, regression_rows)
    wb_out.save(output_path)


def process_all_files(input_path: Path, out_path: Path) -> Tuple[int, int, int, Path]:
    source_files = sorted(input_path.iterdir())
    empirical_rows: List[Dict[str, Any]] = []
    regression_rows: List[Dict[str, Any]] = []
    processed_files = 0

    with xw.App(visible=False, add_book=False) as app:
        app.display_alerts = False
        app.screen_updating = False
        try:
            app.calculation = "manual"
        except Exception:
            # Some Excel backends do not expose this property.
            pass

        for file_path in source_files:
            if not file_path.is_file():
                continue
            if file_path.name.startswith("~"):
                print(f"skipped {file_path.name}: temp file")
                continue
            if file_path.suffix.lower() != ".xlsx":
                print(f"skipped {file_path.name}: not .xlsx")
                continue

            print(f"processing {file_path.name}")
            workbook: Optional[xw.Book] = None
            try:
                workbook = app.books.open(str(file_path), update_links=False)
                labels = parse_file_labels(file_path.name)
                empirical_rows.extend(extract_empirical_rows(workbook, labels, file_path.name))
                regression_rows.extend(extract_regression_rows(workbook, labels, file_path.name))
                processed_files += 1
                print(f"processed {file_path.name}")
            except Exception as exc:
                print(f"skipped {file_path.name}: processing error ({exc})")
            finally:
                if workbook is not None:
                    safe_close_workbook(workbook)

    output_file = ensure_unique_output_path(input_path, out_path)
    write_output_workbook(output_file, empirical_rows, regression_rows)
    return processed_files, len(empirical_rows), len(regression_rows), output_file


def main() -> None:
    input_path = Path(input_dir).expanduser().resolve()
    out_path = Path(output_dir).expanduser().resolve()

    if not input_path.exists() or not input_path.is_dir():
        raise SystemExit(f"Input folder does not exist or is not a directory: {input_path}")
    out_path.mkdir(parents=True, exist_ok=True)

    processed_files, empirical_count, regression_count, output_file = process_all_files(
        input_path,
        out_path,
    )

    print(f"output path: {output_file}")
    print(f"number of files processed: {processed_files}")
    print(f"number of empirical rows: {empirical_count}")
    print(f"number of regression rows: {regression_count}")


if __name__ == "__main__":
    main()
