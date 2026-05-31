#!/usr/bin/env python3
"""Extract empirical/regression candidate rows from model workbooks.

This script opens each source workbook exactly once, processes both
"Empirical Model" and "Regression Model" sheets while the workbook is open,
and writes one output workbook with:
  - empirical_candidates
  - regression_candidates
"""

from __future__ import annotations

import re
from dataclasses import dataclass
from datetime import date
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Sequence, Tuple

import xlwings as xw
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

# ----------------------------
# User-configurable directories
# ----------------------------
input_dir = Path("./input")
output_dir = Path("./output")


EMPIRICAL_COLUMNS: List[str] = [
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

REGRESSION_COLUMNS: List[str] = [
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

# Anchor-based offsets (from "max" cell).
# These keep layout-specific assumptions in one place for easier tuning.
EMPIRICAL_OFFSETS: Dict[str, int] = {
    "last_quarter_used": -10,
    "sales_captured_in_db_pct": -9,
    "avg_penetration_source": -8,
    "growth_rate_pct": -7,
    "quarterly_sales": -6,
    "reported_sales": -5,
    "estimated_total_sold": -1,
    "forecast_max": 0,
    "forecast_min": 1,
}

MONTHS = {
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

PERIOD_DAY = {"early": 5, "mid": 15, "late": 25}


@dataclass(frozen=True)
class FileLabel:
    model: str
    ticker: str
    model_period: str
    model_date: str


def to_number(value: Any) -> Optional[float]:
    """Convert Excel values to float when possible."""
    if value is None or value == "":
        return None
    if isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        return float(value)
    if isinstance(value, str):
        text = value.strip().replace(",", "")
        if not text:
            return None
        if text.endswith("%"):
            text = text[:-1]
            try:
                return float(text) / 100.0
            except ValueError:
                return None
        try:
            return float(text)
        except ValueError:
            return None
    return None


def as_matrix(values: Any) -> List[List[Any]]:
    """Normalize xlwings values to 2D matrix."""
    if values is None:
        return []
    if not isinstance(values, list):
        return [[values]]
    if values and not isinstance(values[0], list):
        return [values]
    return values


def parse_file_label(file_path: Path) -> FileLabel:
    """Parse ticker/model period/date from file names like:
    MedMiner_Model - AORT - MidJan2026_Send.xlsx
    """
    stem = file_path.stem
    parts = [part.strip() for part in stem.split(" - ")]

    ticker = parts[1].upper() if len(parts) >= 2 and parts[1].strip() else "UNKNOWN"
    period_token = parts[2] if len(parts) >= 3 else ""
    period_token = re.sub(r"_?send.*$", "", period_token, flags=re.IGNORECASE).strip()

    # Supports MidJan2026, MidJanuary2026, Mid Jan 2026.
    match = re.search(
        r"(Early|Mid|Late)\s*([A-Za-z]+)\s*(\d{4})",
        period_token,
        flags=re.IGNORECASE,
    )

    if match:
        period_tag = match.group(1).capitalize()
        month_token = match.group(2).strip()
        month_key = month_token[:3].lower()
        year = int(match.group(3))

        month_num = MONTHS.get(month_key)
        if month_num is not None:
            model_period = f"{period_tag}{month_token[:3].title()}_{year}"
            model_day = PERIOD_DAY[period_tag.lower()]
            model_date = date(year, month_num, model_day).isoformat()
        else:
            model_period = period_token.replace(" ", "_")
            model_date = ""
    else:
        model_period = period_token.replace(" ", "_") if period_token else "UNKNOWN"
        model_date = ""

    model = f"{ticker}_{model_period}" if model_period else ticker
    return FileLabel(model=model, ticker=ticker, model_period=model_period, model_date=model_date)


def output_path_for_run(source_input_dir: Path, destination_dir: Path) -> Path:
    """Create unique output path:
    inputFolder_PARAM.xlsx, inputFolder_PARAM.1.xlsx, ...
    """
    destination_dir.mkdir(parents=True, exist_ok=True)
    folder_name = source_input_dir.name or "input"
    base = destination_dir / f"{folder_name}_PARAM.xlsx"
    if not base.exists():
        return base

    suffix = 1
    while True:
        candidate = destination_dir / f"{folder_name}_PARAM.{suffix}.xlsx"
        if not candidate.exists():
            return candidate
        suffix += 1


def discover_files(source_input_dir: Path) -> Iterable[Tuple[Optional[Path], str]]:
    """Yield work items and skip reasons.

    Returns tuples:
      (file_path, "") for processable files
      (None, reason) for global errors
    """
    if not source_input_dir.exists():
        yield None, f"input directory does not exist: {source_input_dir}"
        return
    if not source_input_dir.is_dir():
        yield None, f"input path is not a directory: {source_input_dir}"
        return

    for path in sorted(source_input_dir.iterdir()):
        if not path.is_file():
            continue
        if path.suffix.lower() != ".xlsx":
            print(f"Skipped: {path.name} (not .xlsx)")
            continue
        if path.name.startswith("~"):
            print(f"Skipped: {path.name} (temporary file)")
            continue
        yield path, ""


def read_used_block(sheet: xw.main.Sheet) -> Tuple[int, int, int, int, List[List[Any]]]:
    """Read used range once (for anchor search and bounds)."""
    used = sheet.used_range
    values = as_matrix(used.options(ndim=2).value)
    start_row = used.row
    start_col = used.column
    last_row = used.last_cell.row
    last_col = used.last_cell.column
    return start_row, start_col, last_row, last_col, values


def find_max_anchor(
    start_row: int, start_col: int, values: Sequence[Sequence[Any]]
) -> Optional[Tuple[int, int]]:
    """Locate the first cell containing 'max' (case-insensitive)."""
    for r_idx, row in enumerate(values):
        for c_idx, cell_value in enumerate(row):
            if isinstance(cell_value, str) and cell_value.strip().lower() == "max":
                return (start_row + r_idx, start_col + c_idx)
    return None


def get_col_from_row_block(row_values: Sequence[Any], min_col: int, target_col: int) -> Any:
    idx = target_col - min_col
    if idx < 0 or idx >= len(row_values):
        return None
    return row_values[idx]


def collect_numeric_history_rows(
    sheet: xw.main.Sheet,
    start_row: int,
    end_row: int,
    cols: Sequence[int],
) -> Dict[int, Dict[int, Optional[float]]]:
    """Bulk-read a row block and return numeric values by row/col."""
    if end_row < start_row:
        return {}

    min_col = min(cols)
    max_col = max(cols)
    block = as_matrix(sheet.range((start_row, min_col), (end_row, max_col)).options(ndim=2).value)

    output: Dict[int, Dict[int, Optional[float]]] = {}
    for row_idx, row_values in enumerate(block):
        absolute_row = start_row + row_idx
        row_map: Dict[int, Optional[float]] = {}
        for col in cols:
            row_map[col] = to_number(get_col_from_row_block(row_values, min_col, col))
        output[absolute_row] = row_map
    return output


def process_empirical_sheet(
    wb: xw.main.Book,
    label: FileLabel,
    source_file: str,
) -> List[Dict[str, Any]]:
    try:
        sheet = wb.sheets["Empirical Model"]
    except Exception:
        print(f"Skipped empirical in {source_file}: sheet 'Empirical Model' not found")
        return []

    start_row, start_col, last_row, last_col, used_values = read_used_block(sheet)
    anchor = find_max_anchor(start_row, start_col, used_values)
    if anchor is None:
        print(f"Skipped empirical in {source_file}: 'max' anchor not found")
        return []

    anchor_row, anchor_col = anchor
    n_quarters = 10

    # Historical input columns (anchor-based).
    penetration_col = anchor_col + EMPIRICAL_OFFSETS["avg_penetration_source"]
    quarter_col = anchor_col + EMPIRICAL_OFFSETS["last_quarter_used"]
    captured_col = anchor_col + EMPIRICAL_OFFSETS["sales_captured_in_db_pct"]
    growth_col = anchor_col + EMPIRICAL_OFFSETS["growth_rate_pct"]
    quarterly_sales_col = anchor_col + EMPIRICAL_OFFSETS["quarterly_sales"]
    reported_sales_col = anchor_col + EMPIRICAL_OFFSETS["reported_sales"]

    history_cols = [
        penetration_col,
        quarter_col,
        captured_col,
        growth_col,
        quarterly_sales_col,
        reported_sales_col,
    ]

    history = collect_numeric_history_rows(
        sheet,
        start_row=1,
        end_row=max(anchor_row - 1, 1),
        cols=[penetration_col, captured_col, growth_col, quarterly_sales_col, reported_sales_col],
    )

    # Last quarter text may be non-numeric, read separately in one pass.
    quarter_text_block = as_matrix(
        sheet.range((1, quarter_col), (max(anchor_row - 1, 1), quarter_col)).options(ndim=2).value
    )

    valid_pen_rows = [
        row
        for row, col_map in history.items()
        if col_map.get(penetration_col) is not None
    ]
    if not valid_pen_rows:
        print(f"Skipped empirical in {source_file}: no penetration history found")
        return []

    max_n = min(n_quarters, len(valid_pen_rows))
    latest_history_row = valid_pen_rows[-1]

    # Use temporary formula cells far to the right.
    temp_col = last_col + 2
    temp_start_row = max(last_row + 2, 2)

    # Write AVERAGE formulas using R1C1 .formula2 (one recalc for all rows).
    for idx in range(max_n):
        n_used = idx + 1
        calc_row = temp_start_row + idx
        start_hist_row = valid_pen_rows[-n_used]
        formula = (
            f"=AVERAGE(R{start_hist_row}C{penetration_col}:"
            f"R{latest_history_row}C{penetration_col})"
        )
        sheet.range((calc_row, temp_col)).formula2 = formula

    wb.app.calculate()

    avg_block = as_matrix(
        sheet.range((temp_start_row, temp_col), (temp_start_row + max_n - 1, temp_col)).options(ndim=2).value
    )
    avg_values = [to_number(row[0]) for row in avg_block]

    # Candidate rows are expected beneath anchor.
    min_candidate_offset = min(EMPIRICAL_OFFSETS.values())
    max_candidate_offset = max(EMPIRICAL_OFFSETS.values())
    candidate_min_col = anchor_col + min_candidate_offset
    candidate_max_col = anchor_col + max_candidate_offset
    candidate_block = as_matrix(
        sheet.range(
            (anchor_row + 1, candidate_min_col),
            (anchor_row + max_n, candidate_max_col),
        ).options(ndim=2).value
    )

    def candidate_value(row_values: Sequence[Any], key: str) -> Any:
        col = anchor_col + EMPIRICAL_OFFSETS[key]
        return get_col_from_row_block(row_values, candidate_min_col, col)

    empirical_rows: List[Dict[str, Any]] = []
    for idx in range(max_n):
        n_used = idx + 1
        row_values = candidate_block[idx] if idx < len(candidate_block) else []

        avg_penetration_pct = avg_values[idx] if idx < len(avg_values) else None
        forecast_max = to_number(candidate_value(row_values, "forecast_max"))
        forecast_min = to_number(candidate_value(row_values, "forecast_min"))
        forecast_est = to_number(candidate_value(row_values, "estimated_total_sold"))

        quarterly_sales = to_number(candidate_value(row_values, "quarterly_sales"))
        reported_sales = to_number(candidate_value(row_values, "reported_sales"))
        growth_rate_pct = to_number(candidate_value(row_values, "growth_rate_pct"))
        sales_captured_pct = to_number(candidate_value(row_values, "sales_captured_in_db_pct"))

        if quarterly_sales is None:
            quarterly_sales = history.get(latest_history_row, {}).get(quarterly_sales_col)
        if reported_sales is None:
            reported_sales = history.get(latest_history_row, {}).get(reported_sales_col)
        if growth_rate_pct is None:
            growth_rate_pct = history.get(latest_history_row, {}).get(growth_col)
        if sales_captured_pct is None:
            sales_captured_pct = history.get(latest_history_row, {}).get(captured_col)

        # Last quarter used can be text; fallback to latest historical row.
        candidate_last_q = candidate_value(row_values, "last_quarter_used")
        if candidate_last_q in (None, ""):
            last_q_idx = latest_history_row - 1
            if 0 <= last_q_idx < len(quarter_text_block):
                candidate_last_q = quarter_text_block[last_q_idx][0]

        # Fallback forecast estimate if sheet cell is empty.
        if forecast_est is None and quarterly_sales is not None and avg_penetration_pct not in (None, 0):
            forecast_est = quarterly_sales / avg_penetration_pct
        if forecast_est is None and reported_sales is not None and avg_penetration_pct not in (None, 0):
            forecast_est = reported_sales / avg_penetration_pct

        range_width = None
        if forecast_max is not None and forecast_min is not None:
            range_width = forecast_max - forecast_min

        empirical_rows.append(
            {
                "model": label.model,
                "ticker": label.ticker,
                "model_period": label.model_period,
                "model_date": label.model_date,
                "method": "empirical",
                "parameter_name": "avg_penetration_pct",
                "parameter_value": avg_penetration_pct,
                "num_quarters_used": n_used,
                "last_quarter_used": candidate_last_q,
                "forecast_value": forecast_est,
                "actual_value": reported_sales,
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

    return empirical_rows


def process_regression_sheet(
    wb: xw.main.Book,
    label: FileLabel,
    source_file: str,
) -> List[Dict[str, Any]]:
    try:
        sheet = wb.sheets["Regression Model"]
    except Exception:
        print(f"Skipped regression in {source_file}: sheet 'Regression Model' not found")
        return []

    start_row, start_col, last_row, last_col, used_values = read_used_block(sheet)
    anchor = find_max_anchor(start_row, start_col, used_values)
    if anchor is None:
        print(f"Skipped regression in {source_file}: 'max' anchor not found")
        return []

    anchor_row, anchor_col = anchor
    y_col = anchor_col - 7
    x_col = anchor_col - 11

    if x_col < 1 or y_col < 1:
        print(f"Skipped regression in {source_file}: invalid x/y offsets from anchor")
        return []

    xy_history = collect_numeric_history_rows(
        sheet=sheet,
        start_row=1,
        end_row=max(anchor_row - 1, 1),
        cols=[x_col, y_col],
    )
    valid_xy_rows = [
        row
        for row, vals in xy_history.items()
        if vals.get(x_col) is not None and vals.get(y_col) is not None
    ]

    if len(valid_xy_rows) < 2:
        print(f"Skipped regression in {source_file}: not enough x/y history rows")
        return []

    # Limit rows for speed while still keeping candidate spread.
    max_n = min(20, len(valid_xy_rows))
    n_candidates = list(range(2, max_n + 1))
    if not n_candidates:
        return []

    latest_row = valid_xy_rows[-1]
    next_x = to_number(sheet.range((latest_row + 1, x_col)).value)
    if next_x is None:
        next_x = xy_history[latest_row][x_col]

    temp_col = last_col + 2
    temp_start_row = max(last_row + 2, 2)

    # Temp table layout:
    # c0=n, c1=intercept, c2=slope, c3=forecast, c4=steyx, c5=max, c6=min
    for idx, n_used in enumerate(n_candidates):
        calc_row = temp_start_row + idx
        start_hist_row = valid_xy_rows[-n_used]
        end_hist_row = latest_row
        y_range = f"R{start_hist_row}C{y_col}:R{end_hist_row}C{y_col}"
        x_range = f"R{start_hist_row}C{x_col}:R{end_hist_row}C{x_col}"

        sheet.range((calc_row, temp_col)).value = n_used
        sheet.range((calc_row, temp_col + 1)).formula2 = f"=INTERCEPT({y_range},{x_range})"
        sheet.range((calc_row, temp_col + 2)).formula2 = f"=SLOPE({y_range},{x_range})"
        sheet.range((calc_row, temp_col + 3)).formula2 = f"=RC[-2]+RC[-1]*{next_x}"
        sheet.range((calc_row, temp_col + 4)).formula2 = f"=STEYX({y_range},{x_range})"
        sheet.range((calc_row, temp_col + 5)).formula2 = "=RC[-2]+RC[-1]"
        sheet.range((calc_row, temp_col + 6)).formula2 = "=RC[-3]-RC[-2]"

    wb.app.calculate()

    calc_block = as_matrix(
        sheet.range(
            (temp_start_row, temp_col),
            (temp_start_row + len(n_candidates) - 1, temp_col + 6),
        ).options(ndim=2).value
    )

    regression_rows: List[Dict[str, Any]] = []
    for row in calc_block:
        if not row:
            continue
        n_used = int(to_number(row[0]) or 0)
        intercept = to_number(row[1])
        slope = to_number(row[2])
        forecast_total_wo_sa = to_number(row[3])
        forecast_max = to_number(row[5])
        forecast_min = to_number(row[6])

        if n_used <= 0:
            continue

        range_width = None
        if forecast_max is not None and forecast_min is not None:
            range_width = forecast_max - forecast_min

        regression_rows.append(
            {
                "model": label.model,
                "ticker": label.ticker,
                "model_period": label.model_period,
                "model_date": label.model_date,
                "method": "regression",
                "parameter_name": "num_quarters_used",
                "parameter_value": n_used,
                "num_quarters_used": n_used,
                "forecast_value": forecast_total_wo_sa,
                "actual_value": "",
                "forecast_max": forecast_max,
                "forecast_min": forecast_min,
                "range_width": range_width,
                "intercept": intercept,
                "slope": slope,
                "source_file": source_file,
            }
        )

    # Guard against duplicate terminal row.
    if len(regression_rows) >= 2:
        prev = regression_rows[-2]
        cur = regression_rows[-1]
        signature_cols = ["intercept", "slope", "forecast_value", "forecast_max", "forecast_min"]
        same_terminal = all(
            (
                (prev[k] is None and cur[k] is None)
                or (
                    isinstance(prev[k], (int, float))
                    and isinstance(cur[k], (int, float))
                    and abs(float(prev[k]) - float(cur[k])) < 1e-12
                )
                or prev[k] == cur[k]
            )
            for k in signature_cols
        )
        if same_terminal:
            regression_rows.pop()

    return regression_rows


def safe_close_source_workbook(wb: Optional[xw.main.Book]) -> None:
    """Close source workbook without saving with fallback methods."""
    if wb is None:
        return
    try:
        wb.close(save=False)
        return
    except TypeError:
        pass
    except Exception:
        pass

    try:
        wb.api.Close(SaveChanges=False)
        return
    except Exception:
        pass

    try:
        wb.close()
    except Exception:
        pass


def write_output_sheet(
    ws: Any,
    columns: Sequence[str],
    rows: Sequence[Dict[str, Any]],
) -> None:
    ws.append(list(columns))
    for row in rows:
        ws.append([row.get(col, "") for col in columns])

    # Header styling and usability formatting.
    for cell in ws[1]:
        cell.font = Font(bold=True)

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions

    for col_idx, col_name in enumerate(columns, start=1):
        max_len = len(col_name)
        for row in rows:
            value = row.get(col_name, "")
            txt = "" if value is None else str(value)
            if len(txt) > max_len:
                max_len = len(txt)
        ws.column_dimensions[get_column_letter(col_idx)].width = min(max_len + 2, 40)


def write_output_workbook(
    output_path: Path,
    empirical_rows: Sequence[Dict[str, Any]],
    regression_rows: Sequence[Dict[str, Any]],
) -> None:
    wb = Workbook()
    ws_empirical = wb.active
    ws_empirical.title = "empirical_candidates"
    ws_regression = wb.create_sheet("regression_candidates")

    write_output_sheet(ws_empirical, EMPIRICAL_COLUMNS, empirical_rows)
    write_output_sheet(ws_regression, REGRESSION_COLUMNS, regression_rows)

    wb.save(output_path)


def run() -> None:
    source_input_dir = input_dir.expanduser().resolve()
    destination_dir = output_dir.expanduser().resolve()

    empirical_rows: List[Dict[str, Any]] = []
    regression_rows: List[Dict[str, Any]] = []
    files_processed = 0

    source_files: List[Path] = []
    for path, reason in discover_files(source_input_dir):
        if path is None:
            print(f"Skipped: {reason}")
            continue
        source_files.append(path)

    if not source_files:
        print("No processable .xlsx files found.")
        output_path = output_path_for_run(source_input_dir, destination_dir)
        write_output_workbook(output_path, empirical_rows, regression_rows)
        print(f"Output workbook: {output_path}")
        print("Files processed: 0")
        print("Empirical rows: 0")
        print("Regression rows: 0")
        return

    app = xw.App(visible=False, add_book=False)
    app.display_alerts = False
    app.screen_updating = False

    try:
        for file_path in source_files:
            print(f"Processing: {file_path.name}")
            wb: Optional[xw.main.Book] = None
            try:
                wb = app.books.open(str(file_path), update_links=False)
                label = parse_file_label(file_path)
                empirical_rows.extend(
                    process_empirical_sheet(wb=wb, label=label, source_file=file_path.name)
                )
                regression_rows.extend(
                    process_regression_sheet(wb=wb, label=label, source_file=file_path.name)
                )
                files_processed += 1
            except Exception as exc:
                print(f"Skipped: {file_path.name} (error: {exc})")
            finally:
                safe_close_source_workbook(wb)
    finally:
        app.quit()

    output_path = output_path_for_run(source_input_dir, destination_dir)
    write_output_workbook(output_path, empirical_rows, regression_rows)

    print(f"Output workbook: {output_path}")
    print(f"Files processed: {files_processed}")
    print(f"Empirical rows: {len(empirical_rows)}")
    print(f"Regression rows: {len(regression_rows)}")


if __name__ == "__main__":
    run()
