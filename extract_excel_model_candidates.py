#!/usr/bin/env python3
"""Extract empirical/regression model candidates from .xlsx workbooks.

This script opens each source workbook exactly once, extracts both model tabs
while the workbook is open, and writes a single output workbook with:
  - empirical_candidates
  - regression_candidates
"""

from __future__ import annotations

import re
from datetime import date
from pathlib import Path
from typing import Any, Dict, List, Optional, Sequence, Tuple

import xlwings as xw
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.worksheet.worksheet import Worksheet

# ---------------------------
# User-configurable locations
# ---------------------------
input_dir = Path("./input")
output_dir = Path("./output")


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


def to_2d(values: Any) -> List[List[Any]]:
    if values is None:
        return []
    if isinstance(values, (list, tuple)):
        if values and isinstance(values[0], (list, tuple)):
            return [list(row) for row in values]
        return [list(values)]
    return [[values]]


def norm_text(value: Any) -> str:
    if value is None:
        return ""
    return re.sub(r"\s+", " ", str(value).strip().lower())


def as_float(value: Any) -> Optional[float]:
    if value is None or value == "":
        return None
    if isinstance(value, (int, float)):
        return float(value)
    text = str(value).strip().replace(",", "")
    if not text:
        return None
    try:
        return float(text)
    except ValueError:
        return None


def as_int(value: Any) -> Optional[int]:
    f_value = as_float(value)
    if f_value is None:
        return None
    return int(round(f_value))


def subtract_if_both(left: Any, right: Any) -> Optional[float]:
    l_val = as_float(left)
    r_val = as_float(right)
    if l_val is None or r_val is None:
        return None
    return l_val - r_val


def close_workbook_safely(wb: xw.Book) -> None:
    try:
        wb.close(save=False)
        return
    except TypeError:
        pass
    except Exception:
        pass

    try:
        wb.api.Close(SaveChanges=False)  # noqa: N802 - Excel COM naming
        return
    except Exception:
        pass

    try:
        wb.close()
    except Exception:
        pass


def month_number(month_token: str) -> Optional[int]:
    month_map = {
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
    return month_map.get(month_token[:3].lower())


def parse_file_labels(file_name: str) -> Dict[str, str]:
    stem = Path(file_name).stem
    parts = [part.strip() for part in stem.split("-")]

    ticker = ""
    if len(parts) >= 2:
        ticker = parts[1].split()[0].upper()
    if not ticker:
        ticker_match = re.search(r"\b([A-Z]{2,6})\b", stem)
        if ticker_match:
            ticker = ticker_match.group(1).upper()

    period_match = re.search(
        r"(Early|Mid|Late)\s*([A-Za-z]{3,9})\s*(\d{4})",
        stem,
        flags=re.IGNORECASE,
    )

    model_period = ""
    model_date = ""
    if period_match:
        phase = period_match.group(1).title()
        month_token = period_match.group(2).title()
        year = period_match.group(3)
        month_num = month_number(month_token)
        if month_num is not None:
            day = {"Early": 5, "Mid": 15, "Late": 25}[phase]
            model_date = date(int(year), month_num, day).isoformat()
            model_period = f"{phase}{month_token[:3]}_{year}"

    model = f"{ticker}_{model_period}" if ticker and model_period else stem
    return {
        "ticker": ticker,
        "model_period": model_period,
        "model_date": model_date,
        "model": model,
    }


def get_unique_output_path(input_path: Path, out_dir: Path) -> Path:
    base_name = f"{input_path.name}_PARAM"
    candidate = out_dir / f"{base_name}.xlsx"
    if not candidate.exists():
        return candidate

    counter = 1
    while True:
        candidate = out_dir / f"{base_name}.{counter}.xlsx"
        if not candidate.exists():
            return candidate
        counter += 1


def read_used_grid(sheet: xw.Sheet) -> Tuple[int, int, List[List[Any]]]:
    used = sheet.used_range
    return used.row, used.column, to_2d(used.value)


def grid_get(
    grid: Sequence[Sequence[Any]],
    start_row: int,
    start_col: int,
    row: int,
    col: int,
) -> Any:
    r_idx = row - start_row
    c_idx = col - start_col
    if r_idx < 0 or c_idx < 0:
        return None
    if r_idx >= len(grid):
        return None
    if c_idx >= len(grid[r_idx]):
        return None
    return grid[r_idx][c_idx]


def find_anchor_max(
    grid: Sequence[Sequence[Any]],
    start_row: int,
    start_col: int,
) -> Optional[Tuple[int, int]]:
    for r_off, row_values in enumerate(grid):
        for c_off, value in enumerate(row_values):
            if norm_text(value) == "max":
                return start_row + r_off, start_col + c_off
    return None


def infer_offsets_by_tokens(
    grid: Sequence[Sequence[Any]],
    start_row: int,
    start_col: int,
    anchor_row: int,
    anchor_col: int,
    token_map: Dict[str, Tuple[str, ...]],
    row_window: int = 1,
    col_window: int = 20,
) -> Dict[str, int]:
    offsets: Dict[str, int] = {}
    row_candidates = [
        row
        for row in range(anchor_row - row_window, anchor_row + row_window + 1)
        if row >= 1
    ]

    for col in range(anchor_col - col_window, anchor_col + col_window + 1):
        if col < 1:
            continue
        for row in row_candidates:
            label = norm_text(grid_get(grid, start_row, start_col, row, col))
            if not label:
                continue
            for key, tokens in token_map.items():
                if key in offsets:
                    continue
                if any(token in label for token in tokens):
                    offsets[key] = col - anchor_col
    return offsets


def collect_numeric_rows_above(
    grid: Sequence[Sequence[Any]],
    start_row: int,
    start_col: int,
    target_col: int,
    anchor_row: int,
    max_scan: int = 500,
) -> List[int]:
    rows: List[int] = []
    scanned = 0
    gap_after_start = 0
    for row in range(anchor_row - 1, 0, -1):
        scanned += 1
        if scanned > max_scan:
            break
        value = grid_get(grid, start_row, start_col, row, target_col)
        if as_float(value) is not None:
            rows.append(row)
            gap_after_start = 0
            continue
        if rows:
            gap_after_start += 1
            if gap_after_start >= 3:
                break

    rows.reverse()
    return rows


def get_sheet(wb: xw.Book, sheet_name: str) -> Optional[xw.Sheet]:
    try:
        return wb.sheets[sheet_name]
    except Exception:
        return None


def extract_empirical_candidates(
    wb: xw.Book,
    sheet: xw.Sheet,
    labels: Dict[str, str],
    source_file: str,
    n_quarters: int = 10,
) -> List[Dict[str, Any]]:
    start_row, start_col, grid = read_used_grid(sheet)
    anchor = find_anchor_max(grid, start_row, start_col)
    if anchor is None:
        return []
    anchor_row, anchor_col = anchor

    token_map = {
        "num_quarters_used": ("num quarter", "quarters used", "n quarter"),
        "last_quarter_used": ("last quarter", "quarter used"),
        "forecast_value": ("estimated total sold", "est total sold", "forecast value", "tot fcst"),
        "actual_value": ("reported sales", "actual sales", "actual value"),
        "forecast_min": ("min",),
        "forecast_max": ("max",),
        "quarterly_sales": ("quarterly sales", "qtr sales"),
        "reported_sales": ("reported sales",),
        "growth_rate_pct": ("growth rate", "growth %", "growth pct"),
        "sales_captured_in_db_pct": ("captured in db", "sales captured", "db pct"),
        "penetration_source": ("penetration",),
    }

    inferred = infer_offsets_by_tokens(
        grid=grid,
        start_row=start_row,
        start_col=start_col,
        anchor_row=anchor_row,
        anchor_col=anchor_col,
        token_map=token_map,
    )

    defaults = {
        "num_quarters_used": -6,
        "last_quarter_used": -5,
        "forecast_value": -3,
        "actual_value": -2,
        "forecast_max": 0,
        "forecast_min": 1,
        "quarterly_sales": -4,
        "reported_sales": -2,
        "growth_rate_pct": 2,
        "sales_captured_in_db_pct": 3,
    }
    offsets = {**defaults, **inferred}

    helper_col = sheet.used_range.last_cell.column + 2
    penetration_col: Optional[int] = None
    if "penetration_source" in offsets:
        penetration_col = anchor_col + offsets["penetration_source"]
    penetration_rows: List[int] = []
    if penetration_col is not None:
        penetration_rows = collect_numeric_rows_above(
            grid=grid,
            start_row=start_row,
            start_col=start_col,
            target_col=penetration_col,
            anchor_row=anchor_row,
        )

    rows: List[Dict[str, Any]] = []
    avg_cells: Dict[int, Tuple[int, int]] = {}

    for i in range(1, n_quarters + 1):
        row = anchor_row + i
        num_q_raw = grid_get(
            grid, start_row, start_col, row, anchor_col + offsets["num_quarters_used"]
        )
        num_quarters_used = as_int(num_q_raw) or i
        if num_quarters_used <= 0:
            num_quarters_used = i

        if penetration_col is not None and penetration_rows and num_quarters_used <= len(penetration_rows):
            start_data_row = penetration_rows[-num_quarters_used]
            end_data_row = penetration_rows[-1]
            avg_formula = (
                f"=AVERAGE(R{start_data_row}C{penetration_col}:R{end_data_row}C{penetration_col})"
            )
            avg_cell = sheet.cells(row, helper_col)
            avg_cell.formula2 = avg_formula
            avg_cells[row] = (row, helper_col)

        forecast_value = grid_get(
            grid, start_row, start_col, row, anchor_col + offsets["forecast_value"]
        )
        actual_value = grid_get(
            grid, start_row, start_col, row, anchor_col + offsets["actual_value"]
        )
        forecast_max = grid_get(
            grid, start_row, start_col, row, anchor_col + offsets["forecast_max"]
        )
        forecast_min = grid_get(
            grid, start_row, start_col, row, anchor_col + offsets["forecast_min"]
        )
        last_quarter_used = grid_get(
            grid, start_row, start_col, row, anchor_col + offsets["last_quarter_used"]
        )
        quarterly_sales = grid_get(
            grid, start_row, start_col, row, anchor_col + offsets["quarterly_sales"]
        )
        reported_sales = grid_get(
            grid, start_row, start_col, row, anchor_col + offsets["reported_sales"]
        )
        growth_rate_pct = grid_get(
            grid, start_row, start_col, row, anchor_col + offsets["growth_rate_pct"]
        )
        sales_captured_in_db_pct = grid_get(
            grid, start_row, start_col, row, anchor_col + offsets["sales_captured_in_db_pct"]
        )

        if (
            forecast_value in (None, "")
            and forecast_max in (None, "")
            and forecast_min in (None, "")
            and actual_value in (None, "")
        ):
            continue

        rows.append(
            {
                "model": labels["model"],
                "ticker": labels["ticker"],
                "model_period": labels["model_period"],
                "model_date": labels["model_date"],
                "method": "empirical",
                "parameter_name": "avg_penetration_pct",
                "parameter_value": None,  # populated after calc
                "num_quarters_used": num_quarters_used,
                "last_quarter_used": last_quarter_used,
                "forecast_value": forecast_value,
                "actual_value": actual_value,
                "forecast_max": forecast_max,
                "forecast_min": forecast_min,
                "range_width": subtract_if_both(forecast_max, forecast_min),
                "avg_penetration_pct": None,  # populated after calc
                "quarterly_sales": quarterly_sales,
                "reported_sales": reported_sales,
                "growth_rate_pct": growth_rate_pct,
                "sales_captured_in_db_pct": sales_captured_in_db_pct,
                "source_file": source_file,
                "_row_index": row,
            }
        )

    if avg_cells:
        wb.app.calculate()

    for row_data in rows:
        row_index = row_data["_row_index"]
        avg_value = None
        if row_index in avg_cells:
            avg_row, avg_col = avg_cells[row_index]
            avg_value = sheet.cells(avg_row, avg_col).value
        if avg_value in (None, ""):
            avg_value = row_data.get("avg_penetration_pct")
        row_data["avg_penetration_pct"] = avg_value
        row_data["parameter_value"] = avg_value
        row_data.pop("_row_index", None)

    return rows


def almost_equal(left: Any, right: Any, tol: float = 1e-9) -> bool:
    lf = as_float(left)
    rf = as_float(right)
    if lf is not None and rf is not None:
        return abs(lf - rf) <= tol
    return (left in (None, "")) and (right in (None, ""))


def extract_regression_candidates(
    wb: xw.Book,
    sheet: xw.Sheet,
    labels: Dict[str, str],
    source_file: str,
    n_quarters: int = 10,
) -> List[Dict[str, Any]]:
    start_row, start_col, grid = read_used_grid(sheet)
    anchor = find_anchor_max(grid, start_row, start_col)
    if anchor is None:
        return []
    anchor_row, anchor_col = anchor

    y_col = anchor_col - 7
    x_col = anchor_col - 11

    token_map = {
        "num_quarters_used": ("num quarter", "quarters used", "n quarter"),
        "forecast_value": ("tot fcst w/o sa", "fcst w/o sa", "without sa"),
        "forecast_min": ("min",),
        "forecast_max": ("max",),
    }
    inferred = infer_offsets_by_tokens(
        grid=grid,
        start_row=start_row,
        start_col=start_col,
        anchor_row=anchor_row,
        anchor_col=anchor_col,
        token_map=token_map,
    )

    defaults = {
        "num_quarters_used": -10,
        "forecast_value": -2,
        "forecast_max": 0,
        "forecast_min": 1,
    }
    offsets = {**defaults, **inferred}

    series_rows = []
    max_scan_rows = collect_numeric_rows_above(
        grid=grid,
        start_row=start_row,
        start_col=start_col,
        target_col=x_col,
        anchor_row=anchor_row,
    )
    for row in max_scan_rows:
        x_value = grid_get(grid, start_row, start_col, row, x_col)
        y_value = grid_get(grid, start_row, start_col, row, y_col)
        if as_float(x_value) is not None and as_float(y_value) is not None:
            series_rows.append(row)

    helper_col = sheet.used_range.last_cell.column + 4
    calc_rows: Dict[int, Tuple[int, int]] = {}

    rows: List[Dict[str, Any]] = []
    for i in range(1, n_quarters + 1):
        row = anchor_row + i
        raw_num = grid_get(
            grid, start_row, start_col, row, anchor_col + offsets["num_quarters_used"]
        )
        num_quarters_used = as_int(raw_num) or i
        if num_quarters_used <= 0:
            num_quarters_used = i

        if len(series_rows) >= 2 and num_quarters_used <= len(series_rows) and num_quarters_used >= 2:
            series_start = series_rows[-num_quarters_used]
            series_end = series_rows[-1]

            intercept_formula = (
                f"=INTERCEPT(R{series_start}C{y_col}:R{series_end}C{y_col},"
                f"R{series_start}C{x_col}:R{series_end}C{x_col})"
            )
            slope_formula = (
                f"=SLOPE(R{series_start}C{y_col}:R{series_end}C{y_col},"
                f"R{series_start}C{x_col}:R{series_end}C{x_col})"
            )
            intercept_cell = sheet.cells(row, helper_col)
            slope_cell = sheet.cells(row, helper_col + 1)
            intercept_cell.formula2 = intercept_formula
            slope_cell.formula2 = slope_formula
            calc_rows[row] = (helper_col, helper_col + 1)

        forecast_value = grid_get(
            grid, start_row, start_col, row, anchor_col + offsets["forecast_value"]
        )
        forecast_max = grid_get(
            grid, start_row, start_col, row, anchor_col + offsets["forecast_max"]
        )
        forecast_min = grid_get(
            grid, start_row, start_col, row, anchor_col + offsets["forecast_min"]
        )

        if forecast_value in (None, "") and forecast_max in (None, "") and forecast_min in (None, ""):
            continue

        rows.append(
            {
                "model": labels["model"],
                "ticker": labels["ticker"],
                "model_period": labels["model_period"],
                "model_date": labels["model_date"],
                "method": "regression",
                "parameter_name": "num_quarters_used",
                "parameter_value": num_quarters_used,
                "num_quarters_used": num_quarters_used,
                "forecast_value": forecast_value,
                "actual_value": "",
                "forecast_max": forecast_max,
                "forecast_min": forecast_min,
                "range_width": subtract_if_both(forecast_max, forecast_min),
                "intercept": None,
                "slope": None,
                "source_file": source_file,
                "_row_index": row,
            }
        )

    if calc_rows:
        wb.app.calculate()

    deduped_rows: List[Dict[str, Any]] = []
    for row_data in rows:
        row_index = row_data["_row_index"]
        if row_index in calc_rows:
            int_col, slp_col = calc_rows[row_index]
            row_data["intercept"] = sheet.cells(row_index, int_col).value
            row_data["slope"] = sheet.cells(row_index, slp_col).value
        row_data.pop("_row_index", None)

        if deduped_rows:
            prev = deduped_rows[-1]
            is_duplicate = all(
                (
                    almost_equal(prev["forecast_value"], row_data["forecast_value"]),
                    almost_equal(prev["forecast_max"], row_data["forecast_max"]),
                    almost_equal(prev["forecast_min"], row_data["forecast_min"]),
                    almost_equal(prev["intercept"], row_data["intercept"]),
                    almost_equal(prev["slope"], row_data["slope"]),
                )
            )
            if is_duplicate:
                continue
        deduped_rows.append(row_data)

    return deduped_rows


def write_table(ws: Worksheet, columns: List[str], rows: List[Dict[str, Any]]) -> None:
    ws.append(columns)
    for item in rows:
        ws.append([item.get(column, "") for column in columns])

    for cell in ws[1]:
        cell.font = Font(bold=True)

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions

    for col_idx, col_name in enumerate(columns, start=1):
        max_len = len(col_name)
        for row_idx in range(2, ws.max_row + 1):
            value = ws.cell(row=row_idx, column=col_idx).value
            if value is None:
                continue
            max_len = max(max_len, len(str(value)))
        ws.column_dimensions[ws.cell(row=1, column=col_idx).column_letter].width = min(max_len + 2, 48)


def write_output_workbook(
    output_path: Path,
    empirical_rows: List[Dict[str, Any]],
    regression_rows: List[Dict[str, Any]],
) -> None:
    out_wb = Workbook()
    out_wb.remove(out_wb.active)
    emp_ws = out_wb.create_sheet("empirical_candidates")
    reg_ws = out_wb.create_sheet("regression_candidates")

    write_table(emp_ws, EMPIRICAL_COLUMNS, empirical_rows)
    write_table(reg_ws, REGRESSION_COLUMNS, regression_rows)
    out_wb.save(output_path)


def process_workbooks() -> None:
    in_dir = Path(input_dir)
    out_dir = Path(output_dir)
    out_dir.mkdir(parents=True, exist_ok=True)

    if not in_dir.exists() or not in_dir.is_dir():
        raise FileNotFoundError(f"input_dir does not exist or is not a folder: {in_dir}")

    output_path = get_unique_output_path(in_dir.resolve(), out_dir.resolve())

    empirical_rows: List[Dict[str, Any]] = []
    regression_rows: List[Dict[str, Any]] = []

    processed_files = 0

    app = xw.App(visible=False, add_book=False)
    app.display_alerts = False
    app.screen_updating = False
    try:
        for file_path in sorted(in_dir.iterdir()):
            if not file_path.is_file():
                continue
            if file_path.name.startswith("~"):
                print(f"SKIP {file_path.name}: temp file starts with '~'")
                continue
            if file_path.suffix.lower() != ".xlsx":
                print(f"SKIP {file_path.name}: not an .xlsx file")
                continue

            print(f"PROCESS {file_path.name}")
            labels = parse_file_labels(file_path.name)
            wb: Optional[xw.Book] = None
            try:
                wb = app.books.open(str(file_path), update_links=False)
                processed_files += 1

                empirical_sheet = get_sheet(wb, "Empirical Model")
                if empirical_sheet is None:
                    print(f"SKIP {file_path.name}: missing sheet 'Empirical Model'")
                else:
                    empirical_rows.extend(
                        extract_empirical_candidates(
                            wb=wb,
                            sheet=empirical_sheet,
                            labels=labels,
                            source_file=file_path.name,
                            n_quarters=10,
                        )
                    )

                regression_sheet = get_sheet(wb, "Regression Model")
                if regression_sheet is None:
                    print(f"SKIP {file_path.name}: missing sheet 'Regression Model'")
                else:
                    regression_rows.extend(
                        extract_regression_candidates(
                            wb=wb,
                            sheet=regression_sheet,
                            labels=labels,
                            source_file=file_path.name,
                            n_quarters=10,
                        )
                    )
            except Exception as exc:
                print(f"SKIP {file_path.name}: processing error: {exc}")
            finally:
                if wb is not None:
                    close_workbook_safely(wb)
    finally:
        app.quit()

    write_output_workbook(output_path, empirical_rows, regression_rows)

    print(f"OUTPUT {output_path}")
    print(f"FILES_PROCESSED {processed_files}")
    print(f"EMPIRICAL_ROWS {len(empirical_rows)}")
    print(f"REGRESSION_ROWS {len(regression_rows)}")


if __name__ == "__main__":
    process_workbooks()
