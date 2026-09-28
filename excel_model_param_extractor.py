#!/usr/bin/env python3
"""Extract empirical and regression candidates from .xlsx model workbooks."""

from __future__ import annotations

import re
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

import xlwings as xw

# Configure these paths before running.
input_dir = Path("input")
output_dir = Path("output")

EMPIRICAL_SHEET_NAME = "Empirical Model"
REGRESSION_SHEET_NAME = "Regression Model"
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

MONTH_NUM = {
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

DAY_BY_PERIOD = {"early": 5, "mid": 15, "late": 25}


def is_blank(value: Any) -> bool:
    return value is None or (isinstance(value, str) and value.strip() == "")


def to_float(value: Any) -> Optional[float]:
    if isinstance(value, (int, float)):
        return float(value)
    if isinstance(value, str):
        text = value.strip().replace(",", "")
        if not text:
            return None
        try:
            return float(text)
        except ValueError:
            return None
    return None


def safe_subtract(a: Any, b: Any) -> Any:
    a_num = to_float(a)
    b_num = to_float(b)
    if a_num is None or b_num is None:
        return ""
    return a_num - b_num


def parse_file_label(file_name: str) -> Dict[str, str]:
    stem = Path(file_name).stem
    match = re.search(
        r"-\s*(?P<ticker>[A-Za-z0-9]+)\s*-\s*"
        r"(?P<period>(?P<phase>Early|Mid|Late)"
        r"(?P<month>Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)"
        r"(?P<year>\d{4}))",
        stem,
        flags=re.IGNORECASE,
    )

    if not match:
        return {
            "ticker": "",
            "model_period": "",
            "model_date": "",
            "model": "",
        }

    ticker = match.group("ticker").upper()
    phase_raw = match.group("phase").lower()
    month_raw = match.group("month").lower()
    year = match.group("year")

    phase = phase_raw.capitalize()
    month = month_raw.capitalize()
    model_period = f"{phase}{month}_{year}"

    month_num = MONTH_NUM.get(month_raw)
    day_num = DAY_BY_PERIOD.get(phase_raw)
    if month_num is None or day_num is None:
        model_date = ""
    else:
        model_date = f"{year}-{month_num:02d}-{day_num:02d}"

    model = f"{ticker}_{model_period}" if model_period else ticker

    return {
        "ticker": ticker,
        "model_period": model_period,
        "model_date": model_date,
        "model": model,
    }


def close_source_workbook(wb: xw.Book) -> None:
    close_attempts = (
        lambda: wb.close(save=False),
        lambda: wb.close(False),
        lambda: wb.api.Close(SaveChanges=False),
        lambda: wb.api.Close(False),
    )
    for attempt in close_attempts:
        try:
            attempt()
            return
        except Exception:
            continue


def set_formula2(cell: xw.main.Range, formula: str) -> None:
    try:
        cell.formula2 = formula
    except Exception:
        # Rare fallback if formula2 isn't exposed on this object.
        cell.formula = formula


def find_anchor(sheet: xw.Sheet, anchor_text: str = "max") -> Optional[Tuple[int, int]]:
    used = sheet.used_range
    values = used.value
    if values is None:
        return None

    if not isinstance(values, list):
        matrix = [[values]]
    elif values and isinstance(values[0], list):
        matrix = values
    else:
        matrix = [values]

    for row_idx, row_values in enumerate(matrix):
        for col_idx, cell_value in enumerate(row_values):
            if isinstance(cell_value, str) and cell_value.strip().lower() == anchor_text:
                return used.row + row_idx, used.column + col_idx
    return None


def extract_empirical_rows(
    wb: xw.Book,
    sheet: xw.Sheet,
    meta: Dict[str, str],
    source_file: str,
) -> List[Dict[str, Any]]:
    anchor = find_anchor(sheet, "max")
    if anchor is None:
        print(f"Skipped {source_file} {EMPIRICAL_SHEET_NAME}: max anchor not found")
        return []

    anchor_row, anchor_col = anchor
    start_row = anchor_row + 1
    helper_col = anchor_col + 8
    sales_captured_col = anchor_col - 5
    rows = []

    formula_rows: List[int] = []
    for quarter_idx in range(N_QUARTERS):
        row = start_row + quarter_idx
        start_for_avg = max(start_row, row - quarter_idx)
        formula = (
            f'=IFERROR(AVERAGE(R{start_for_avg}C{sales_captured_col}:'
            f'R{row}C{sales_captured_col}),"")'
        )
        set_formula2(sheet.cells(row, helper_col), formula)
        formula_rows.append(row)

    if formula_rows:
        wb.app.calculate()

    for quarter_idx, row in enumerate(formula_rows, start=1):
        num_quarters_used = sheet.cells(row, anchor_col - 8).value
        if is_blank(num_quarters_used):
            num_quarters_used = quarter_idx

        reported_sales = sheet.cells(row, anchor_col - 7).value
        quarterly_sales = sheet.cells(row, anchor_col - 6).value
        sales_captured_in_db_pct = sheet.cells(row, anchor_col - 5).value
        growth_rate_pct = sheet.cells(row, anchor_col - 4).value
        last_quarter_used = sheet.cells(row, anchor_col - 3).value
        actual_value = sheet.cells(row, anchor_col - 2).value
        forecast_value = sheet.cells(row, anchor_col - 1).value
        forecast_max = sheet.cells(row, anchor_col).value
        forecast_min = sheet.cells(row, anchor_col + 1).value
        avg_penetration_pct = sheet.cells(row, helper_col).value

        if all(
            is_blank(v)
            for v in (
                forecast_value,
                forecast_max,
                forecast_min,
                actual_value,
                quarterly_sales,
                reported_sales,
            )
        ):
            continue

        row_out = {
            "model": meta["model"],
            "ticker": meta["ticker"],
            "model_period": meta["model_period"],
            "model_date": meta["model_date"],
            "method": "empirical",
            "parameter_name": "avg_penetration_pct",
            "parameter_value": avg_penetration_pct,
            "num_quarters_used": num_quarters_used,
            "last_quarter_used": last_quarter_used,
            "forecast_value": forecast_value,  # estimated total sold
            "actual_value": actual_value,  # reported sales
            "forecast_max": forecast_max,
            "forecast_min": forecast_min,
            "range_width": safe_subtract(forecast_max, forecast_min),
            "avg_penetration_pct": avg_penetration_pct,
            "quarterly_sales": quarterly_sales,
            "reported_sales": reported_sales,
            "growth_rate_pct": growth_rate_pct,
            "sales_captured_in_db_pct": sales_captured_in_db_pct,
            "source_file": source_file,
        }
        rows.append(row_out)

    if formula_rows:
        sheet.range((formula_rows[0], helper_col), (formula_rows[-1], helper_col)).clear_contents()

    return rows


def extract_regression_rows(
    wb: xw.Book,
    sheet: xw.Sheet,
    meta: Dict[str, str],
    source_file: str,
) -> List[Dict[str, Any]]:
    anchor = find_anchor(sheet, "max")
    if anchor is None:
        print(f"Skipped {source_file} {REGRESSION_SHEET_NAME}: max anchor not found")
        return []

    anchor_row, anchor_col = anchor
    y_col = anchor_col - 7
    x_col = anchor_col - 11
    helper_intercept_col = anchor_col + 8
    helper_slope_col = anchor_col + 9

    formula_rows: List[int] = []
    data_end_row = anchor_row - 1

    for quarter_count in range(1, N_QUARTERS + 1):
        data_start_row = data_end_row - quarter_count + 1
        if data_start_row < 1:
            continue

        row = anchor_row + quarter_count
        intercept_formula = (
            f'=IFERROR(INTERCEPT(R{data_start_row}C{y_col}:R{data_end_row}C{y_col},'
            f'R{data_start_row}C{x_col}:R{data_end_row}C{x_col}),"")'
        )
        slope_formula = (
            f'=IFERROR(SLOPE(R{data_start_row}C{y_col}:R{data_end_row}C{y_col},'
            f'R{data_start_row}C{x_col}:R{data_end_row}C{x_col}),"")'
        )
        set_formula2(sheet.cells(row, helper_intercept_col), intercept_formula)
        set_formula2(sheet.cells(row, helper_slope_col), slope_formula)
        formula_rows.append(row)

    if formula_rows:
        wb.app.calculate()

    rows: List[Dict[str, Any]] = []
    for default_quarters, row in enumerate(formula_rows, start=1):
        num_quarters_used = sheet.cells(row, anchor_col - 2).value
        if is_blank(num_quarters_used):
            num_quarters_used = default_quarters

        forecast_value = sheet.cells(row, anchor_col - 1).value  # TOT FCST w/o SA
        actual_value = sheet.cells(row, anchor_col - 3).value
        if is_blank(actual_value):
            actual_value = ""

        forecast_max = sheet.cells(row, anchor_col).value
        forecast_min = sheet.cells(row, anchor_col + 1).value
        intercept = sheet.cells(row, helper_intercept_col).value
        slope = sheet.cells(row, helper_slope_col).value

        if all(is_blank(v) for v in (forecast_value, forecast_max, forecast_min, intercept, slope)):
            continue

        row_out = {
            "model": meta["model"],
            "ticker": meta["ticker"],
            "model_period": meta["model_period"],
            "model_date": meta["model_date"],
            "method": "regression",
            "parameter_name": "num_quarters_used",
            "parameter_value": num_quarters_used,
            "num_quarters_used": num_quarters_used,
            "forecast_value": forecast_value,
            "actual_value": actual_value,
            "forecast_max": forecast_max,
            "forecast_min": forecast_min,
            "range_width": safe_subtract(forecast_max, forecast_min),
            "intercept": intercept,
            "slope": slope,
            "source_file": source_file,
        }
        rows.append(row_out)

    if len(rows) >= 2:
        last = rows[-1]
        prev = rows[-2]
        compare_fields = ("forecast_value", "forecast_max", "forecast_min", "intercept", "slope")
        same_as_previous = True
        for field in compare_fields:
            v1 = to_float(last[field])
            v2 = to_float(prev[field])
            if v1 is not None and v2 is not None:
                if abs(v1 - v2) > 1e-9:
                    same_as_previous = False
                    break
            elif str(last[field]) != str(prev[field]):
                same_as_previous = False
                break
        if same_as_previous:
            rows.pop()

    if formula_rows:
        sheet.range(
            (formula_rows[0], helper_intercept_col),
            (formula_rows[-1], helper_slope_col),
        ).clear_contents()

    return rows


def next_output_path(input_path: Path, output_path: Path) -> Path:
    output_path.mkdir(parents=True, exist_ok=True)
    base_stem = f"{input_path.name}_PARAM"
    candidate = output_path / f"{base_stem}.xlsx"
    if not candidate.exists():
        return candidate

    idx = 1
    while True:
        candidate = output_path / f"{base_stem}.{idx}.xlsx"
        if not candidate.exists():
            return candidate
        idx += 1


def write_table(
    app: xw.App,
    sheet: xw.Sheet,
    headers: List[str],
    rows: List[Dict[str, Any]],
) -> None:
    sheet.clear_contents()
    sheet.range((1, 1)).value = headers

    if rows:
        values = [[row.get(col, "") for col in headers] for row in rows]
        sheet.range((2, 1)).value = values

    last_row = max(1, len(rows) + 1)
    last_col = len(headers)

    header_range = sheet.range((1, 1), (1, last_col))
    header_range.api.Font.Bold = True
    sheet.range((1, 1), (last_row, last_col)).api.AutoFilter()

    # Freeze header row.
    sheet.activate()
    app.api.ActiveWindow.SplitRow = 1
    app.api.ActiveWindow.SplitColumn = 0
    app.api.ActiveWindow.FreezePanes = True

    # Reasonable column widths based on headers and a sample of values.
    sample_rows = rows[:200]
    for col_idx, header in enumerate(headers, start=1):
        max_len = len(header)
        for row in sample_rows:
            value = row.get(header, "")
            if value is None:
                continue
            max_len = max(max_len, len(str(value)))
        width = min(50, max(12, max_len + 2))
        sheet.range((1, col_idx)).column_width = width


def collect_source_files(input_path: Path) -> List[Path]:
    files: List[Path] = []
    param_prefix = f"{input_path.name.lower()}_param"
    for file_path in sorted(input_path.iterdir(), key=lambda p: p.name.lower()):
        if not file_path.is_file():
            print(f"Skipped {file_path.name}: not a file")
            continue
        if file_path.name.startswith("~"):
            print(f"Skipped {file_path.name}: temp file")
            continue
        if file_path.suffix.lower() != ".xlsx":
            print(f"Skipped {file_path.name}: not an .xlsx file")
            continue
        # Prevent re-processing prior output files when output_dir == input_dir.
        if file_path.stem.lower().startswith(param_prefix):
            print(f"Skipped {file_path.name}: prior PARAM output file")
            continue
        files.append(file_path)
    return files


def get_sheet_if_exists(wb: xw.Book, sheet_name: str) -> Optional[xw.Sheet]:
    try:
        return wb.sheets[sheet_name]
    except Exception:
        return None


def process_workbook_once(
    wb: xw.Book,
    source_file: str,
    meta: Dict[str, str],
) -> Tuple[List[Dict[str, Any]], List[Dict[str, Any]]]:
    empirical_rows: List[Dict[str, Any]] = []
    regression_rows: List[Dict[str, Any]] = []

    empirical_sheet = get_sheet_if_exists(wb, EMPIRICAL_SHEET_NAME)
    if empirical_sheet is None:
        print(f"Skipped {source_file} {EMPIRICAL_SHEET_NAME}: sheet not found")
    else:
        empirical_rows.extend(extract_empirical_rows(wb, empirical_sheet, meta, source_file))

    regression_sheet = get_sheet_if_exists(wb, REGRESSION_SHEET_NAME)
    if regression_sheet is None:
        print(f"Skipped {source_file} {REGRESSION_SHEET_NAME}: sheet not found")
    else:
        regression_rows.extend(extract_regression_rows(wb, regression_sheet, meta, source_file))

    return empirical_rows, regression_rows


def main() -> None:
    input_path = Path(input_dir).expanduser().resolve()
    output_path = Path(output_dir).expanduser().resolve()

    if not input_path.exists() or not input_path.is_dir():
        raise FileNotFoundError(f"input_dir does not exist or is not a directory: {input_path}")

    source_files = collect_source_files(input_path)
    final_output_path = next_output_path(input_path, output_path)

    all_empirical_rows: List[Dict[str, Any]] = []
    all_regression_rows: List[Dict[str, Any]] = []
    processed_file_count = 0

    app = xw.App(visible=False, add_book=False)
    try:
        app.display_alerts = False
        app.screen_updating = False
        try:
            app.calculation = "manual"
        except Exception:
            pass

        for file_path in source_files:
            wb = None
            try:
                # Source workbook safety requirement.
                wb = app.books.open(str(file_path), update_links=False)
            except Exception as exc:
                print(f"Skipped {file_path.name}: failed to open ({exc})")
                continue

            try:
                metadata = parse_file_label(file_path.name)
                emp_rows, reg_rows = process_workbook_once(wb, file_path.name, metadata)
                all_empirical_rows.extend(emp_rows)
                all_regression_rows.extend(reg_rows)
                processed_file_count += 1
                print(f"Processed {file_path.name}")
            except Exception as exc:
                print(f"Skipped {file_path.name}: processing error ({exc})")
            finally:
                if wb is not None:
                    close_source_workbook(wb)

        output_wb = app.books.add()
        try:
            empirical_sheet = output_wb.sheets[0]
            empirical_sheet.name = "empirical_candidates"
            regression_sheet = output_wb.sheets.add("regression_candidates", after=empirical_sheet)

            write_table(app, empirical_sheet, EMPIRICAL_COLUMNS, all_empirical_rows)
            write_table(app, regression_sheet, REGRESSION_COLUMNS, all_regression_rows)

            output_wb.save(str(final_output_path))
        finally:
            output_wb.close()
    finally:
        app.quit()

    print(f"Output path: {final_output_path}")
    print(f"Number of files processed: {processed_file_count}")
    print(f"Number of empirical rows: {len(all_empirical_rows)}")
    print(f"Number of regression rows: {len(all_regression_rows)}")


if __name__ == "__main__":
    main()
