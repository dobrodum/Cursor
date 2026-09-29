#!/usr/bin/env python3
from __future__ import annotations

import re
from dataclasses import dataclass
from datetime import date
from pathlib import Path
from typing import Any, Dict, List, Optional, Sequence, Tuple

import xlwings as xw
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

# Inputs (edit these two paths as needed)
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


@dataclass
class ModelLabels:
    model: str
    ticker: str
    model_period: str
    model_date: str


def to_2d(values: Any) -> List[List[Any]]:
    if values is None:
        return []
    if isinstance(values, tuple):
        values = [list(row) if isinstance(row, tuple) else row for row in values]
    if not isinstance(values, list):
        return [[values]]
    if not values:
        return []
    if isinstance(values[0], list):
        return values
    return [values]


def as_float(value: Any) -> Optional[float]:
    if value is None or value == "":
        return None
    if isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        return float(value)
    try:
        return float(str(value).replace(",", "").strip())
    except Exception:
        return None


def as_int(value: Any) -> Optional[int]:
    numeric = as_float(value)
    if numeric is None:
        return None
    return int(round(numeric))


def parse_model_labels(file_name: str) -> Optional[ModelLabels]:
    stem = Path(file_name).stem
    pattern = re.compile(
        r"-\s*([A-Za-z0-9]+)\s*-\s*((Early|Mid|Late)([A-Za-z]{3})(\d{4}))",
        re.IGNORECASE,
    )
    match = pattern.search(stem)
    if not match:
        return None

    ticker = match.group(1).upper()
    phase = match.group(3).title()
    month_abbrev = match.group(4).title()
    year = int(match.group(5))

    month_lookup = {
        "Jan": 1,
        "Feb": 2,
        "Mar": 3,
        "Apr": 4,
        "May": 5,
        "Jun": 6,
        "Jul": 7,
        "Aug": 8,
        "Sep": 9,
        "Oct": 10,
        "Nov": 11,
        "Dec": 12,
    }
    day_lookup = {"Early": 5, "Mid": 15, "Late": 25}

    month = month_lookup.get(month_abbrev)
    day = day_lookup.get(phase)
    if month is None or day is None:
        return None

    model_period = f"{phase}{month_abbrev}_{year}"
    model_date = date(year, month, day).isoformat()
    model = f"{ticker}_{model_period}"
    return ModelLabels(
        model=model,
        ticker=ticker,
        model_period=model_period,
        model_date=model_date,
    )


def unique_output_path(base_output_dir: Path, input_folder_name: str) -> Path:
    base_output_dir.mkdir(parents=True, exist_ok=True)
    first_path = base_output_dir / f"{input_folder_name}_PARAM.xlsx"
    if not first_path.exists():
        return first_path

    counter = 1
    while True:
        candidate = base_output_dir / f"{input_folder_name}_PARAM.{counter}.xlsx"
        if not candidate.exists():
            return candidate
        counter += 1


def close_workbook_safe(wb: xw.main.Book) -> None:
    try:
        wb.close(save=False)
        return
    except TypeError:
        pass
    except Exception:
        pass

    try:
        wb.api.Close(False)
        return
    except Exception:
        pass

    try:
        wb.close()
    except Exception:
        pass


def find_anchor_max(sheet: xw.main.Sheet) -> Tuple[int, int]:
    used = sheet.used_range
    values = to_2d(used.value)
    if not values:
        raise ValueError("sheet is empty")

    start_row = used.row
    start_col = used.column
    for r_offset, row_values in enumerate(values):
        for c_offset, value in enumerate(row_values):
            if isinstance(value, str) and value.strip().lower() == "max":
                return start_row + r_offset, start_col + c_offset
    raise ValueError('could not find "max" anchor')


def normalize_header(value: Any) -> str:
    if value is None:
        return ""
    text = re.sub(r"[^a-z0-9]+", " ", str(value).strip().lower())
    return re.sub(r"\s+", " ", text).strip()


def find_columns_from_headers(
    header_values: Sequence[Any],
    start_col: int,
    patterns: Dict[str, List[Tuple[str, ...]]],
) -> Dict[str, int]:
    headers = [normalize_header(v) for v in header_values]
    col_map: Dict[str, int] = {}

    for field, alternatives in patterns.items():
        for index, header in enumerate(headers):
            if not header:
                continue
            for parts in alternatives:
                if all(part in header for part in parts):
                    col_map[field] = start_col + index
                    break
            if field in col_map:
                break
    return col_map


def read_cell(sheet: xw.main.Sheet, row: int, col: int) -> Any:
    if row < 1 or col < 1:
        return None
    return sheet.cells(row, col).value


def apply_empirical_avg_formula(
    sheet: xw.main.Sheet,
    row: int,
    avg_col: int,
    num_quarters: int,
) -> bool:
    if avg_col < 2 or num_quarters < 1:
        return False
    try:
        # R1C1 formula avoids letter conversion overhead.
        sheet.cells(row, avg_col).formula2 = f"=AVERAGE(RC[-{num_quarters}]:RC[-1])"
        return True
    except Exception:
        return False


def process_empirical_sheet(
    wb: xw.main.Book,
    sheet: xw.main.Sheet,
    labels: ModelLabels,
    source_file: str,
) -> List[Dict[str, Any]]:
    anchor_row, anchor_col = find_anchor_max(sheet)

    header_start_col = max(1, anchor_col - 30)
    header_end_col = anchor_col + 12
    header_values = to_2d(
        sheet.range((anchor_row, header_start_col), (anchor_row, header_end_col)).value
    )
    header_row = header_values[0] if header_values else []
    patterns = {
        "num_quarters_used": [("num", "quarter"), ("quarters", "used")],
        "last_quarter_used": [("last", "quarter")],
        "forecast_value": [("estimated", "total", "sold"), ("forecast", "value")],
        "actual_value": [("reported", "sales"), ("actual", "value")],
        "forecast_max": [("max",)],
        "forecast_min": [("min",)],
        "avg_penetration_pct": [("avg", "penetration"), ("penetration", "avg")],
        "quarterly_sales": [("quarterly", "sales"), ("qtr", "sales")],
        "reported_sales": [("reported", "sales")],
        "growth_rate_pct": [("growth", "rate")],
        "sales_captured_in_db_pct": [("captured", "db"), ("captured", "database")],
    }
    col_map = find_columns_from_headers(header_row, header_start_col, patterns)

    defaults = {
        "num_quarters_used": anchor_col - 10,
        "last_quarter_used": anchor_col - 9,
        "avg_penetration_pct": anchor_col - 6,
        "quarterly_sales": anchor_col - 5,
        "forecast_value": anchor_col - 4,
        "actual_value": anchor_col - 3,
        "reported_sales": anchor_col - 3,
        "growth_rate_pct": anchor_col - 2,
        "sales_captured_in_db_pct": anchor_col - 1,
        "forecast_max": anchor_col,
        "forecast_min": anchor_col + 1,
    }
    for key, default_col in defaults.items():
        if key not in col_map:
            col_map[key] = default_col

    updated_formulas = False
    for n in range(1, N_QUARTERS + 1):
        row = anchor_row + n
        updated_formulas = (
            apply_empirical_avg_formula(sheet, row, col_map["avg_penetration_pct"], n)
            or updated_formulas
        )
    if updated_formulas:
        wb.app.calculate()

    rows: List[Dict[str, Any]] = []
    for n in range(1, N_QUARTERS + 1):
        row = anchor_row + n
        num_quarters_used = as_int(read_cell(sheet, row, col_map["num_quarters_used"])) or n
        last_quarter_used = read_cell(sheet, row, col_map["last_quarter_used"])
        forecast_value = as_float(read_cell(sheet, row, col_map["forecast_value"]))
        actual_value = as_float(read_cell(sheet, row, col_map["actual_value"]))
        forecast_max = as_float(read_cell(sheet, row, col_map["forecast_max"]))
        forecast_min = as_float(read_cell(sheet, row, col_map["forecast_min"]))
        avg_penetration = as_float(read_cell(sheet, row, col_map["avg_penetration_pct"]))
        quarterly_sales = as_float(read_cell(sheet, row, col_map["quarterly_sales"]))
        reported_sales = as_float(read_cell(sheet, row, col_map["reported_sales"]))
        growth_rate_pct = as_float(read_cell(sheet, row, col_map["growth_rate_pct"]))
        sales_captured = as_float(read_cell(sheet, row, col_map["sales_captured_in_db_pct"]))

        if (
            forecast_max is None
            and forecast_min is None
            and forecast_value is None
            and avg_penetration is None
        ):
            continue

        range_width = None
        if forecast_max is not None and forecast_min is not None:
            range_width = forecast_max - forecast_min

        rows.append(
            {
                "model": labels.model,
                "ticker": labels.ticker,
                "model_period": labels.model_period,
                "model_date": labels.model_date,
                "method": "empirical",
                "parameter_name": "avg_penetration_pct",
                "parameter_value": avg_penetration,
                "num_quarters_used": num_quarters_used,
                "last_quarter_used": last_quarter_used,
                "forecast_value": forecast_value,
                "actual_value": actual_value,
                "forecast_max": forecast_max,
                "forecast_min": forecast_min,
                "range_width": range_width,
                "avg_penetration_pct": avg_penetration,
                "quarterly_sales": quarterly_sales,
                "reported_sales": reported_sales,
                "growth_rate_pct": growth_rate_pct,
                "sales_captured_in_db_pct": sales_captured,
                "source_file": source_file,
            }
        )
    return rows


def process_regression_sheet(
    wb: xw.main.Book,
    sheet: xw.main.Sheet,
    labels: ModelLabels,
    source_file: str,
) -> List[Dict[str, Any]]:
    anchor_row, anchor_col = find_anchor_max(sheet)
    y_col = anchor_col - 7
    x_col = anchor_col - 11

    header_start_col = max(1, anchor_col - 30)
    header_end_col = anchor_col + 12
    header_values = to_2d(
        sheet.range((anchor_row, header_start_col), (anchor_row, header_end_col)).value
    )
    header_row = header_values[0] if header_values else []
    patterns = {
        "num_quarters_used": [("num", "quarter"), ("quarters", "used")],
        "forecast_value": [("tot", "fcst", "w", "sa"), ("forecast", "without", "sa")],
        "actual_value": [("actual", "value"), ("reported", "sales")],
        "forecast_max": [("max",)],
        "forecast_min": [("min",)],
    }
    col_map = find_columns_from_headers(header_row, header_start_col, patterns)
    defaults = {
        "num_quarters_used": anchor_col - 10,
        "forecast_value": anchor_col - 1,
        "actual_value": anchor_col - 2,
        "forecast_max": anchor_col,
        "forecast_min": anchor_col + 1,
    }
    for key, default_col in defaults.items():
        if key not in col_map:
            col_map[key] = default_col

    used = sheet.used_range
    start_row = used.row
    end_row = used.row + used.rows.count - 1

    numeric_rows: List[int] = []
    for row in range(start_row, anchor_row):
        x_val = as_float(read_cell(sheet, row, x_col))
        y_val = as_float(read_cell(sheet, row, y_col))
        if x_val is not None and y_val is not None:
            numeric_rows.append(row)

    formula_cells: List[Tuple[xw.main.Range, xw.main.Range, int]] = []
    for n in range(1, N_QUARTERS + 1):
        if len(numeric_rows) < max(2, n):
            continue
        calc_row = anchor_row + n
        first_row = numeric_rows[-n]
        last_row = numeric_rows[-1]
        intercept_cell = sheet.cells(calc_row, anchor_col + 4)
        slope_cell = sheet.cells(calc_row, anchor_col + 5)
        intercept_cell.formula2 = (
            f"=INTERCEPT(R{first_row}C{y_col}:R{last_row}C{y_col},"
            f"R{first_row}C{x_col}:R{last_row}C{x_col})"
        )
        slope_cell.formula2 = (
            f"=SLOPE(R{first_row}C{y_col}:R{last_row}C{y_col},"
            f"R{first_row}C{x_col}:R{last_row}C{x_col})"
        )
        formula_cells.append((intercept_cell, slope_cell, n))

    if formula_cells:
        wb.app.calculate()

    rows: List[Dict[str, Any]] = []
    previous_key: Optional[Tuple[Optional[float], Optional[float], Optional[float], Optional[float], Optional[float]]] = None
    latest_x = as_float(read_cell(sheet, numeric_rows[-1], x_col)) if numeric_rows else None
    for intercept_cell, slope_cell, n in formula_cells:
        calc_row = intercept_cell.row
        num_quarters_used = as_int(read_cell(sheet, calc_row, col_map["num_quarters_used"])) or n
        intercept = as_float(intercept_cell.value)
        slope = as_float(slope_cell.value)
        forecast_value = as_float(read_cell(sheet, calc_row, col_map["forecast_value"]))
        if forecast_value is None and latest_x is not None and intercept is not None and slope is not None:
            forecast_value = intercept + (slope * latest_x)
        actual_value = as_float(read_cell(sheet, calc_row, col_map["actual_value"]))
        forecast_max = as_float(read_cell(sheet, calc_row, col_map["forecast_max"]))
        forecast_min = as_float(read_cell(sheet, calc_row, col_map["forecast_min"]))
        if (
            forecast_value is None
            and forecast_max is None
            and forecast_min is None
            and intercept is None
            and slope is None
        ):
            continue

        current_key = (intercept, slope, forecast_value, forecast_max, forecast_min)
        if previous_key is not None and current_key == previous_key:
            continue
        previous_key = current_key

        range_width = None
        if forecast_max is not None and forecast_min is not None:
            range_width = forecast_max - forecast_min

        rows.append(
            {
                "model": labels.model,
                "ticker": labels.ticker,
                "model_period": labels.model_period,
                "model_date": labels.model_date,
                "method": "regression",
                "parameter_name": "num_quarters_used",
                "parameter_value": num_quarters_used,
                "num_quarters_used": num_quarters_used,
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


def write_rows_to_sheet(ws: Any, columns: List[str], rows: List[Dict[str, Any]]) -> None:
    ws.append(columns)
    for row in rows:
        ws.append([row.get(col) for col in columns])

    for header_cell in ws[1]:
        header_cell.font = Font(bold=True)

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions

    for col_index, column in enumerate(columns, start=1):
        values = [column]
        values.extend(str(row.get(column, "")) for row in rows if row.get(column) is not None)
        max_len = max(len(v) for v in values) if values else len(column)
        width = min(max(max_len + 2, 12), 50)
        ws.column_dimensions[get_column_letter(col_index)].width = width


def write_output_workbook(
    output_path: Path,
    empirical_rows: List[Dict[str, Any]],
    regression_rows: List[Dict[str, Any]],
) -> None:
    out_wb = Workbook()
    ws_empirical = out_wb.active
    ws_empirical.title = "empirical_candidates"
    ws_regression = out_wb.create_sheet("regression_candidates")

    write_rows_to_sheet(ws_empirical, EMPIRICAL_COLUMNS, empirical_rows)
    write_rows_to_sheet(ws_regression, REGRESSION_COLUMNS, regression_rows)
    out_wb.save(output_path)


def run() -> None:
    source_dir = Path(input_dir)
    destination_dir = Path(output_dir)
    input_folder_name = source_dir.name or "input"
    output_path = unique_output_path(destination_dir, input_folder_name)

    empirical_rows: List[Dict[str, Any]] = []
    regression_rows: List[Dict[str, Any]] = []
    files_processed = 0

    if not source_dir.exists():
        raise FileNotFoundError(f"Input directory not found: {source_dir}")

    source_files = sorted(source_dir.iterdir(), key=lambda p: p.name.lower())

    app = xw.App(visible=False, add_book=False)
    app.display_alerts = False
    app.screen_updating = False
    app.calculation = "manual"
    try:
        for file_path in source_files:
            if not file_path.is_file():
                continue
            if file_path.name.startswith("~"):
                print(f"SKIP {file_path.name}: temp file")
                continue
            if file_path.suffix.lower() != ".xlsx":
                print(f"SKIP {file_path.name}: not .xlsx")
                continue

            labels = parse_model_labels(file_path.name)
            if labels is None:
                print(f"SKIP {file_path.name}: filename format not recognized")
                continue

            print(f"PROCESS {file_path.name}")
            wb = None
            try:
                wb = app.books.open(str(file_path), update_links=False)
                sheet_names = {s.name for s in wb.sheets}
                if "Empirical Model" not in sheet_names:
                    print(f"SKIP {file_path.name}: missing 'Empirical Model' sheet")
                    continue
                if "Regression Model" not in sheet_names:
                    print(f"SKIP {file_path.name}: missing 'Regression Model' sheet")
                    continue

                empirical_rows.extend(
                    process_empirical_sheet(
                        wb=wb,
                        sheet=wb.sheets["Empirical Model"],
                        labels=labels,
                        source_file=file_path.name,
                    )
                )
                regression_rows.extend(
                    process_regression_sheet(
                        wb=wb,
                        sheet=wb.sheets["Regression Model"],
                        labels=labels,
                        source_file=file_path.name,
                    )
                )
                files_processed += 1
            except Exception as exc:
                print(f"SKIP {file_path.name}: processing error ({exc})")
            finally:
                if wb is not None:
                    close_workbook_safe(wb)
    finally:
        app.quit()

    write_output_workbook(output_path, empirical_rows, regression_rows)

    print(f"OUTPUT {output_path}")
    print(f"FILES_PROCESSED {files_processed}")
    print(f"EMPIRICAL_ROWS {len(empirical_rows)}")
    print(f"REGRESSION_ROWS {len(regression_rows)}")


if __name__ == "__main__":
    run()
