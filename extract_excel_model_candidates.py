#!/usr/bin/env python3
from __future__ import annotations

import re
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

import xlwings as xw
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

# -----------------------------
# User-configurable directories
# -----------------------------
input_dir = "/path/to/input"
output_dir = "/path/to/output"


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

# Offsets are relative to the "max" anchor column in each source model sheet.
EMPIRICAL_OFFSETS = {
    "num_quarters_used": -9,
    "last_quarter_used": -8,
    "avg_pen_source": -7,
    "quarterly_sales": -6,
    "reported_sales": -5,
    "growth_rate_pct": -4,
    "sales_captured_in_db_pct": -3,
    "forecast_value": -2,  # estimated total sold
    "forecast_max": 0,
    "forecast_min": 1,
    "temp_avg_calc": 5,  # temporary formula write location
}

REGRESSION_OFFSETS = {
    "num_quarters_used": -8,
    "actual_value": -2,
    "forecast_total_without_sa": -1,  # TOT FCST w/o SA
    "forecast_max": 0,
    "forecast_min": 1,
    "temp_intercept": 5,  # temporary formula write location
    "temp_slope": 6,  # temporary formula write location
}


@dataclass
class FileLabels:
    model: str
    ticker: str
    model_period: str
    model_date: str


def to_float(value: Any) -> Optional[float]:
    if value is None or value == "":
        return None
    if isinstance(value, (int, float)):
        return float(value)
    if isinstance(value, str):
        cleaned = value.strip().replace(",", "")
        if cleaned.endswith("%"):
            cleaned = cleaned[:-1]
        if cleaned == "":
            return None
        try:
            return float(cleaned)
        except ValueError:
            return None
    return None


def to_int(value: Any) -> Optional[int]:
    numeric = to_float(value)
    if numeric is None:
        return None
    return int(round(numeric))


def parse_file_labels(file_name: str) -> Optional[FileLabels]:
    stem = Path(file_name).stem
    parts = [part.strip() for part in stem.split(" - ")]
    if len(parts) < 3:
        return None

    ticker = parts[1].upper()
    period_match = re.search(
        r"\b(Early|Mid|Late)\s*([A-Za-z]{3,9})\s*([12]\d{3})\b",
        stem,
        flags=re.IGNORECASE,
    )
    if not period_match:
        return None

    timing_raw, month_raw, year_raw = period_match.groups()
    timing = timing_raw.capitalize()
    year = int(year_raw)

    month_abbr = month_raw[:3].title()
    try:
        month_number = datetime.strptime(month_abbr, "%b").month
    except ValueError:
        return None

    day_map = {"Early": 5, "Mid": 15, "Late": 25}
    day = day_map[timing]
    model_period = f"{timing}{month_abbr}_{year}"
    model_date = f"{year:04d}-{month_number:02d}-{day:02d}"
    model = f"{ticker}_{model_period}"
    return FileLabels(model=model, ticker=ticker, model_period=model_period, model_date=model_date)


def find_anchor_cell(sheet: xw.Sheet, anchor_text: str = "max") -> Optional[Tuple[int, int]]:
    used_range = sheet.used_range
    values = used_range.options(ndim=2).value
    start_row = used_range.row
    start_col = used_range.column

    for row_idx, row_values in enumerate(values):
        for col_idx, cell_value in enumerate(row_values):
            if isinstance(cell_value, str) and cell_value.strip().lower() == anchor_text:
                return start_row + row_idx, start_col + col_idx
    return None


def get_cell_value(sheet: xw.Sheet, row: int, col: int) -> Any:
    if row < 1 or col < 1:
        return None
    return sheet.cells(row, col).value


def set_formula2(cell: xw.Range, formula_r1c1: str) -> None:
    try:
        cell.formula2 = formula_r1c1
    except Exception:
        cell.formula = formula_r1c1


def clear_range(sheet: xw.Sheet, start_row: int, end_row: int, col: int) -> None:
    if start_row > end_row or col < 1:
        return
    sheet.range((start_row, col), (end_row, col)).clear_contents()


def extract_empirical_rows(wb: xw.Book, labels: FileLabels, source_file: str) -> List[Dict[str, Any]]:
    try:
        sheet = wb.sheets["Empirical Model"]
    except Exception:
        print(f"  Skipped empirical extraction for {source_file}: missing 'Empirical Model' sheet")
        return []

    anchor = find_anchor_cell(sheet, "max")
    if anchor is None:
        print(f"  Skipped empirical extraction for {source_file}: 'max' anchor not found")
        return []

    anchor_row, anchor_col = anchor
    data_start_row = anchor_row + 1
    temp_avg_col = anchor_col + EMPIRICAL_OFFSETS["temp_avg_calc"]
    avg_pen_col = anchor_col + EMPIRICAL_OFFSETS["avg_pen_source"]

    # Write all temporary average formulas first, then calculate once.
    for idx in range(N_QUARTERS):
        row = data_start_row + idx
        num_quarters = idx + 1
        window_start = max(data_start_row, row - num_quarters + 1)
        formula = f"=AVERAGE(R{window_start}C{avg_pen_col}:R{row}C{avg_pen_col})"
        set_formula2(sheet.cells(row, temp_avg_col), formula)

    wb.app.calculate()

    avg_values = sheet.range(
        (data_start_row, temp_avg_col),
        (data_start_row + N_QUARTERS - 1, temp_avg_col),
    ).options(ndim=2).value
    avg_values = [row_vals[0] for row_vals in avg_values]

    rows: List[Dict[str, Any]] = []
    for idx in range(N_QUARTERS):
        row = data_start_row + idx
        num_quarters_used = (
            to_int(get_cell_value(sheet, row, anchor_col + EMPIRICAL_OFFSETS["num_quarters_used"]))
            or idx + 1
        )
        last_quarter_used = get_cell_value(sheet, row, anchor_col + EMPIRICAL_OFFSETS["last_quarter_used"])
        forecast_value = to_float(get_cell_value(sheet, row, anchor_col + EMPIRICAL_OFFSETS["forecast_value"]))
        forecast_max = to_float(get_cell_value(sheet, row, anchor_col + EMPIRICAL_OFFSETS["forecast_max"]))
        forecast_min = to_float(get_cell_value(sheet, row, anchor_col + EMPIRICAL_OFFSETS["forecast_min"]))
        quarterly_sales = to_float(get_cell_value(sheet, row, anchor_col + EMPIRICAL_OFFSETS["quarterly_sales"]))
        reported_sales = to_float(get_cell_value(sheet, row, anchor_col + EMPIRICAL_OFFSETS["reported_sales"]))
        growth_rate_pct = to_float(get_cell_value(sheet, row, anchor_col + EMPIRICAL_OFFSETS["growth_rate_pct"]))
        sales_captured_in_db_pct = to_float(
            get_cell_value(sheet, row, anchor_col + EMPIRICAL_OFFSETS["sales_captured_in_db_pct"])
        )
        avg_penetration_pct = to_float(avg_values[idx])
        actual_value = reported_sales

        if (
            forecast_value is None
            and forecast_max is None
            and forecast_min is None
            and quarterly_sales is None
            and reported_sales is None
        ):
            continue

        range_width = (
            (forecast_max - forecast_min)
            if forecast_max is not None and forecast_min is not None
            else None
        )

        rows.append(
            {
                "model": labels.model,
                "ticker": labels.ticker,
                "model_period": labels.model_period,
                "model_date": labels.model_date,
                "method": "empirical",
                "parameter_name": "avg_penetration_pct",
                "parameter_value": avg_penetration_pct,
                "num_quarters_used": num_quarters_used,
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
                "sales_captured_in_db_pct": sales_captured_in_db_pct,
                "source_file": source_file,
            }
        )

    clear_range(sheet, data_start_row, data_start_row + N_QUARTERS - 1, temp_avg_col)
    return rows


def extract_regression_rows(wb: xw.Book, labels: FileLabels, source_file: str) -> List[Dict[str, Any]]:
    try:
        sheet = wb.sheets["Regression Model"]
    except Exception:
        print(f"  Skipped regression extraction for {source_file}: missing 'Regression Model' sheet")
        return []

    anchor = find_anchor_cell(sheet, "max")
    if anchor is None:
        print(f"  Skipped regression extraction for {source_file}: 'max' anchor not found")
        return []

    anchor_row, anchor_col = anchor
    data_start_row = anchor_row + 1
    y_col = anchor_col - 7
    x_col = anchor_col - 11
    temp_intercept_col = anchor_col + REGRESSION_OFFSETS["temp_intercept"]
    temp_slope_col = anchor_col + REGRESSION_OFFSETS["temp_slope"]

    # Write all temporary intercept/slope formulas first, then calculate once.
    row_meta: List[Tuple[int, int]] = []
    for idx in range(N_QUARTERS):
        row = data_start_row + idx
        num_quarters_used = (
            to_int(get_cell_value(sheet, row, anchor_col + REGRESSION_OFFSETS["num_quarters_used"]))
            or idx + 1
        )
        if num_quarters_used < 2:
            num_quarters_used = 2
        y_end = row
        y_start = max(data_start_row, y_end - num_quarters_used + 1)
        x_end = y_end
        x_start = y_start

        intercept_formula = (
            f"=INTERCEPT(R{y_start}C{y_col}:R{y_end}C{y_col},R{x_start}C{x_col}:R{x_end}C{x_col})"
        )
        slope_formula = f"=SLOPE(R{y_start}C{y_col}:R{y_end}C{y_col},R{x_start}C{x_col}:R{x_end}C{x_col})"

        set_formula2(sheet.cells(row, temp_intercept_col), intercept_formula)
        set_formula2(sheet.cells(row, temp_slope_col), slope_formula)
        row_meta.append((row, num_quarters_used))

    wb.app.calculate()

    intercept_vals = sheet.range(
        (data_start_row, temp_intercept_col),
        (data_start_row + N_QUARTERS - 1, temp_intercept_col),
    ).options(ndim=2).value
    slope_vals = sheet.range(
        (data_start_row, temp_slope_col),
        (data_start_row + N_QUARTERS - 1, temp_slope_col),
    ).options(ndim=2).value
    intercept_vals = [row_vals[0] for row_vals in intercept_vals]
    slope_vals = [row_vals[0] for row_vals in slope_vals]

    rows: List[Dict[str, Any]] = []
    previous_signature: Optional[Tuple[Any, ...]] = None
    for idx, (row, num_quarters_used) in enumerate(row_meta):
        forecast_value = to_float(
            get_cell_value(sheet, row, anchor_col + REGRESSION_OFFSETS["forecast_total_without_sa"])
        )
        actual_value = to_float(get_cell_value(sheet, row, anchor_col + REGRESSION_OFFSETS["actual_value"]))
        forecast_max = to_float(get_cell_value(sheet, row, anchor_col + REGRESSION_OFFSETS["forecast_max"]))
        forecast_min = to_float(get_cell_value(sheet, row, anchor_col + REGRESSION_OFFSETS["forecast_min"]))
        intercept = to_float(intercept_vals[idx])
        slope = to_float(slope_vals[idx])

        if (
            forecast_value is None
            and forecast_max is None
            and forecast_min is None
            and intercept is None
            and slope is None
        ):
            continue

        signature = (
            num_quarters_used,
            forecast_value,
            forecast_max,
            forecast_min,
            intercept,
            slope,
        )
        if previous_signature is not None and signature == previous_signature:
            continue
        previous_signature = signature

        range_width = (
            (forecast_max - forecast_min)
            if forecast_max is not None and forecast_min is not None
            else None
        )
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

    clear_range(sheet, data_start_row, data_start_row + N_QUARTERS - 1, temp_intercept_col)
    clear_range(sheet, data_start_row, data_start_row + N_QUARTERS - 1, temp_slope_col)
    return rows


def build_output_path(input_path: Path, output_path: Path) -> Path:
    base_name = f"{input_path.name}_PARAM"
    candidate = output_path / f"{base_name}.xlsx"
    counter = 1
    while candidate.exists():
        candidate = output_path / f"{base_name}.{counter}.xlsx"
        counter += 1
    return candidate


def close_source_workbook(wb: xw.Book) -> None:
    try:
        wb.close(save=False)
        return
    except TypeError:
        pass
    except Exception:
        # Continue to fallback approaches.
        pass

    try:
        wb.api.Close(SaveChanges=False)
        return
    except Exception:
        pass

    try:
        wb.api.Saved = True
        wb.close()
    except Exception as exc:
        print(f"  Warning: unable to close source workbook safely: {exc}")


def write_output_sheet(ws, columns: List[str], rows: List[Dict[str, Any]]) -> None:
    ws.append(columns)
    for row_data in rows:
        ws.append([row_data.get(col) for col in columns])

    header_font = Font(bold=True)
    for cell in ws[1]:
        cell.font = header_font

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(columns))}{max(1, ws.max_row)}"

    for col_idx, column_name in enumerate(columns, start=1):
        max_length = len(column_name)
        for row_idx in range(2, ws.max_row + 1):
            value = ws.cell(row=row_idx, column=col_idx).value
            text = "" if value is None else str(value)
            max_length = max(max_length, len(text))
        ws.column_dimensions[get_column_letter(col_idx)].width = min(max(max_length + 2, 12), 40)


def create_output_workbook(
    output_file_path: Path,
    empirical_rows: List[Dict[str, Any]],
    regression_rows: List[Dict[str, Any]],
) -> None:
    wb = Workbook()
    ws_empirical = wb.active
    ws_empirical.title = "empirical_candidates"
    ws_regression = wb.create_sheet("regression_candidates")

    write_output_sheet(ws_empirical, EMPIRICAL_COLUMNS, empirical_rows)
    write_output_sheet(ws_regression, REGRESSION_COLUMNS, regression_rows)

    wb.save(output_file_path)


def main() -> None:
    input_path = Path(input_dir).expanduser().resolve()
    output_path = Path(output_dir).expanduser().resolve()

    if not input_path.exists() or not input_path.is_dir():
        raise ValueError(f"input_dir does not exist or is not a directory: {input_path}")
    output_path.mkdir(parents=True, exist_ok=True)

    output_file_path = build_output_path(input_path, output_path)

    empirical_rows: List[Dict[str, Any]] = []
    regression_rows: List[Dict[str, Any]] = []
    files_processed = 0

    source_files: List[Path] = []
    for file_path in sorted(input_path.iterdir()):
        if not file_path.is_file():
            continue
        if file_path.name.startswith("~"):
            print(f"Skipped {file_path.name}: temp file")
            continue
        if file_path.suffix.lower() != ".xlsx":
            print(f"Skipped {file_path.name}: not an .xlsx file")
            continue
        source_files.append(file_path)

    app = xw.App(visible=False, add_book=False)
    app.display_alerts = False
    app.screen_updating = False
    try:
        try:
            app.calculation = "manual"
        except Exception:
            pass

        for file_path in source_files:
            print(f"Processing {file_path.name}")
            labels = parse_file_labels(file_path.name)
            if labels is None:
                print(
                    f"Skipped {file_path.name}: filename does not match expected pattern "
                    "(... - TICKER - Early/Mid/LateMonYYYY_...)"
                )
                continue

            wb: Optional[xw.Book] = None
            try:
                wb = app.books.open(str(file_path), update_links=False)
                empirical_rows.extend(extract_empirical_rows(wb, labels, file_path.name))
                regression_rows.extend(extract_regression_rows(wb, labels, file_path.name))
                files_processed += 1
            except Exception as exc:
                print(f"Skipped {file_path.name}: extraction error ({exc})")
            finally:
                if wb is not None:
                    close_source_workbook(wb)
    finally:
        try:
            app.quit()
        except Exception:
            pass

    create_output_workbook(output_file_path, empirical_rows, regression_rows)

    print(f"Output written to: {output_file_path}")
    print(f"Number of files processed: {files_processed}")
    print(f"Number of empirical rows: {len(empirical_rows)}")
    print(f"Number of regression rows: {len(regression_rows)}")


if __name__ == "__main__":
    main()
