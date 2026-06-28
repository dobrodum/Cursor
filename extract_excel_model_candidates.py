#!/usr/bin/env python3
from __future__ import annotations

import math
import re
from dataclasses import dataclass
from datetime import date
from pathlib import Path
from typing import Any, Iterable

import xlwings as xw
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

# -----------------------------
# User inputs
# -----------------------------
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

MAX_QUARTERS = 10


@dataclass(frozen=True)
class EmpiricalLayout:
    penetration_col_offset: int = -7
    quarterly_sales_col_offset: int = -5
    reported_sales_col_offset: int = -4
    growth_rate_col_offset: int = -3
    captured_col_offset: int = -2
    forecast_col_offset: int = -1
    max_col_offset: int = 0
    min_col_offset: int = 1
    last_quarter_col_offset: int = -8
    helper_row_offset: int = 2
    helper_col_offset: int = 6


@dataclass(frozen=True)
class RegressionLayout:
    x_col_offset: int = -11
    y_col_offset: int = -7
    forecast_col_offset: int = -1
    max_col_offset: int = 0
    min_col_offset: int = 1
    helper_row_offset: int = 2
    helper_intercept_col_offset: int = 5
    helper_slope_col_offset: int = 6
    helper_forecast_col_offset: int = 7


EMP_LAYOUT = EmpiricalLayout()
REG_LAYOUT = RegressionLayout()


def ensure_2d(values: Any) -> list[list[Any]]:
    if values is None:
        return []
    if not isinstance(values, (list, tuple)):
        return [[values]]
    if len(values) == 0:
        return []
    if isinstance(values[0], (list, tuple)):
        return [list(row) for row in values]
    return [list(values)]


def to_float(value: Any) -> float | None:
    if value is None or value == "":
        return None
    if isinstance(value, bool):
        return float(value)
    if isinstance(value, (int, float)):
        if isinstance(value, float) and math.isnan(value):
            return None
        return float(value)
    if isinstance(value, str):
        cleaned = value.strip().replace(",", "")
        if cleaned.endswith("%"):
            cleaned = cleaned[:-1]
            try:
                return float(cleaned) / 100.0
            except ValueError:
                return None
        try:
            return float(cleaned)
        except ValueError:
            return None
    return None


def to_serializable(value: Any) -> Any:
    if value is None:
        return None
    if isinstance(value, (int, float, str)):
        return value
    if hasattr(value, "isoformat"):
        return value.isoformat()
    return str(value)


def safe_sub(a: float | None, b: float | None) -> float | None:
    if a is None or b is None:
        return None
    return a - b


def safe_close_workbook(wb: xw.Book) -> None:
    try:
        wb.close(save=False)
        return
    except TypeError:
        pass
    except Exception:
        pass

    try:
        wb.close(False)
        return
    except Exception:
        pass

    try:
        wb.api.Close(SaveChanges=False)
        return
    except Exception:
        pass

    try:
        wb.api.Saved = True
        wb.close()
    except Exception:
        print(f"Warning: unable to close workbook safely: {wb.name}")


def set_formula2_r1c1(cell: xw.Range, formula_r1c1: str) -> None:
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

    cell.api.FormulaR1C1 = formula_r1c1


def get_sheet_matrix(sheet: xw.Sheet) -> tuple[int, int, list[list[Any]]]:
    used = sheet.used_range
    start_row = used.row
    start_col = used.column
    matrix = ensure_2d(used.value)
    return start_row, start_col, matrix


def find_anchor(
    matrix: list[list[Any]],
    start_row: int,
    start_col: int,
    label: str = "max",
) -> tuple[int, int] | None:
    target = label.strip().lower()
    for row_idx, row_values in enumerate(matrix):
        for col_idx, value in enumerate(row_values):
            if isinstance(value, str) and value.strip().lower() == target:
                return start_row + row_idx, start_col + col_idx
    return None


def get_numeric_cell(sheet: xw.Sheet, row: int, col: int) -> float | None:
    if row < 1 or col < 1:
        return None
    return to_float(sheet.cells(row, col).value)


def get_candidate_data_rows(
    sheet: xw.Sheet,
    start_row: int,
    matrix: list[list[Any]],
    first_col: int,
    second_col: int,
    row_limit: int | None = None,
) -> list[int]:
    rows: list[int] = []
    row_max = row_limit or (start_row + len(matrix) - 1)
    for row in range(start_row, row_max + 1):
        first_value = get_numeric_cell(sheet, row, first_col)
        second_value = get_numeric_cell(sheet, row, second_col)
        if first_value is not None and second_value is not None:
            rows.append(row)
    return rows


def parse_filename_metadata(file_name: str) -> dict[str, str] | None:
    name_without_ext = Path(file_name).stem
    parts = [p.strip() for p in name_without_ext.split(" - ")]
    if len(parts) < 3:
        return None

    ticker = parts[1]
    period_token = parts[2]
    period_match = re.search(r"(Early|Mid|Late)([A-Za-z]+)(\d{4})", period_token, flags=re.IGNORECASE)
    if not period_match:
        return None

    period_label = period_match.group(1).title()
    month_text = period_match.group(2).title()
    year_text = period_match.group(3)

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
    month_abbrev = month_text[:3]
    month_num = month_lookup.get(month_abbrev)
    if month_num is None:
        return None

    day_lookup = {"Early": 5, "Mid": 15, "Late": 25}
    day_num = day_lookup[period_label]

    year_num = int(year_text)
    model_period = f"{period_label}{month_abbrev}_{year_num}"
    model_date = date(year_num, month_num, day_num).isoformat()
    model = f"{ticker}_{model_period}"

    return {
        "model": model,
        "ticker": ticker,
        "model_period": model_period,
        "model_date": model_date,
    }


def build_output_path(input_folder: Path, output_folder: Path) -> Path:
    output_folder.mkdir(parents=True, exist_ok=True)
    base = f"{input_folder.name}_PARAM"
    candidate = output_folder / f"{base}.xlsx"
    idx = 1
    while candidate.exists():
        candidate = output_folder / f"{base}.{idx}.xlsx"
        idx += 1
    return candidate


def process_empirical_sheet(
    wb: xw.Book,
    sheet: xw.Sheet,
    metadata: dict[str, str],
    source_file: str,
) -> list[dict[str, Any]]:
    start_row, start_col, matrix = get_sheet_matrix(sheet)
    anchor = find_anchor(matrix, start_row, start_col, "max")
    if not anchor:
        print(f"Skipped empirical extraction for {source_file}: 'max' anchor not found")
        return []

    anchor_row, anchor_col = anchor
    penetration_col = anchor_col + EMP_LAYOUT.penetration_col_offset
    quarterly_sales_col = anchor_col + EMP_LAYOUT.quarterly_sales_col_offset
    reported_sales_col = anchor_col + EMP_LAYOUT.reported_sales_col_offset
    growth_col = anchor_col + EMP_LAYOUT.growth_rate_col_offset
    captured_col = anchor_col + EMP_LAYOUT.captured_col_offset
    last_quarter_col = anchor_col + EMP_LAYOUT.last_quarter_col_offset

    data_rows = get_candidate_data_rows(
        sheet=sheet,
        start_row=start_row,
        matrix=matrix,
        first_col=penetration_col,
        second_col=quarterly_sales_col,
        row_limit=anchor_row - 1,
    )
    if not data_rows:
        print(f"Skipped empirical extraction for {source_file}: no usable historical rows")
        return []

    latest_row = data_rows[-1]
    helper_cell = sheet.cells(anchor_row + EMP_LAYOUT.helper_row_offset, anchor_col + EMP_LAYOUT.helper_col_offset)
    helper_original = helper_cell.value

    quarterly_sales = get_numeric_cell(sheet, latest_row, quarterly_sales_col)
    reported_sales = get_numeric_cell(sheet, latest_row, reported_sales_col)
    growth_rate = get_numeric_cell(sheet, latest_row, growth_col)
    captured_pct = get_numeric_cell(sheet, latest_row, captured_col)
    last_quarter_used = to_serializable(sheet.cells(latest_row, last_quarter_col).value)

    rows: list[dict[str, Any]] = []
    max_n = min(MAX_QUARTERS, len(data_rows))
    for n_quarters in range(1, max_n + 1):
        start_n_row = data_rows[-n_quarters]
        avg_formula = f"=AVERAGE(R{start_n_row}C{penetration_col}:R{latest_row}C{penetration_col})"
        set_formula2_r1c1(helper_cell, avg_formula)
        wb.app.calculate()

        avg_penetration = to_float(helper_cell.value)
        table_row = anchor_row + n_quarters

        forecast_value = get_numeric_cell(sheet, table_row, anchor_col + EMP_LAYOUT.forecast_col_offset)
        if forecast_value is None and quarterly_sales is not None and avg_penetration not in (None, 0.0):
            forecast_value = quarterly_sales / avg_penetration

        actual_value = get_numeric_cell(sheet, table_row, reported_sales_col)
        if actual_value is None:
            actual_value = reported_sales

        forecast_max = get_numeric_cell(sheet, table_row, anchor_col + EMP_LAYOUT.max_col_offset)
        forecast_min = get_numeric_cell(sheet, table_row, anchor_col + EMP_LAYOUT.min_col_offset)
        if forecast_max is None:
            forecast_max = forecast_value
        if forecast_min is None:
            forecast_min = forecast_value

        rows.append(
            {
                "model": metadata["model"],
                "ticker": metadata["ticker"],
                "model_period": metadata["model_period"],
                "model_date": metadata["model_date"],
                "method": "empirical",
                "parameter_name": "avg_penetration_pct",
                "parameter_value": avg_penetration,
                "num_quarters_used": n_quarters,
                "last_quarter_used": last_quarter_used,
                "forecast_value": forecast_value,
                "actual_value": actual_value,
                "forecast_max": forecast_max,
                "forecast_min": forecast_min,
                "range_width": safe_sub(forecast_max, forecast_min),
                "avg_penetration_pct": avg_penetration,
                "quarterly_sales": quarterly_sales,
                "reported_sales": reported_sales,
                "growth_rate_pct": growth_rate,
                "sales_captured_in_db_pct": captured_pct,
                "source_file": source_file,
            }
        )

    helper_cell.value = helper_original
    return rows


def infer_next_x(x_values: Iterable[float]) -> float | None:
    values = list(x_values)
    if not values:
        return None
    if len(values) == 1:
        return values[0]
    step = values[-1] - values[-2]
    return values[-1] + step


def rounded_signature(values: Iterable[Any]) -> tuple[Any, ...]:
    out: list[Any] = []
    for value in values:
        if isinstance(value, float):
            out.append(round(value, 10))
        else:
            out.append(value)
    return tuple(out)


def process_regression_sheet(
    wb: xw.Book,
    sheet: xw.Sheet,
    metadata: dict[str, str],
    source_file: str,
) -> list[dict[str, Any]]:
    start_row, start_col, matrix = get_sheet_matrix(sheet)
    anchor = find_anchor(matrix, start_row, start_col, "max")
    if not anchor:
        print(f"Skipped regression extraction for {source_file}: 'max' anchor not found")
        return []

    anchor_row, anchor_col = anchor
    y_col = anchor_col + REG_LAYOUT.y_col_offset
    x_col = anchor_col + REG_LAYOUT.x_col_offset

    data_rows = get_candidate_data_rows(
        sheet=sheet,
        start_row=start_row,
        matrix=matrix,
        first_col=x_col,
        second_col=y_col,
        row_limit=anchor_row - 1,
    )
    if len(data_rows) < 2:
        print(f"Skipped regression extraction for {source_file}: not enough rows for regression")
        return []

    latest_row = data_rows[-1]
    x_history = [get_numeric_cell(sheet, row, x_col) for row in data_rows]
    x_values = [x for x in x_history if x is not None]
    forecast_x = get_numeric_cell(sheet, latest_row + 1, x_col)
    if forecast_x is None:
        forecast_x = infer_next_x(x_values)

    helper_intercept = sheet.cells(anchor_row + REG_LAYOUT.helper_row_offset, anchor_col + REG_LAYOUT.helper_intercept_col_offset)
    helper_slope = sheet.cells(anchor_row + REG_LAYOUT.helper_row_offset, anchor_col + REG_LAYOUT.helper_slope_col_offset)
    helper_forecast = sheet.cells(anchor_row + REG_LAYOUT.helper_row_offset, anchor_col + REG_LAYOUT.helper_forecast_col_offset)

    original_intercept = helper_intercept.value
    original_slope = helper_slope.value
    original_forecast = helper_forecast.value

    rows: list[dict[str, Any]] = []
    previous_signature: tuple[Any, ...] | None = None
    max_n = min(MAX_QUARTERS, len(data_rows))
    for n_quarters in range(1, max_n + 1):
        start_n_row = data_rows[-n_quarters]
        intercept_formula = (
            f"=INTERCEPT("
            f"R{start_n_row}C{y_col}:R{latest_row}C{y_col},"
            f"R{start_n_row}C{x_col}:R{latest_row}C{x_col}"
            f")"
        )
        slope_formula = (
            f"=SLOPE("
            f"R{start_n_row}C{y_col}:R{latest_row}C{y_col},"
            f"R{start_n_row}C{x_col}:R{latest_row}C{x_col}"
            f")"
        )

        set_formula2_r1c1(helper_intercept, intercept_formula)
        set_formula2_r1c1(helper_slope, slope_formula)
        if forecast_x is not None:
            forecast_formula = (
                f"=R{helper_intercept.row}C{helper_intercept.column}"
                f"+R{helper_slope.row}C{helper_slope.column}*{forecast_x}"
            )
        else:
            forecast_formula = f"=R{latest_row}C{y_col}"
        set_formula2_r1c1(helper_forecast, forecast_formula)

        wb.app.calculate()

        intercept = to_float(helper_intercept.value)
        slope = to_float(helper_slope.value)
        forecast_calc = to_float(helper_forecast.value)

        table_row = anchor_row + n_quarters
        forecast_total_without_sa = get_numeric_cell(sheet, table_row, anchor_col + REG_LAYOUT.forecast_col_offset)
        if forecast_total_without_sa is None:
            forecast_total_without_sa = forecast_calc

        forecast_max = get_numeric_cell(sheet, table_row, anchor_col + REG_LAYOUT.max_col_offset)
        forecast_min = get_numeric_cell(sheet, table_row, anchor_col + REG_LAYOUT.min_col_offset)
        if forecast_max is None:
            forecast_max = forecast_total_without_sa
        if forecast_min is None:
            forecast_min = forecast_total_without_sa

        candidate_signature = rounded_signature(
            [
                n_quarters,
                forecast_total_without_sa,
                forecast_max,
                forecast_min,
                intercept,
                slope,
            ]
        )
        if candidate_signature == previous_signature:
            continue
        previous_signature = candidate_signature

        rows.append(
            {
                "model": metadata["model"],
                "ticker": metadata["ticker"],
                "model_period": metadata["model_period"],
                "model_date": metadata["model_date"],
                "method": "regression",
                "parameter_name": "num_quarters_used",
                "parameter_value": n_quarters,
                "num_quarters_used": n_quarters,
                "forecast_value": forecast_total_without_sa,
                "actual_value": None,
                "forecast_max": forecast_max,
                "forecast_min": forecast_min,
                "range_width": safe_sub(forecast_max, forecast_min),
                "intercept": intercept,
                "slope": slope,
                "source_file": source_file,
            }
        )

    helper_intercept.value = original_intercept
    helper_slope.value = original_slope
    helper_forecast.value = original_forecast
    return rows


def format_worksheet(ws: Any, headers: list[str]) -> None:
    ws.freeze_panes = "A2"
    for cell in ws[1]:
        cell.font = Font(bold=True)
    ws.auto_filter.ref = ws.dimensions

    for col_idx, header in enumerate(headers, start=1):
        max_len = len(header)
        for row_idx in range(2, ws.max_row + 1):
            value = ws.cell(row=row_idx, column=col_idx).value
            if value is None:
                continue
            max_len = max(max_len, len(str(value)))
        ws.column_dimensions[get_column_letter(col_idx)].width = min(max_len + 2, 40)


def write_output_workbook(
    output_path: Path,
    empirical_rows: list[dict[str, Any]],
    regression_rows: list[dict[str, Any]],
) -> None:
    wb = Workbook()
    wb.remove(wb.active)

    empirical_ws = wb.create_sheet("empirical_candidates")
    regression_ws = wb.create_sheet("regression_candidates")

    empirical_ws.append(EMPIRICAL_COLUMNS)
    for row in empirical_rows:
        empirical_ws.append([row.get(col) for col in EMPIRICAL_COLUMNS])
    format_worksheet(empirical_ws, EMPIRICAL_COLUMNS)

    regression_ws.append(REGRESSION_COLUMNS)
    for row in regression_rows:
        regression_ws.append([row.get(col) for col in REGRESSION_COLUMNS])
    format_worksheet(regression_ws, REGRESSION_COLUMNS)

    wb.save(output_path)


def process_all_files() -> None:
    if not input_dir.exists():
        raise FileNotFoundError(f"Input folder not found: {input_dir.resolve()}")

    output_path = build_output_path(input_dir.resolve(), output_dir.resolve())
    empirical_rows: list[dict[str, Any]] = []
    regression_rows: list[dict[str, Any]] = []
    files_processed = 0

    app = xw.App(visible=False, add_book=False)
    app.display_alerts = False
    app.screen_updating = False
    try:
        app.calculation = "manual"
    except Exception:
        pass

    try:
        for file_path in sorted(input_dir.iterdir(), key=lambda p: p.name.lower()):
            if not file_path.is_file():
                continue
            if file_path.name.startswith("~"):
                print(f"Skipped {file_path.name}: temporary file")
                continue
            if file_path.suffix.lower() != ".xlsx":
                print(f"Skipped {file_path.name}: not an .xlsx file")
                continue

            metadata = parse_filename_metadata(file_path.name)
            if not metadata:
                print(f"Skipped {file_path.name}: filename does not match expected pattern")
                continue

            print(f"Processing {file_path.name}")
            wb: xw.Book | None = None
            try:
                wb = app.books.open(str(file_path), update_links=False)
                try:
                    empirical_sheet = wb.sheets["Empirical Model"]
                except Exception:
                    empirical_sheet = None
                try:
                    regression_sheet = wb.sheets["Regression Model"]
                except Exception:
                    regression_sheet = None

                if empirical_sheet is None and regression_sheet is None:
                    print(f"Skipped {file_path.name}: required model sheets not found")
                    continue

                if empirical_sheet is not None:
                    empirical_rows.extend(
                        process_empirical_sheet(
                            wb=wb,
                            sheet=empirical_sheet,
                            metadata=metadata,
                            source_file=file_path.name,
                        )
                    )

                if regression_sheet is not None:
                    regression_rows.extend(
                        process_regression_sheet(
                            wb=wb,
                            sheet=regression_sheet,
                            metadata=metadata,
                            source_file=file_path.name,
                        )
                    )

                files_processed += 1
            except Exception as exc:
                print(f"Skipped {file_path.name}: {exc}")
            finally:
                if wb is not None:
                    safe_close_workbook(wb)
    finally:
        app.quit()

    write_output_workbook(
        output_path=output_path,
        empirical_rows=empirical_rows,
        regression_rows=regression_rows,
    )

    print(f"Output path: {output_path}")
    print(f"Number of files processed: {files_processed}")
    print(f"Number of empirical rows: {len(empirical_rows)}")
    print(f"Number of regression rows: {len(regression_rows)}")


if __name__ == "__main__":
    process_all_files()
