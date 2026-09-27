#!/usr/bin/env python3
from __future__ import annotations

import calendar
import math
import re
from datetime import date
from pathlib import Path
from typing import Any

import xlwings as xw
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

# ---------------------------
# Configure these two values.
# ---------------------------
input_dir = Path("/path/to/input")
output_dir = Path("/path/to/output")

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


def to_float(value: Any) -> float | None:
    if value is None:
        return None
    if isinstance(value, bool):
        return float(value)
    if isinstance(value, (int, float)):
        if isinstance(value, float) and math.isnan(value):
            return None
        return float(value)
    if isinstance(value, str):
        cleaned = value.strip().replace(",", "")
        if not cleaned or cleaned.startswith("#"):
            return None
        if cleaned.endswith("%"):
            cleaned = cleaned[:-1]
        try:
            return float(cleaned)
        except ValueError:
            return None
    return None


def maybe_round(value: float | None, digits: int = 8) -> float | None:
    if value is None:
        return None
    return round(value, digits)


def resolve_bound(bound_raw: float | None, forecast_value: float | None) -> float | None:
    if bound_raw is None:
        return None
    if forecast_value is None:
        return bound_raw

    # Many templates store max/min as ratios (0.95) or percentages (95).
    if abs(bound_raw) <= 2:
        return forecast_value * bound_raw
    if abs(bound_raw) <= 200:
        return forecast_value * (bound_raw / 100.0)
    return bound_raw


def normalize_2d(values: Any) -> list[list[Any]]:
    if values is None:
        return []
    if isinstance(values, tuple):
        if not values:
            return []
        first = values[0]
        if isinstance(first, tuple):
            return [list(row) for row in values]
        return [list(values)]
    return [[values]]


def month_to_number(token: str) -> int:
    token_l = token.strip().lower()
    month_map: dict[str, int] = {}
    for month_num in range(1, 13):
        month_map[calendar.month_abbr[month_num].lower()] = month_num
        month_map[calendar.month_name[month_num].lower()] = month_num
    if token_l in month_map:
        return month_map[token_l]
    short = token_l[:3]
    if short in month_map:
        return month_map[short]
    raise ValueError(f"Unsupported month token: {token}")


def parse_file_label(file_path: Path) -> dict[str, str]:
    stem = file_path.stem
    parts = [part.strip() for part in stem.split(" - ")]

    ticker = ""
    if len(parts) >= 2 and parts[1]:
        ticker = parts[1].split("_")[0].strip()
    if not ticker:
        ticker = stem.split("_")[0].strip()

    period_segment = parts[2] if len(parts) >= 3 else stem
    period_segment = period_segment.split("_")[0]

    match = re.search(
        r"(Early|Mid|Late)\s*([A-Za-z]{3,9})\s*([12][0-9]{3})",
        period_segment,
        flags=re.IGNORECASE,
    )

    if not match:
        model_period = "unknown_period"
        model_date = ""
    else:
        period_tag = match.group(1).title()
        month_text = match.group(2)
        year = int(match.group(3))
        month_num = month_to_number(month_text)
        month_abbr = calendar.month_abbr[month_num]
        day_lookup = {"Early": 5, "Mid": 15, "Late": 25}
        day = day_lookup[period_tag]
        model_period = f"{period_tag}{month_abbr}_{year}"
        model_date = date(year, month_num, day).isoformat()

    model = f"{ticker}_{model_period}" if ticker else model_period
    return {
        "ticker": ticker,
        "model_period": model_period,
        "model_date": model_date,
        "model": model,
    }


def next_output_path(input_directory: Path, output_directory: Path) -> Path:
    base_name = f"{input_directory.name}_PARAM.xlsx"
    base_path = output_directory / base_name
    if not base_path.exists():
        return base_path

    stem = f"{input_directory.name}_PARAM"
    counter = 1
    while True:
        candidate = output_directory / f"{stem}.{counter}.xlsx"
        if not candidate.exists():
            return candidate
        counter += 1


def find_anchor_max(sheet: xw.Sheet) -> tuple[int, int] | None:
    used = sheet.used_range
    values_2d = normalize_2d(used.value)
    if not values_2d:
        return None

    base_row = used.row
    base_col = used.column

    for row_idx, row_values in enumerate(values_2d):
        for col_idx, cell_val in enumerate(row_values):
            if isinstance(cell_val, str) and cell_val.strip().lower() == "max":
                return base_row + row_idx, base_col + col_idx
    return None


def set_formula2(cell: xw.Range, formula: str) -> None:
    try:
        cell.formula2 = formula
    except Exception:
        cell.formula = formula


def safe_close_workbook(workbook: xw.Book) -> None:
    try:
        workbook.close(save=False)
        return
    except TypeError:
        pass
    except Exception:
        pass

    try:
        workbook.api.Close(SaveChanges=False)
        return
    except Exception:
        pass

    try:
        workbook.close()
    except Exception:
        pass


def process_empirical_sheet(
    workbook: xw.Book, metadata: dict[str, str], source_file: str
) -> tuple[list[dict[str, Any]], str | None]:
    try:
        sheet = workbook.sheets["Empirical Model"]
    except Exception:
        return [], "missing sheet: Empirical Model"

    anchor = find_anchor_max(sheet)
    if anchor is None:
        return [], "anchor 'max' not found in Empirical Model"

    anchor_row, anchor_col = anchor
    data_end_row = anchor_row - 2
    if data_end_row < 2:
        return [], "insufficient data rows in Empirical Model"

    # Anchor-based offsets from the Empirical Model "max" cell.
    penetration_col = anchor_col - 4
    quarterly_sales_col = anchor_col - 2
    reported_sales_col = anchor_col - 1
    quarter_label_col = anchor_col - 6

    max_value_raw = to_float(sheet.range((anchor_row, anchor_col + 1)).value)

    min_row = anchor_row + 1
    min_label = sheet.range((min_row, anchor_col)).value
    if isinstance(min_label, str) and min_label.strip().lower() != "min":
        min_row = anchor_row + 2
    min_value_raw = to_float(sheet.range((min_row, anchor_col + 1)).value)

    scratch_row = anchor_row + 8
    scratch_col = anchor_col + 8

    rows: list[dict[str, Any]] = []
    for num_quarters_used in range(1, N_QUARTERS + 1):
        data_start_row = data_end_row - num_quarters_used + 1
        if data_start_row < 2:
            break

        avg_pen_cell = sheet.range((scratch_row, scratch_col))
        quarterly_sales_sum_cell = sheet.range((scratch_row, scratch_col + 1))
        reported_sales_sum_cell = sheet.range((scratch_row, scratch_col + 2))

        set_formula2(
            avg_pen_cell,
            (
                f'=IFERROR(AVERAGE(R{data_start_row}C{penetration_col}:'
                f'R{data_end_row}C{penetration_col}),"")'
            ),
        )
        set_formula2(
            quarterly_sales_sum_cell,
            (
                f'=IFERROR(SUM(R{data_start_row}C{quarterly_sales_col}:'
                f'R{data_end_row}C{quarterly_sales_col}),"")'
            ),
        )
        set_formula2(
            reported_sales_sum_cell,
            (
                f'=IFERROR(SUM(R{data_start_row}C{reported_sales_col}:'
                f'R{data_end_row}C{reported_sales_col}),"")'
            ),
        )

        workbook.app.calculate()

        avg_pen_raw = to_float(avg_pen_cell.value)
        quarterly_sales = to_float(quarterly_sales_sum_cell.value)
        reported_sales = to_float(reported_sales_sum_cell.value)
        if avg_pen_raw is None or quarterly_sales is None:
            continue

        avg_pen_ratio = avg_pen_raw / 100.0 if abs(avg_pen_raw) > 1 else avg_pen_raw
        if avg_pen_ratio == 0:
            continue

        forecast_value = quarterly_sales / avg_pen_ratio
        actual_value = reported_sales

        forecast_max = resolve_bound(max_value_raw, forecast_value)
        forecast_min = resolve_bound(min_value_raw, forecast_value)
        range_width = None
        if forecast_max is not None and forecast_min is not None:
            range_width = forecast_max - forecast_min

        start_sales = to_float(sheet.range((data_start_row, quarterly_sales_col)).value)
        end_sales = to_float(sheet.range((data_end_row, quarterly_sales_col)).value)
        growth_rate_pct = None
        if (
            start_sales is not None
            and end_sales is not None
            and start_sales > 0
            and num_quarters_used > 1
        ):
            growth_rate_pct = ((end_sales / start_sales) ** (1 / (num_quarters_used - 1)) - 1) * 100

        sales_captured_in_db_pct = None
        if reported_sales is not None and quarterly_sales:
            sales_captured_in_db_pct = (reported_sales / quarterly_sales) * 100

        last_quarter_used = sheet.range((data_end_row, quarter_label_col)).value
        if last_quarter_used is None:
            last_quarter_used = data_end_row

        rows.append(
            {
                "model": metadata["model"],
                "ticker": metadata["ticker"],
                "model_period": metadata["model_period"],
                "model_date": metadata["model_date"],
                "method": "empirical",
                "parameter_name": "avg_penetration_pct",
                "parameter_value": avg_pen_ratio * 100,
                "num_quarters_used": num_quarters_used,
                "last_quarter_used": last_quarter_used,
                "forecast_value": forecast_value,
                "actual_value": actual_value,
                "forecast_max": forecast_max,
                "forecast_min": forecast_min,
                "range_width": range_width,
                "avg_penetration_pct": avg_pen_ratio * 100,
                "quarterly_sales": quarterly_sales,
                "reported_sales": reported_sales,
                "growth_rate_pct": growth_rate_pct,
                "sales_captured_in_db_pct": sales_captured_in_db_pct,
                "source_file": source_file,
            }
        )

    sheet.range((scratch_row, scratch_col), (scratch_row, scratch_col + 2)).value = None
    return rows, None


def process_regression_sheet(
    workbook: xw.Book, metadata: dict[str, str], source_file: str
) -> tuple[list[dict[str, Any]], str | None]:
    try:
        sheet = workbook.sheets["Regression Model"]
    except Exception:
        return [], "missing sheet: Regression Model"

    anchor = find_anchor_max(sheet)
    if anchor is None:
        return [], "anchor 'max' not found in Regression Model"

    anchor_row, anchor_col = anchor
    data_end_row = anchor_row - 2
    if data_end_row < 2:
        return [], "insufficient data rows in Regression Model"

    # Required anchor-based columns.
    y_col = anchor_col - 7
    x_col = anchor_col - 11

    scratch_row = anchor_row + 8
    scratch_col = anchor_col + 8

    rows: list[dict[str, Any]] = []
    last_signature: tuple[float | None, ...] | None = None

    actual_value_candidate = to_float(sheet.range((anchor_row - 1, anchor_col + 1)).value)

    for num_quarters_used in range(1, N_QUARTERS + 1):
        data_start_row = data_end_row - num_quarters_used + 1
        if data_start_row < 2:
            break

        intercept_cell = sheet.range((scratch_row, scratch_col))
        slope_cell = sheet.range((scratch_row, scratch_col + 1))
        max_cell = sheet.range((scratch_row, scratch_col + 2))
        min_cell = sheet.range((scratch_row, scratch_col + 3))
        latest_x_cell = sheet.range((scratch_row, scratch_col + 4))

        set_formula2(
            intercept_cell,
            (
                f'=IFERROR(INTERCEPT(R{data_start_row}C{y_col}:R{data_end_row}C{y_col},'
                f'R{data_start_row}C{x_col}:R{data_end_row}C{x_col}),"")'
            ),
        )
        set_formula2(
            slope_cell,
            (
                f'=IFERROR(SLOPE(R{data_start_row}C{y_col}:R{data_end_row}C{y_col},'
                f'R{data_start_row}C{x_col}:R{data_end_row}C{x_col}),"")'
            ),
        )
        set_formula2(
            max_cell,
            f'=IFERROR(MAX(R{data_start_row}C{y_col}:R{data_end_row}C{y_col}),"")',
        )
        set_formula2(
            min_cell,
            f'=IFERROR(MIN(R{data_start_row}C{y_col}:R{data_end_row}C{y_col}),"")',
        )
        set_formula2(latest_x_cell, f'=R{data_end_row}C{x_col}')

        workbook.app.calculate()

        intercept = to_float(intercept_cell.value)
        slope = to_float(slope_cell.value)
        forecast_max = to_float(max_cell.value)
        forecast_min = to_float(min_cell.value)
        latest_x = to_float(latest_x_cell.value)

        if intercept is None or slope is None or latest_x is None:
            continue

        forecast_total_without_sa = intercept + slope * latest_x

        range_width = None
        if forecast_max is not None and forecast_min is not None:
            range_width = forecast_max - forecast_min

        signature = (
            maybe_round(forecast_total_without_sa),
            maybe_round(intercept),
            maybe_round(slope),
            maybe_round(forecast_max),
            maybe_round(forecast_min),
        )
        if last_signature is not None and signature == last_signature:
            continue
        last_signature = signature

        rows.append(
            {
                "model": metadata["model"],
                "ticker": metadata["ticker"],
                "model_period": metadata["model_period"],
                "model_date": metadata["model_date"],
                "method": "regression",
                "parameter_name": "num_quarters_used",
                "parameter_value": num_quarters_used,
                "num_quarters_used": num_quarters_used,
                "forecast_value": forecast_total_without_sa,
                "actual_value": actual_value_candidate,
                "forecast_max": forecast_max,
                "forecast_min": forecast_min,
                "range_width": range_width,
                "intercept": intercept,
                "slope": slope,
                "source_file": source_file,
            }
        )

    sheet.range((scratch_row, scratch_col), (scratch_row, scratch_col + 4)).value = None
    return rows, None


def format_sheet(sheet, columns: list[str]) -> None:
    for col_idx in range(1, len(columns) + 1):
        sheet.cell(row=1, column=col_idx).font = Font(bold=True)

    sheet.freeze_panes = "A2"

    for col_idx, col_name in enumerate(columns, start=1):
        max_len = len(col_name)
        for row_idx in range(2, sheet.max_row + 1):
            val = sheet.cell(row=row_idx, column=col_idx).value
            if val is None:
                continue
            max_len = max(max_len, len(str(val)))
        sheet.column_dimensions[get_column_letter(col_idx)].width = min(max_len + 2, 48)

    sheet.auto_filter.ref = f"A1:{get_column_letter(sheet.max_column)}{sheet.max_row}"


def write_output_workbook(
    output_path: Path,
    empirical_rows: list[dict[str, Any]],
    regression_rows: list[dict[str, Any]],
) -> None:
    wb_out = Workbook()

    ws_empirical = wb_out.active
    ws_empirical.title = "empirical_candidates"
    ws_regression = wb_out.create_sheet("regression_candidates")

    ws_empirical.append(EMPIRICAL_COLUMNS)
    for row in empirical_rows:
        ws_empirical.append([row.get(col) for col in EMPIRICAL_COLUMNS])

    ws_regression.append(REGRESSION_COLUMNS)
    for row in regression_rows:
        ws_regression.append([row.get(col) for col in REGRESSION_COLUMNS])

    format_sheet(ws_empirical, EMPIRICAL_COLUMNS)
    format_sheet(ws_regression, REGRESSION_COLUMNS)

    wb_out.save(output_path)


def main() -> int:
    in_dir = Path(input_dir)
    out_dir = Path(output_dir)
    out_dir.mkdir(parents=True, exist_ok=True)

    if not in_dir.exists() or not in_dir.is_dir():
        print(f"Input directory not found: {in_dir}")
        return 1

    output_path = next_output_path(in_dir, out_dir)

    empirical_rows: list[dict[str, Any]] = []
    regression_rows: list[dict[str, Any]] = []
    processed_files = 0

    app = xw.App(visible=False, add_book=False)
    app.display_alerts = False
    app.screen_updating = False
    app.enable_events = False

    # Manual calc keeps Excel idle until formulas are updated.
    try:
        app.calculation = "manual"
    except Exception:
        pass

    try:
        for file_path in sorted(in_dir.iterdir()):
            if not file_path.is_file():
                continue

            if file_path.name.startswith("~"):
                print(f"Skipped {file_path.name}: temp file")
                continue

            if file_path.suffix.lower() != ".xlsx":
                print(f"Skipped {file_path.name}: not an .xlsx file")
                continue

            print(f"Processing {file_path.name}")
            workbook = None
            try:
                metadata = parse_file_label(file_path)
                workbook = app.books.open(str(file_path), update_links=False)

                empirical_result, empirical_error = process_empirical_sheet(
                    workbook, metadata, file_path.name
                )
                if empirical_error:
                    print(f"  Empirical Model skipped: {empirical_error}")
                empirical_rows.extend(empirical_result)

                regression_result, regression_error = process_regression_sheet(
                    workbook, metadata, file_path.name
                )
                if regression_error:
                    print(f"  Regression Model skipped: {regression_error}")
                regression_rows.extend(regression_result)

                processed_files += 1
            except Exception as exc:
                print(f"Skipped {file_path.name}: {exc}")
            finally:
                if workbook is not None:
                    safe_close_workbook(workbook)
    finally:
        app.quit()

    write_output_workbook(output_path, empirical_rows, regression_rows)

    print(f"Output path: {output_path}")
    print(f"Number of files processed: {processed_files}")
    print(f"Number of empirical rows: {len(empirical_rows)}")
    print(f"Number of regression rows: {len(regression_rows)}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
