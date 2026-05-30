#!/usr/bin/env python3
"""
Extract empirical and regression model candidates from source Excel workbooks.

This script opens each source workbook once, processes both target model sheets
while the workbook is open, and writes one consolidated output workbook.
"""

from __future__ import annotations

from dataclasses import dataclass
from datetime import date
from pathlib import Path
import re
from typing import Any

import xlwings as xw
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter


# ---------------------------------------------------------------------------
# Configure paths here
# ---------------------------------------------------------------------------
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


DAY_BY_PERIOD = {"early": 5, "mid": 15, "late": 25}
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


@dataclass
class FileLabel:
    model: str
    ticker: str
    model_period: str
    model_date: str


def read_range_2d(rng: xw.Range) -> list[list[Any]]:
    values = rng.options(ndim=2).value
    return values or []


def normalize_text(value: Any) -> str:
    if value is None:
        return ""
    return re.sub(r"\s+", " ", str(value).strip()).lower()


def to_number(value: Any) -> float | None:
    if value is None:
        return None
    if isinstance(value, (int, float)):
        return float(value)
    if isinstance(value, str):
        stripped = value.strip().replace(",", "")
        if not stripped:
            return None
        try:
            return float(stripped)
        except ValueError:
            return None
    return None


def value_or_blank(value: Any) -> Any:
    return "" if value is None else value


def safe_close_workbook(wb: xw.Book) -> None:
    # Close source files without saving under multiple xlwings variants.
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
    except Exception as exc:
        print(f"warning: failed to close workbook safely ({exc})")


def find_anchor_max(sheet: xw.Sheet) -> tuple[int, int] | None:
    used = sheet.used_range
    first_row = used.row
    first_col = used.column
    matrix = read_range_2d(used)
    for r_idx, row in enumerate(matrix):
        for c_idx, cell in enumerate(row):
            if normalize_text(cell) == "max":
                return first_row + r_idx, first_col + c_idx
    return None


def parse_file_label(file_path: Path) -> FileLabel:
    stem = file_path.stem
    parts = [p.strip() for p in stem.split(" - ")]

    ticker = parts[1].upper() if len(parts) > 1 and parts[1] else "UNKNOWN"
    period_token_source = parts[2] if len(parts) > 2 else stem
    period_token = period_token_source.split("_")[0]

    match = re.search(r"(Early|Mid|Late)([A-Za-z]{3,9})(\d{4})", period_token, re.IGNORECASE)
    if match:
        period_name = match.group(1).title()
        month_token = match.group(2)[:3].lower()
        year_int = int(match.group(3))
        month_int = MONTHS.get(month_token)
        if month_int:
            day_int = DAY_BY_PERIOD[period_name.lower()]
            model_period = f"{period_name}{match.group(2)[:3].title()}_{year_int}"
            model_date = date(year_int, month_int, day_int).isoformat()
            model = f"{ticker}_{model_period}"
            return FileLabel(model=model, ticker=ticker, model_period=model_period, model_date=model_date)

    fallback_period = re.sub(r"\s+", "", period_token) or "UnknownPeriod"
    model_period = fallback_period
    model_date = ""
    model = f"{ticker}_{model_period}"
    return FileLabel(model=model, ticker=ticker, model_period=model_period, model_date=model_date)


def infer_last_quarter(sheet: xw.Sheet, anchor_row: int, anchor_col: int) -> Any:
    # Default quarter label position relative to "max" anchor.
    return sheet.range((anchor_row - 1, anchor_col - 12)).value


def extract_empirical_candidates(
    wb: xw.Book,
    sheet_name: str,
    file_label: FileLabel,
    source_file: str,
) -> list[dict[str, Any]]:
    if sheet_name not in {sht.name for sht in wb.sheets}:
        print(f"  skipped sheet '{sheet_name}' (not found)")
        return []

    sheet = wb.sheets[sheet_name]
    anchor = find_anchor_max(sheet)
    if not anchor:
        print(f"  skipped sheet '{sheet_name}' (max anchor not found)")
        return []

    anchor_row, anchor_col = anchor
    n_quarters = 10
    data_end_row = anchor_row - 1
    helper_col = anchor_col + 10
    x_col = anchor_col - 11

    # Offsets from "max" anchor (kept explicit and centralized).
    forecast_max_offset = 0
    forecast_min_offset = 1
    forecast_value_offset = -1
    reported_sales_offset = -2
    quarterly_sales_offset = -7
    growth_rate_offset = -5
    captured_pct_offset = -4

    # Write avg penetration formulas for all rows first, then calculate once.
    for i in range(1, n_quarters + 1):
        row = anchor_row + i
        start_row = max(data_end_row - i + 1, 1)
        avg_formula = f"=AVERAGE(R{start_row}C{x_col}:R{data_end_row}C{x_col})"
        sheet.range((row, helper_col)).formula2 = avg_formula

    wb.app.calculate()
    avg_values = read_range_2d(
        sheet.range((anchor_row + 1, helper_col), (anchor_row + n_quarters, helper_col))
    )

    table_start_row = anchor_row + 1
    table_end_row = anchor_row + n_quarters
    table_start_col = anchor_col - 12
    table_end_col = anchor_col + 1
    table_values = read_range_2d(
        sheet.range((table_start_row, table_start_col), (table_end_row, table_end_col))
    )

    rows: list[dict[str, Any]] = []
    last_quarter_used = infer_last_quarter(sheet, anchor_row, anchor_col)

    for idx, row_values in enumerate(table_values):
        n_used = idx + 1
        get_rel = lambda rel: row_values[rel + 12] if 0 <= rel + 12 < len(row_values) else None

        forecast_value = get_rel(forecast_value_offset)
        forecast_max = get_rel(forecast_max_offset)
        forecast_min = get_rel(forecast_min_offset)
        reported_sales = get_rel(reported_sales_offset)
        quarterly_sales = get_rel(quarterly_sales_offset)
        growth_rate = get_rel(growth_rate_offset)
        captured_pct = get_rel(captured_pct_offset)

        if all(v in (None, "") for v in (forecast_value, forecast_max, forecast_min, reported_sales)):
            continue

        avg_pen_val = avg_values[idx][0] if idx < len(avg_values) and avg_values[idx] else None
        max_num = to_number(forecast_max)
        min_num = to_number(forecast_min)
        range_width = (max_num - min_num) if max_num is not None and min_num is not None else None

        rows.append(
            {
                "model": file_label.model,
                "ticker": file_label.ticker,
                "model_period": file_label.model_period,
                "model_date": file_label.model_date,
                "method": "empirical",
                "parameter_name": "avg_penetration_pct",
                "parameter_value": avg_pen_val,
                "num_quarters_used": n_used,
                "last_quarter_used": value_or_blank(last_quarter_used),
                "forecast_value": value_or_blank(forecast_value),
                "actual_value": value_or_blank(reported_sales),
                "forecast_max": value_or_blank(forecast_max),
                "forecast_min": value_or_blank(forecast_min),
                "range_width": value_or_blank(range_width),
                "avg_penetration_pct": value_or_blank(avg_pen_val),
                "quarterly_sales": value_or_blank(quarterly_sales),
                "reported_sales": value_or_blank(reported_sales),
                "growth_rate_pct": value_or_blank(growth_rate),
                "sales_captured_in_db_pct": value_or_blank(captured_pct),
                "source_file": source_file,
            }
        )

    return rows


def extract_regression_candidates(
    wb: xw.Book,
    sheet_name: str,
    file_label: FileLabel,
    source_file: str,
) -> list[dict[str, Any]]:
    if sheet_name not in {sht.name for sht in wb.sheets}:
        print(f"  skipped sheet '{sheet_name}' (not found)")
        return []

    sheet = wb.sheets[sheet_name]
    anchor = find_anchor_max(sheet)
    if not anchor:
        print(f"  skipped sheet '{sheet_name}' (max anchor not found)")
        return []

    anchor_row, anchor_col = anchor
    n_quarters = 10
    data_end_row = anchor_row - 1
    y_col = anchor_col - 7
    x_col = anchor_col - 11
    helper_intercept_col = anchor_col + 10
    helper_slope_col = anchor_col + 11

    forecast_value_offset = -1  # TOT FCST w/o SA
    forecast_max_offset = 0
    forecast_min_offset = 1
    actual_value_offset = -2  # optional / workbook-dependent

    # Write formulas first, then calculate once.
    for i in range(1, n_quarters + 1):
        row = anchor_row + i
        start_row = max(data_end_row - i + 1, 1)
        intercept_formula = (
            f"=INTERCEPT(R{start_row}C{y_col}:R{data_end_row}C{y_col},"
            f"R{start_row}C{x_col}:R{data_end_row}C{x_col})"
        )
        slope_formula = (
            f"=SLOPE(R{start_row}C{y_col}:R{data_end_row}C{y_col},"
            f"R{start_row}C{x_col}:R{data_end_row}C{x_col})"
        )
        sheet.range((row, helper_intercept_col)).formula2 = intercept_formula
        sheet.range((row, helper_slope_col)).formula2 = slope_formula

    wb.app.calculate()
    intercept_values = read_range_2d(
        sheet.range(
            (anchor_row + 1, helper_intercept_col),
            (anchor_row + n_quarters, helper_intercept_col),
        )
    )
    slope_values = read_range_2d(
        sheet.range(
            (anchor_row + 1, helper_slope_col),
            (anchor_row + n_quarters, helper_slope_col),
        )
    )

    table_start_row = anchor_row + 1
    table_end_row = anchor_row + n_quarters
    table_start_col = anchor_col - 12
    table_end_col = anchor_col + 1
    table_values = read_range_2d(
        sheet.range((table_start_row, table_start_col), (table_end_row, table_end_col))
    )

    rows: list[dict[str, Any]] = []
    prev_signature: tuple[Any, ...] | None = None

    for idx, row_values in enumerate(table_values):
        n_used = idx + 1
        get_rel = lambda rel: row_values[rel + 12] if 0 <= rel + 12 < len(row_values) else None

        forecast_value = get_rel(forecast_value_offset)
        forecast_max = get_rel(forecast_max_offset)
        forecast_min = get_rel(forecast_min_offset)
        actual_value = get_rel(actual_value_offset)
        intercept = intercept_values[idx][0] if idx < len(intercept_values) and intercept_values[idx] else None
        slope = slope_values[idx][0] if idx < len(slope_values) and slope_values[idx] else None

        if all(v in (None, "") for v in (forecast_value, forecast_max, forecast_min, intercept, slope)):
            continue

        max_num = to_number(forecast_max)
        min_num = to_number(forecast_min)
        range_width = (max_num - min_num) if max_num is not None and min_num is not None else None

        signature = (
            round(to_number(forecast_value) or 0.0, 8),
            round(to_number(forecast_max) or 0.0, 8),
            round(to_number(forecast_min) or 0.0, 8),
            round(to_number(intercept) or 0.0, 8),
            round(to_number(slope) or 0.0, 8),
        )
        if signature == prev_signature:
            # Prevent duplicate final row when terminal ranges converge.
            continue
        prev_signature = signature

        rows.append(
            {
                "model": file_label.model,
                "ticker": file_label.ticker,
                "model_period": file_label.model_period,
                "model_date": file_label.model_date,
                "method": "regression",
                "parameter_name": "num_quarters_used",
                "parameter_value": n_used,
                "num_quarters_used": n_used,
                "forecast_value": value_or_blank(forecast_value),
                "actual_value": value_or_blank(actual_value),
                "forecast_max": value_or_blank(forecast_max),
                "forecast_min": value_or_blank(forecast_min),
                "range_width": value_or_blank(range_width),
                "intercept": value_or_blank(intercept),
                "slope": value_or_blank(slope),
                "source_file": source_file,
            }
        )

    return rows


def choose_output_path(source_input_dir: Path, destination_dir: Path) -> Path:
    destination_dir.mkdir(parents=True, exist_ok=True)
    base_name = f"{source_input_dir.name}_PARAM.xlsx"
    output_path = destination_dir / base_name
    if not output_path.exists():
        return output_path

    suffix = 1
    while True:
        candidate = destination_dir / f"{source_input_dir.name}_PARAM.{suffix}.xlsx"
        if not candidate.exists():
            return candidate
        suffix += 1


def write_output_workbook(
    output_path: Path,
    empirical_rows: list[dict[str, Any]],
    regression_rows: list[dict[str, Any]],
) -> None:
    wb = Workbook()
    default_sheet = wb.active
    wb.remove(default_sheet)

    ws_emp = wb.create_sheet("empirical_candidates")
    ws_reg = wb.create_sheet("regression_candidates")

    ws_emp.append(EMPIRICAL_COLUMNS)
    for row in empirical_rows:
        ws_emp.append([row.get(col, "") for col in EMPIRICAL_COLUMNS])

    ws_reg.append(REGRESSION_COLUMNS)
    for row in regression_rows:
        ws_reg.append([row.get(col, "") for col in REGRESSION_COLUMNS])

    for ws in (ws_emp, ws_reg):
        for cell in ws[1]:
            cell.font = Font(bold=True)

        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions

        for col_idx, col_cells in enumerate(ws.columns, start=1):
            max_len = max(len(str(cell.value)) if cell.value is not None else 0 for cell in col_cells)
            ws.column_dimensions[get_column_letter(col_idx)].width = min(max(max_len + 2, 12), 48)

    wb.save(output_path)


def iter_source_files(directory: Path) -> tuple[list[Path], list[tuple[Path, str]]]:
    valid: list[Path] = []
    skipped: list[tuple[Path, str]] = []

    for path in sorted(directory.iterdir()):
        if not path.is_file():
            skipped.append((path, "not a file"))
            continue
        if path.name.startswith("~"):
            skipped.append((path, "temp file"))
            continue
        if path.suffix.lower() != ".xlsx":
            skipped.append((path, "not .xlsx"))
            continue
        valid.append(path)

    return valid, skipped


def main() -> None:
    if not input_dir.exists():
        print(f"input_dir does not exist: {input_dir}")
        return

    source_files, skipped_files = iter_source_files(input_dir)
    for skipped_file, reason in skipped_files:
        print(f"skipped file: {skipped_file.name} ({reason})")

    output_path = choose_output_path(input_dir, output_dir)
    all_empirical_rows: list[dict[str, Any]] = []
    all_regression_rows: list[dict[str, Any]] = []
    processed_files = 0

    app = xw.App(visible=False, add_book=False)
    app.display_alerts = False
    app.screen_updating = False
    try:
        # Manual calc mode reduces recalc overhead; explicit calculate() calls are used.
        try:
            app.api.Calculation = -4135  # xlCalculationManual
        except Exception:
            pass

        for file_path in source_files:
            print(f"processed file: {file_path.name}")
            wb = None
            try:
                wb = app.books.open(str(file_path), update_links=False)
                label = parse_file_label(file_path)

                empirical_rows = extract_empirical_candidates(
                    wb=wb,
                    sheet_name="Empirical Model",
                    file_label=label,
                    source_file=file_path.name,
                )
                regression_rows = extract_regression_candidates(
                    wb=wb,
                    sheet_name="Regression Model",
                    file_label=label,
                    source_file=file_path.name,
                )

                all_empirical_rows.extend(empirical_rows)
                all_regression_rows.extend(regression_rows)
                processed_files += 1
            except Exception as exc:
                print(f"skipped file: {file_path.name} (error: {exc})")
            finally:
                if wb is not None:
                    safe_close_workbook(wb)
    finally:
        app.quit()

    write_output_workbook(output_path, all_empirical_rows, all_regression_rows)
    print(f"output path: {output_path}")
    print(f"number of files processed: {processed_files}")
    print(f"number of empirical rows: {len(all_empirical_rows)}")
    print(f"number of regression rows: {len(all_regression_rows)}")


if __name__ == "__main__":
    main()
