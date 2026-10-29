from __future__ import annotations

import calendar
import re
from dataclasses import dataclass
from datetime import date
from pathlib import Path
from statistics import mean
from typing import Any

import xlwings as xw
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

# =========================
# User-configurable paths
# =========================
input_dir = Path("input")
output_dir = Path("output")


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

# Anchor-relative offsets for both model sheets.
MAX_ROWS_TO_USE = 10
LAST_DATA_ROW_OFFSET = -2
QUARTER_COL_OFFSET = -12
DB_SALES_COL_OFFSET = -11
REPORTED_SALES_COL_OFFSET = -7
GROWTH_RATE_COL_OFFSET = -6
CAPTURE_PCT_COL_OFFSET = -5


@dataclass(frozen=True)
class FileMeta:
    ticker: str
    model_period: str
    model_date: str
    model: str


def to_float(value: Any) -> float | None:
    if value is None:
        return None
    try:
        return float(value)
    except (TypeError, ValueError):
        return None


def normalize_2d(values: Any) -> list[list[Any]]:
    if values is None:
        return []
    if not isinstance(values, list):
        return [[values]]
    if not values:
        return []
    if isinstance(values[0], list):
        return values
    return [values]


def normalize_col(values: Any) -> list[Any]:
    if values is None:
        return []
    if not isinstance(values, list):
        return [values]
    if not values:
        return []
    if isinstance(values[0], list):
        return [row[0] if row else None for row in values]
    return values


def round_sig(value: float | None, digits: int = 10) -> float | None:
    if value is None:
        return None
    return round(value, digits)


def parse_file_meta(file_path: Path) -> FileMeta:
    stem = file_path.stem

    ticker = ""
    ticker_match = re.search(r"-\s*([A-Za-z0-9]+)\s*-\s*", stem)
    if ticker_match:
        ticker = ticker_match.group(1).upper()

    period_match = re.search(
        r"(?i)\b(Early|Mid|Late)(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)(20\d{2})\b",
        stem,
    )

    model_period = ""
    model_date = ""
    if period_match:
        bucket = period_match.group(1).capitalize()
        month_abbr = period_match.group(2).title()
        year = int(period_match.group(3))
        month = list(calendar.month_abbr).index(month_abbr)
        day_map = {"Early": 5, "Mid": 15, "Late": 25}
        month_day = day_map[bucket]
        model_period = f"{bucket}{month_abbr}_{year}"
        model_date = date(year, month, month_day).isoformat()

    if ticker and model_period:
        model = f"{ticker}_{model_period}"
    elif ticker:
        model = ticker
    else:
        model = stem

    return FileMeta(
        ticker=ticker,
        model_period=model_period,
        model_date=model_date,
        model=model,
    )


def next_output_path(source_dir: Path, target_dir: Path) -> Path:
    base_name = f"{source_dir.name}_PARAM"
    candidate = target_dir / f"{base_name}.xlsx"
    suffix = 1
    while candidate.exists():
        candidate = target_dir / f"{base_name}.{suffix}.xlsx"
        suffix += 1
    return candidate


def get_sheet_or_none(workbook: xw.Book, sheet_name: str) -> xw.Sheet | None:
    try:
        return workbook.sheets[sheet_name]
    except Exception:
        return None


def find_max_anchor(sheet: xw.Sheet) -> tuple[int, int] | None:
    used = sheet.used_range
    values = normalize_2d(used.value)
    if not values:
        return None

    for row_idx, row in enumerate(values):
        for col_idx, value in enumerate(row):
            if isinstance(value, str) and value.strip().lower() == "max":
                return used.row + row_idx, used.column + col_idx
    return None


def set_formula2(cell: xw.Range, formula: str) -> None:
    try:
        cell.formula2 = formula
    except Exception:
        # Fallback if Formula2 is unavailable in a specific Excel build.
        cell.formula = formula


def safe_close_source_wb(workbook: xw.Book) -> None:
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
        workbook.api.Close(False)
        return
    except Exception:
        pass

    try:
        workbook.close()
    except Exception:
        pass


def safe_subtract(a: float | None, b: float | None) -> float | None:
    if a is None or b is None:
        return None
    return a - b


def process_empirical_sheet(
    workbook: xw.Book, sheet: xw.Sheet, meta: FileMeta, source_file: str
) -> list[dict[str, Any]]:
    anchor = find_max_anchor(sheet)
    if not anchor:
        print(f"  - skipped empirical extraction: 'max' anchor not found")
        return []

    anchor_row, anchor_col = anchor
    last_data_row = anchor_row + LAST_DATA_ROW_OFFSET
    first_data_row = last_data_row - (MAX_ROWS_TO_USE - 1)
    if first_data_row < 1:
        print("  - skipped empirical extraction: insufficient data rows above anchor")
        return []

    quarter_col = anchor_col + QUARTER_COL_OFFSET
    db_sales_col = anchor_col + DB_SALES_COL_OFFSET
    reported_sales_col = anchor_col + REPORTED_SALES_COL_OFFSET
    growth_col = anchor_col + GROWTH_RATE_COL_OFFSET
    captured_col = anchor_col + CAPTURE_PCT_COL_OFFSET

    if min(quarter_col, db_sales_col, reported_sales_col, growth_col, captured_col) < 1:
        print("  - skipped empirical extraction: computed source column out of bounds")
        return []

    min_col = min(quarter_col, db_sales_col, reported_sales_col, growth_col, captured_col)
    max_col = max(quarter_col, db_sales_col, reported_sales_col, growth_col, captured_col)
    block = normalize_2d(sheet.range((first_data_row, min_col), (last_data_row, max_col)).value)
    if len(block) < MAX_ROWS_TO_USE:
        print("  - skipped empirical extraction: unable to read expected data block")
        return []

    def col_slice(abs_col: int) -> list[Any]:
        idx = abs_col - min_col
        return [row[idx] if idx < len(row) else None for row in block]

    quarter_values = col_slice(quarter_col)
    db_sales_values = [to_float(v) for v in col_slice(db_sales_col)]
    reported_sales_values = [to_float(v) for v in col_slice(reported_sales_col)]
    growth_values = [to_float(v) for v in col_slice(growth_col)]
    captured_values = [to_float(v) for v in col_slice(captured_col)]

    base_forecast_max = to_float(sheet.range((anchor_row, anchor_col + 1)).value)
    base_forecast_min = None
    for row_offset in (1, 2, 3):
        probe_label = sheet.range((anchor_row + row_offset, anchor_col)).value
        if isinstance(probe_label, str) and probe_label.strip().lower() == "min":
            base_forecast_min = to_float(sheet.range((anchor_row + row_offset, anchor_col + 1)).value)
            break
    if base_forecast_min is None:
        base_forecast_min = to_float(sheet.range((anchor_row + 1, anchor_col + 1)).value)

    calc_cell = sheet.range((anchor_row + 6, anchor_col + 2))
    rows: list[dict[str, Any]] = []

    for quarters_used in range(1, MAX_ROWS_TO_USE + 1):
        start_row = last_data_row - quarters_used + 1
        avg_formula = (
            f"=AVERAGE(R{start_row}C{captured_col}:R{last_data_row}C{captured_col})"
        )
        set_formula2(calc_cell, avg_formula)
        workbook.app.calculate()

        avg_penetration = to_float(calc_cell.value)
        if avg_penetration is None:
            sample = [v for v in captured_values[-quarters_used:] if v is not None]
            avg_penetration = mean(sample) if sample else None

        last_idx = len(quarter_values) - 1
        quarterly_sales = db_sales_values[last_idx]
        reported_sales = reported_sales_values[last_idx]
        growth_rate_pct = growth_values[last_idx]
        sales_captured_pct = captured_values[last_idx]

        forecast_value = None
        if quarterly_sales is not None and avg_penetration not in (None, 0):
            forecast_value = quarterly_sales / avg_penetration

        forecast_max = base_forecast_max
        forecast_min = base_forecast_min
        if forecast_value is not None and forecast_max is None:
            forecast_max = forecast_value * 1.05
        if forecast_value is not None and forecast_min is None:
            forecast_min = forecast_value * 0.95

        rows.append(
            {
                "model": meta.model,
                "ticker": meta.ticker,
                "model_period": meta.model_period,
                "model_date": meta.model_date,
                "method": "empirical",
                "parameter_name": "avg_penetration_pct",
                "parameter_value": avg_penetration,
                "num_quarters_used": quarters_used,
                "last_quarter_used": quarter_values[last_idx],
                "forecast_value": forecast_value,
                "actual_value": reported_sales,
                "forecast_max": forecast_max,
                "forecast_min": forecast_min,
                "range_width": safe_subtract(forecast_max, forecast_min),
                "avg_penetration_pct": avg_penetration,
                "quarterly_sales": quarterly_sales,
                "reported_sales": reported_sales,
                "growth_rate_pct": growth_rate_pct,
                "sales_captured_in_db_pct": sales_captured_pct,
                "source_file": source_file,
            }
        )

    return rows


def process_regression_sheet(
    workbook: xw.Book, sheet: xw.Sheet, meta: FileMeta, source_file: str
) -> list[dict[str, Any]]:
    anchor = find_max_anchor(sheet)
    if not anchor:
        print(f"  - skipped regression extraction: 'max' anchor not found")
        return []

    anchor_row, anchor_col = anchor
    y_col = anchor_col - 7
    x_col = anchor_col - 11
    last_data_row = anchor_row + LAST_DATA_ROW_OFFSET
    first_data_row = last_data_row - (MAX_ROWS_TO_USE - 1)
    if min(y_col, x_col, first_data_row) < 1:
        print("  - skipped regression extraction: computed source range out of bounds")
        return []

    x_values = [to_float(v) for v in normalize_col(sheet.range((first_data_row, x_col), (last_data_row, x_col)).value)]
    y_values = [to_float(v) for v in normalize_col(sheet.range((first_data_row, y_col), (last_data_row, y_col)).value)]
    if len(x_values) < MAX_ROWS_TO_USE or len(y_values) < MAX_ROWS_TO_USE:
        print("  - skipped regression extraction: unable to read expected x/y data")
        return []

    calc_col = anchor_col + 2
    calc_base_row = anchor_row + 6
    intercept_cell = sheet.range((calc_base_row, calc_col))
    slope_cell = sheet.range((calc_base_row + 1, calc_col))
    forecast_cell = sheet.range((calc_base_row + 2, calc_col))
    max_cell = sheet.range((calc_base_row + 3, calc_col))
    min_cell = sheet.range((calc_base_row + 4, calc_col))

    rows: list[dict[str, Any]] = []
    previous_signature: tuple[float | None, ...] | None = None
    latest_actual = y_values[-1]

    for quarters_used in range(2, MAX_ROWS_TO_USE + 1):
        start_row = last_data_row - quarters_used + 1
        intercept_formula = (
            f"=INTERCEPT(R{start_row}C{y_col}:R{last_data_row}C{y_col},"
            f"R{start_row}C{x_col}:R{last_data_row}C{x_col})"
        )
        slope_formula = (
            f"=SLOPE(R{start_row}C{y_col}:R{last_data_row}C{y_col},"
            f"R{start_row}C{x_col}:R{last_data_row}C{x_col})"
        )
        forecast_formula = (
            f"=R{last_data_row}C{x_col}*R{calc_base_row + 1}C{calc_col}"
            f"+R{calc_base_row}C{calc_col}"
        )

        set_formula2(intercept_cell, intercept_formula)
        set_formula2(slope_cell, slope_formula)
        set_formula2(forecast_cell, forecast_formula)
        set_formula2(max_cell, f"=R{calc_base_row + 2}C{calc_col}*1.05")
        set_formula2(min_cell, f"=R{calc_base_row + 2}C{calc_col}*0.95")
        workbook.app.calculate()

        intercept = to_float(intercept_cell.value)
        slope = to_float(slope_cell.value)
        forecast_total_without_sa = to_float(forecast_cell.value)
        forecast_max = to_float(max_cell.value)
        forecast_min = to_float(min_cell.value)

        current_signature = (
            round_sig(intercept),
            round_sig(slope),
            round_sig(forecast_total_without_sa),
            round_sig(forecast_max),
            round_sig(forecast_min),
        )
        if previous_signature == current_signature:
            continue
        previous_signature = current_signature

        rows.append(
            {
                "model": meta.model,
                "ticker": meta.ticker,
                "model_period": meta.model_period,
                "model_date": meta.model_date,
                "method": "regression",
                "parameter_name": "num_quarters_used",
                "parameter_value": quarters_used,
                "num_quarters_used": quarters_used,
                "forecast_value": forecast_total_without_sa,
                "actual_value": latest_actual,
                "forecast_max": forecast_max,
                "forecast_min": forecast_min,
                "range_width": safe_subtract(forecast_max, forecast_min),
                "intercept": intercept,
                "slope": slope,
                "source_file": source_file,
            }
        )

    return rows


def format_sheet(
    sheet, columns: list[str], rows: list[dict[str, Any]], min_width: int = 12, max_width: int = 50
) -> None:
    sheet.append(columns)
    for row in rows:
        sheet.append([row.get(column) for column in columns])

    for cell in sheet[1]:
        cell.font = Font(bold=True)

    sheet.freeze_panes = "A2"
    sheet.auto_filter.ref = sheet.dimensions

    for col_idx, column in enumerate(columns, start=1):
        values = [column]
        for row in rows:
            value = row.get(column)
            values.append("" if value is None else str(value))
        width = max(min_width, min(max_width, max(len(v) for v in values) + 2))
        sheet.column_dimensions[get_column_letter(col_idx)].width = width


def write_output(
    output_path: Path,
    empirical_rows: list[dict[str, Any]],
    regression_rows: list[dict[str, Any]],
) -> None:
    workbook = Workbook()
    empirical_sheet = workbook.active
    empirical_sheet.title = "empirical_candidates"
    regression_sheet = workbook.create_sheet("regression_candidates")

    format_sheet(empirical_sheet, EMPIRICAL_COLUMNS, empirical_rows)
    format_sheet(regression_sheet, REGRESSION_COLUMNS, regression_rows)

    workbook.save(output_path)


def run() -> None:
    source_dir = input_dir.expanduser().resolve()
    target_dir = output_dir.expanduser().resolve()
    target_dir.mkdir(parents=True, exist_ok=True)

    if not source_dir.exists():
        raise FileNotFoundError(f"Input directory does not exist: {source_dir}")

    output_path = next_output_path(source_dir, target_dir)
    files = sorted(source_dir.iterdir())

    empirical_rows: list[dict[str, Any]] = []
    regression_rows: list[dict[str, Any]] = []
    files_processed = 0

    app = xw.App(visible=False, add_book=False)
    app.display_alerts = False
    app.screen_updating = False
    try:
        try:
            app.calculation = "manual"
        except Exception:
            pass

        for file_path in files:
            if not file_path.is_file():
                continue

            if file_path.name.startswith("~"):
                print(f"Skipped: {file_path.name} (temporary file)")
                continue

            if file_path.suffix.lower() != ".xlsx":
                print(f"Skipped: {file_path.name} (not .xlsx)")
                continue

            print(f"Processing: {file_path.name}")
            workbook = None
            try:
                workbook = app.books.open(str(file_path), update_links=False)
                meta = parse_file_meta(file_path)

                empirical_sheet = get_sheet_or_none(workbook, "Empirical Model")
                if empirical_sheet is None:
                    print("  - skipped empirical extraction: sheet 'Empirical Model' missing")
                else:
                    empirical_rows.extend(
                        process_empirical_sheet(workbook, empirical_sheet, meta, file_path.name)
                    )

                regression_sheet = get_sheet_or_none(workbook, "Regression Model")
                if regression_sheet is None:
                    print("  - skipped regression extraction: sheet 'Regression Model' missing")
                else:
                    regression_rows.extend(
                        process_regression_sheet(workbook, regression_sheet, meta, file_path.name)
                    )

                files_processed += 1
            except Exception as exc:
                print(f"Skipped: {file_path.name} (error: {exc})")
            finally:
                if workbook is not None:
                    safe_close_source_wb(workbook)
    finally:
        app.quit()

    write_output(output_path, empirical_rows, regression_rows)

    print(f"Output written to: {output_path}")
    print(f"Files processed: {files_processed}")
    print(f"Empirical rows: {len(empirical_rows)}")
    print(f"Regression rows: {len(regression_rows)}")


if __name__ == "__main__":
    run()
