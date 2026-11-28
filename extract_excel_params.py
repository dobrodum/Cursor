#!/usr/bin/env python3
from __future__ import annotations

import math
import re
import statistics
import traceback
from dataclasses import dataclass
from datetime import date
from pathlib import Path
from typing import Any, Iterable

import xlwings as xw
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

# =============================
# Configure paths here
# =============================
input_dir = Path("/path/to/input")
output_dir = Path("/path/to/output")

# =============================
# Runtime constants
# =============================
EMPIRICAL_SOURCE_SHEET = "Empirical Model"
REGRESSION_SOURCE_SHEET = "Regression Model"
EMPIRICAL_OUTPUT_SHEET = "empirical_candidates"
REGRESSION_OUTPUT_SHEET = "regression_candidates"
MAX_QUARTERS_TO_TEST = 10

# Anchor-based offsets from the "max" cell.
X_COL_OFFSET = -11
Y_COL_OFFSET = -7

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

MONTH_TO_NUMBER = {
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
PERIOD_PATTERN = re.compile(
    r"^(?P<phase>Early|Mid|Late)(?P<month>Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)(?P<year>\d{4})$",
    re.IGNORECASE,
)
OUTPUT_STEM_PATTERN_TEMPLATE = r"^{name}_PARAM(?:\.\d+)?\.xlsx$"


@dataclass
class FileLabel:
    ticker: str
    model_period: str
    model_date: str
    model: str


def to_float(value: Any) -> float | None:
    if value is None or value == "":
        return None
    if isinstance(value, bool):
        return float(value)
    if isinstance(value, (int, float)):
        if isinstance(value, float) and (math.isnan(value) or math.isinf(value)):
            return None
        return float(value)
    try:
        return float(str(value).replace(",", ""))
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


def normalize_1d(values: Any) -> list[Any]:
    if values is None:
        return []
    if isinstance(values, list):
        return values
    return [values]


def parse_file_label(file_name: str) -> FileLabel:
    stem = Path(file_name).stem
    parts = [part.strip() for part in stem.split(" - ")]

    ticker = ""
    raw_period = ""

    if len(parts) >= 2:
        ticker = parts[1].upper()
    if len(parts) >= 3:
        raw_period = parts[2].split("_")[0]

    model_period = raw_period
    model_date = ""

    period_match = PERIOD_PATTERN.match(raw_period)
    if period_match:
        phase = period_match.group("phase").title()
        month = period_match.group("month").title()
        year = int(period_match.group("year"))
        model_period = f"{phase}{month}_{year}"
        day = PERIOD_DAY[phase.lower()]
        month_num = MONTH_TO_NUMBER[month.lower()]
        model_date = date(year, month_num, day).isoformat()

    if ticker and model_period:
        model = f"{ticker}_{model_period}"
    elif ticker:
        model = ticker
    else:
        model = stem

    return FileLabel(
        ticker=ticker,
        model_period=model_period,
        model_date=model_date,
        model=model,
    )


def resolve_output_path(input_path: Path, output_path: Path) -> Path:
    output_path.mkdir(parents=True, exist_ok=True)
    base_name = f"{input_path.name}_PARAM.xlsx"
    candidate = output_path / base_name
    if not candidate.exists():
        return candidate

    idx = 1
    while True:
        candidate = output_path / f"{input_path.name}_PARAM.{idx}.xlsx"
        if not candidate.exists():
            return candidate
        idx += 1


def find_anchor(sheet: xw.Sheet, token: str = "max") -> tuple[int, int] | None:
    used = sheet.used_range
    matrix = normalize_2d(used.value)
    if not matrix:
        return None

    token_norm = token.strip().lower()
    start_row = used.row
    start_col = used.column

    for r_idx, row in enumerate(matrix):
        for c_idx, value in enumerate(row):
            if isinstance(value, str) and value.strip().lower() == token_norm:
                return (start_row + r_idx, start_col + c_idx)
    return None


def get_sheet_or_none(wb: xw.Book, sheet_name: str) -> xw.Sheet | None:
    try:
        return wb.sheets[sheet_name]
    except Exception:
        return None


def collect_numeric_series(
    sheet: xw.Sheet,
    anchor_row: int,
    anchor_col: int,
    x_offset: int = X_COL_OFFSET,
    y_offset: int = Y_COL_OFFSET,
) -> tuple[int, int, list[tuple[int, float, float]]]:
    x_col = anchor_col + x_offset
    y_col = anchor_col + y_offset
    end_row = anchor_row - 1

    if x_col < 1 or y_col < 1 or end_row < 1:
        return x_col, y_col, []

    x_values = normalize_1d(sheet.range((1, x_col), (end_row, x_col)).value)
    y_values = normalize_1d(sheet.range((1, y_col), (end_row, y_col)).value)

    if len(x_values) < end_row:
        x_values.extend([None] * (end_row - len(x_values)))
    if len(y_values) < end_row:
        y_values.extend([None] * (end_row - len(y_values)))

    series: list[tuple[int, float, float]] = []
    found_first_numeric_row = False

    for row_idx in range(end_row, 0, -1):
        x_val = to_float(x_values[row_idx - 1])
        y_val = to_float(y_values[row_idx - 1])

        if x_val is not None and y_val is not None:
            series.append((row_idx, x_val, y_val))
            found_first_numeric_row = True
            continue

        if found_first_numeric_row:
            break

    series.reverse()
    return x_col, y_col, series


def r1c1_range(start_row: int, end_row: int, col: int) -> str:
    return f"R{start_row}C{col}:R{end_row}C{col}"


def safe_close_source_workbook(wb: xw.Book) -> None:
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
        wb.api.Close(False)
    except Exception:
        pass


def safe_quit_app(app: xw.App) -> None:
    try:
        app.quit()
    except Exception:
        pass


def extract_empirical_rows(
    wb: xw.Book,
    source_file: str,
    label: FileLabel,
) -> list[dict[str, Any]]:
    sheet = get_sheet_or_none(wb, EMPIRICAL_SOURCE_SHEET)
    if sheet is None:
        print(f"  - skipped empirical in {source_file}: sheet '{EMPIRICAL_SOURCE_SHEET}' not found")
        return []

    anchor = find_anchor(sheet, "max")
    if anchor is None:
        print(f"  - skipped empirical in {source_file}: could not find 'max' anchor")
        return []

    anchor_row, anchor_col = anchor
    x_col, y_col, series = collect_numeric_series(sheet, anchor_row, anchor_col)

    if not series:
        print(f"  - skipped empirical in {source_file}: no numeric history near anchor")
        return []

    rows: list[dict[str, Any]] = []
    n_max = min(MAX_QUARTERS_TO_TEST, len(series))

    temp_avg_cell = sheet.cells(anchor_row + 2, anchor_col + 6)
    last_quarter_label_col = max(1, x_col - 1)

    for n_quarters in range(1, n_max + 1):
        window = series[-n_quarters:]
        start_row = window[0][0]
        end_row = window[-1][0]

        avg_formula = (
            f"=LET("
            f"x,{r1c1_range(start_row, end_row, x_col)},"
            f"y,{r1c1_range(start_row, end_row, y_col)},"
            f"AVERAGE(IFERROR(x/y,0))"
            f")"
        )
        temp_avg_cell.formula2 = avg_formula
        wb.app.calculate()
        avg_penetration = to_float(temp_avg_cell.value)

        quarterly_values = [x for _, x, _ in window]
        reported_values = [y for _, _, y in window]
        penetration_values = [x / y for x, y in zip(quarterly_values, reported_values) if y != 0]

        if avg_penetration is None and penetration_values:
            avg_penetration = sum(penetration_values) / len(penetration_values)

        quarterly_sales = sum(quarterly_values)
        reported_sales = sum(reported_values)
        sales_captured_pct = (
            quarterly_sales / reported_sales if reported_sales not in (0, None) else None
        )

        forecast_value = None
        if avg_penetration not in (None, 0):
            forecast_value = quarterly_sales / avg_penetration

        penetration_std = statistics.pstdev(penetration_values) if len(penetration_values) > 1 else 0.0

        forecast_max = None
        forecast_min = None
        if avg_penetration is not None:
            lower_penetration = max(avg_penetration - penetration_std, 1e-9)
            upper_penetration = avg_penetration + penetration_std
            if lower_penetration > 0:
                forecast_max = quarterly_sales / lower_penetration
            if upper_penetration > 0:
                forecast_min = quarterly_sales / upper_penetration

        if forecast_min is not None and forecast_max is not None and forecast_min > forecast_max:
            forecast_min, forecast_max = forecast_max, forecast_min

        range_width = None
        if forecast_min is not None and forecast_max is not None:
            range_width = forecast_max - forecast_min

        growth_rate_pct = None
        if len(reported_values) > 1 and reported_values[0] != 0:
            growth_rate_pct = (reported_values[-1] / reported_values[0]) - 1

        last_quarter_used = sheet.cells(end_row, last_quarter_label_col).value

        rows.append(
            {
                "model": label.model,
                "ticker": label.ticker,
                "model_period": label.model_period,
                "model_date": label.model_date,
                "method": "empirical",
                "parameter_name": "avg_penetration_pct",
                "parameter_value": avg_penetration,
                "num_quarters_used": n_quarters,
                "last_quarter_used": last_quarter_used,
                "forecast_value": forecast_value,
                "actual_value": reported_sales,
                "forecast_max": forecast_max,
                "forecast_min": forecast_min,
                "range_width": range_width,
                "avg_penetration_pct": avg_penetration,
                "quarterly_sales": quarterly_sales,
                "reported_sales": reported_sales,
                "growth_rate_pct": growth_rate_pct,
                "sales_captured_in_db_pct": sales_captured_pct,
                "source_file": source_file,
            }
        )

    temp_avg_cell.value = None
    return rows


def regression_rows_match(a: dict[str, Any], b: dict[str, Any], tol: float = 1e-9) -> bool:
    keys = (
        "num_quarters_used",
        "forecast_value",
        "forecast_max",
        "forecast_min",
        "intercept",
        "slope",
    )
    for key in keys:
        av = a.get(key)
        bv = b.get(key)
        if av is None and bv is None:
            continue
        if av is None or bv is None:
            return False
        if abs(float(av) - float(bv)) > tol:
            return False
    return True


def extract_regression_rows(
    wb: xw.Book,
    source_file: str,
    label: FileLabel,
) -> list[dict[str, Any]]:
    sheet = get_sheet_or_none(wb, REGRESSION_SOURCE_SHEET)
    if sheet is None:
        print(f"  - skipped regression in {source_file}: sheet '{REGRESSION_SOURCE_SHEET}' not found")
        return []

    anchor = find_anchor(sheet, "max")
    if anchor is None:
        print(f"  - skipped regression in {source_file}: could not find 'max' anchor")
        return []

    anchor_row, anchor_col = anchor
    x_col, y_col, series = collect_numeric_series(sheet, anchor_row, anchor_col)

    if not series:
        print(f"  - skipped regression in {source_file}: no numeric history near anchor")
        return []

    rows: list[dict[str, Any]] = []
    n_max = min(MAX_QUARTERS_TO_TEST, len(series))

    temp_intercept_cell = sheet.cells(anchor_row + 2, anchor_col + 6)
    temp_slope_cell = sheet.cells(anchor_row + 3, anchor_col + 6)

    for n_quarters in range(1, n_max + 1):
        window = series[-n_quarters:]
        start_row = window[0][0]
        end_row = window[-1][0]

        y_rng = r1c1_range(start_row, end_row, y_col)
        x_rng = r1c1_range(start_row, end_row, x_col)
        temp_intercept_cell.formula2 = f"=INTERCEPT({y_rng},{x_rng})"
        temp_slope_cell.formula2 = f"=SLOPE({y_rng},{x_rng})"
        wb.app.calculate()

        intercept = to_float(temp_intercept_cell.value)
        slope = to_float(temp_slope_cell.value)
        if intercept is None or slope is None:
            continue

        x_values = [x for _, x, _ in window]
        y_values = [y for _, _, y in window]

        # Uses the most recent x as the base for TOT FCST w/o SA.
        x_latest = x_values[-1]
        forecast_total_without_sa = intercept + slope * x_latest

        residuals = [y - (intercept + slope * x) for x, y in zip(x_values, y_values)]
        residual_std = statistics.pstdev(residuals) if len(residuals) > 1 else 0.0

        forecast_max = forecast_total_without_sa + residual_std
        forecast_min = forecast_total_without_sa - residual_std
        range_width = forecast_max - forecast_min

        row = {
            "model": label.model,
            "ticker": label.ticker,
            "model_period": label.model_period,
            "model_date": label.model_date,
            "method": "regression",
            "parameter_name": "num_quarters_used",
            "parameter_value": n_quarters,
            "num_quarters_used": n_quarters,
            "forecast_value": forecast_total_without_sa,
            "actual_value": None,
            "forecast_max": forecast_max,
            "forecast_min": forecast_min,
            "range_width": range_width,
            "intercept": intercept,
            "slope": slope,
            "source_file": source_file,
        }

        if rows and regression_rows_match(rows[-1], row):
            continue
        rows.append(row)

    temp_intercept_cell.value = None
    temp_slope_cell.value = None
    return rows


def write_sheet(ws, headers: list[str], rows: Iterable[dict[str, Any]]) -> None:
    ws.append(headers)
    for header_cell in ws[1]:
        header_cell.font = Font(bold=True)

    for row in rows:
        ws.append([row.get(col) for col in headers])

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions

    for col_idx, header in enumerate(headers, start=1):
        col_letter = get_column_letter(col_idx)
        max_len = len(header)
        for cell in ws[col_letter]:
            if cell.value is None:
                continue
            max_len = max(max_len, len(str(cell.value)))
        ws.column_dimensions[col_letter].width = min(max_len + 2, 50)


def build_output_workbook(
    output_file_path: Path,
    empirical_rows: list[dict[str, Any]],
    regression_rows: list[dict[str, Any]],
) -> None:
    out_wb = Workbook()
    out_wb.remove(out_wb.active)

    empirical_ws = out_wb.create_sheet(EMPIRICAL_OUTPUT_SHEET)
    regression_ws = out_wb.create_sheet(REGRESSION_OUTPUT_SHEET)

    write_sheet(empirical_ws, EMPIRICAL_COLUMNS, empirical_rows)
    write_sheet(regression_ws, REGRESSION_COLUMNS, regression_rows)

    out_wb.save(output_file_path)


def iter_source_files(input_path: Path, output_name_pattern: re.Pattern[str]) -> Iterable[Path]:
    for path in sorted(input_path.iterdir()):
        if not path.is_file():
            continue
        if path.name.startswith("~"):
            print(f"Skipped {path.name}: temporary file")
            continue
        if path.suffix.lower() != ".xlsx":
            print(f"Skipped {path.name}: not an .xlsx file")
            continue
        if output_name_pattern.match(path.name):
            print(f"Skipped {path.name}: existing output file pattern")
            continue
        yield path


def configure_app(app: xw.App) -> None:
    app.visible = False
    app.display_alerts = False
    app.screen_updating = False
    try:
        app.api.EnableEvents = False
    except Exception:
        pass


def main() -> None:
    input_path = Path(input_dir)
    output_path = Path(output_dir)

    if not input_path.exists() or not input_path.is_dir():
        raise FileNotFoundError(f"input_dir does not exist or is not a directory: {input_path}")

    output_file_path = resolve_output_path(input_path, output_path)
    output_name_pattern = re.compile(
        OUTPUT_STEM_PATTERN_TEMPLATE.format(name=re.escape(input_path.name)),
        re.IGNORECASE,
    )

    empirical_rows: list[dict[str, Any]] = []
    regression_rows: list[dict[str, Any]] = []
    files_processed = 0

    app = xw.App(visible=False, add_book=False)
    configure_app(app)

    try:
        for file_path in iter_source_files(input_path, output_name_pattern):
            print(f"Processing {file_path.name}")
            files_processed += 1
            source_wb: xw.Book | None = None
            try:
                source_wb = app.books.open(str(file_path), update_links=False)
                label = parse_file_label(file_path.name)

                empirical_rows.extend(extract_empirical_rows(source_wb, file_path.name, label))
                regression_rows.extend(extract_regression_rows(source_wb, file_path.name, label))
            except Exception as exc:
                print(f"  - failed {file_path.name}: {exc}")
                traceback.print_exc()
            finally:
                if source_wb is not None:
                    safe_close_source_workbook(source_wb)
    finally:
        safe_quit_app(app)

    build_output_workbook(output_file_path, empirical_rows, regression_rows)

    print(f"Output path: {output_file_path}")
    print(f"Files processed: {files_processed}")
    print(f"Empirical rows: {len(empirical_rows)}")
    print(f"Regression rows: {len(regression_rows)}")


if __name__ == "__main__":
    main()
