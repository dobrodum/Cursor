#!/usr/bin/env python3
"""
Extract empirical/regression model candidates from .xlsx workbooks.

This script opens each source workbook once, processes both model sheets while
the workbook is open, then closes it without saving.
"""

from __future__ import annotations

from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path
import re
from typing import Any, Iterable

import xlwings as xw
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter


# User-configurable paths
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


EMPIRICAL_OFFSETS = {
    "num_quarters_used": -10,
    "last_quarter_used": -9,
    "quarterly_sales": -8,
    "reported_sales": -7,
    "growth_rate_pct": -6,
    "sales_captured_in_db_pct": -5,
    "avg_penetration_pct": -4,
    "forecast_value": -1,
    "actual_value": -2,
    "forecast_max": 0,
    "forecast_min": 1,
}

REGRESSION_OFFSETS = {
    "num_quarters_used": -10,
    "forecast_value": -1,
    "forecast_max": 0,
    "forecast_min": 1,
    "actual_value": -2,
}


PHASE_DAY = {"early": 5, "mid": 15, "late": 25}
PERIOD_PATTERN = re.compile(
    r"(?i)\b(Early|Mid|Late)(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)(\d{4})\b"
)


@dataclass
class FileMeta:
    model: str
    ticker: str
    model_period: str
    model_date: str


@dataclass
class SheetCache:
    values: list[list[Any]]
    first_row: int
    first_col: int


def normalize_header(value: Any) -> str:
    if value is None:
        return ""
    return re.sub(r"[^a-z0-9]+", " ", str(value).lower()).strip()


def to_float(value: Any) -> float | None:
    if value is None or value == "":
        return None
    if isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        return float(value)
    text = str(value).strip()
    if not text:
        return None
    pct = "%" in text
    cleaned = text.replace(",", "").replace("%", "").strip()
    try:
        number = float(cleaned)
        if pct:
            number = number / 100.0
        return number
    except ValueError:
        return None


def to_int(value: Any) -> int | None:
    number = to_float(value)
    if number is None:
        return None
    return int(round(number))


def value_for_output(value: Any) -> Any:
    if isinstance(value, datetime):
        return value.date().isoformat()
    if isinstance(value, date):
        return value.isoformat()
    return value


def parse_file_meta(file_path: Path) -> FileMeta:
    stem = file_path.stem
    parts = [part.strip() for part in stem.split(" - ")]
    ticker = ""
    if len(parts) >= 2:
        ticker = re.sub(r"[^A-Za-z0-9_]+", "", parts[1]).upper()
    if not ticker:
        ticker = "UNKNOWN"

    period_match = PERIOD_PATTERN.search(stem)
    if period_match:
        phase_raw, month_raw, year_raw = period_match.groups()
        phase = phase_raw.title()
        month = month_raw.title()
        year = int(year_raw)
        month_num = datetime.strptime(month, "%b").month
        day = PHASE_DAY[phase.lower()]
        model_period = f"{phase}{month}_{year}"
        model_date = date(year, month_num, day).isoformat()
    else:
        model_period = "UNKNOWN"
        model_date = ""

    model = f"{ticker}_{model_period}" if model_period else ticker
    return FileMeta(model=model, ticker=ticker, model_period=model_period, model_date=model_date)


def output_path_for_run(in_dir: Path, out_dir: Path) -> Path:
    out_dir.mkdir(parents=True, exist_ok=True)
    base_name = f"{in_dir.resolve().name}_PARAM"
    candidate = out_dir / f"{base_name}.xlsx"
    if not candidate.exists():
        return candidate

    idx = 1
    while True:
        candidate = out_dir / f"{base_name}.{idx}.xlsx"
        if not candidate.exists():
            return candidate
        idx += 1


def iter_source_files(in_dir: Path) -> Iterable[Path]:
    if not in_dir.exists():
        print(f"Input directory does not exist: {in_dir}")
        return []
    files = sorted(in_dir.iterdir(), key=lambda p: p.name.lower())
    return files


def get_sheet(book: xw.Book, sheet_name: str) -> xw.Sheet | None:
    try:
        return book.sheets[sheet_name]
    except Exception:
        return None


def sheet_cache(sheet: xw.Sheet) -> SheetCache:
    used = sheet.used_range
    values = used.value
    if values is None:
        matrix: list[list[Any]] = []
    elif isinstance(values, list):
        if values and isinstance(values[0], list):
            matrix = values
        else:
            matrix = [values]
    else:
        matrix = [[values]]
    return SheetCache(values=matrix, first_row=used.row, first_col=used.column)


def find_max_anchor(cache: SheetCache) -> tuple[int, int] | None:
    candidates: list[tuple[int, int, bool]] = []
    for r_idx, row in enumerate(cache.values):
        for c_idx, value in enumerate(row):
            if isinstance(value, str) and value.strip().lower() == "max":
                has_min_right = False
                if c_idx + 1 < len(row):
                    right = row[c_idx + 1]
                    has_min_right = isinstance(right, str) and right.strip().lower() == "min"
                candidates.append((cache.first_row + r_idx, cache.first_col + c_idx, has_min_right))
    if not candidates:
        return None
    candidates.sort(key=lambda x: (not x[2], x[0], x[1]))
    return candidates[0][0], candidates[0][1]


def row_values_from_cache(cache: SheetCache, row_number: int) -> list[Any]:
    idx = row_number - cache.first_row
    if idx < 0 or idx >= len(cache.values):
        return []
    row = cache.values[idx]
    if isinstance(row, list):
        return row
    return [row]


def find_column(
    header_rows: list[list[Any]],
    first_col: int,
    token_sets: list[tuple[str, ...]],
    fallback: int,
) -> int:
    for row in header_rows:
        for offset, value in enumerate(row):
            normalized = normalize_header(value)
            if not normalized:
                continue
            for token_set in token_sets:
                if all(token in normalized for token in token_set):
                    return first_col + offset
    return fallback


def set_formula2(cell: xw.Range, formula_r1c1: str) -> None:
    try:
        cell.formula2 = formula_r1c1
    except Exception:
        cell.formula = formula_r1c1


def safe_close_no_save(book: xw.Book) -> None:
    close_attempts = [
        lambda: book.close(save=False),
        lambda: book.close(SaveChanges=False),  # type: ignore[arg-type]
        lambda: book.api.Close(SaveChanges=False),
        lambda: book.close(),
    ]
    last_error: Exception | None = None
    for closer in close_attempts:
        try:
            closer()
            return
        except Exception as exc:
            last_error = exc
    if last_error:
        raise last_error


def latest_numeric_row(sheet: xw.Sheet, col: int, start_row: int) -> int | None:
    row = start_row
    while row >= 1:
        if to_float(sheet.cells(row, col).value) is not None:
            return row
        row -= 1
    return None


def clamp_min(value: int, minimum: int) -> int:
    return value if value >= minimum else minimum


def extract_empirical_rows(book: xw.Book, meta: FileMeta, source_file: str) -> list[dict[str, Any]]:
    sheet = get_sheet(book, "Empirical Model")
    if sheet is None:
        print(f"Skipped Empirical Model in {source_file}: sheet not found")
        return []

    cache = sheet_cache(sheet)
    anchor = find_max_anchor(cache)
    if anchor is None:
        print(f"Skipped Empirical Model in {source_file}: max anchor not found")
        return []

    anchor_row, anchor_col = anchor
    header_rows = [
        row_values_from_cache(cache, anchor_row),
        row_values_from_cache(cache, anchor_row - 1),
    ]
    first_col = cache.first_col

    col_num_quarters = find_column(
        header_rows,
        first_col,
        token_sets=[("num", "quarter"), ("quarters", "used"), ("n", "qtrs")],
        fallback=anchor_col + EMPIRICAL_OFFSETS["num_quarters_used"],
    )
    col_last_quarter = find_column(
        header_rows,
        first_col,
        token_sets=[("last", "quarter"), ("quarter", "used")],
        fallback=anchor_col + EMPIRICAL_OFFSETS["last_quarter_used"],
    )
    col_quarterly_sales = find_column(
        header_rows,
        first_col,
        token_sets=[("quarterly", "sales"), ("quarter", "sales")],
        fallback=anchor_col + EMPIRICAL_OFFSETS["quarterly_sales"],
    )
    col_reported_sales = find_column(
        header_rows,
        first_col,
        token_sets=[("reported", "sales"), ("actual", "sales")],
        fallback=anchor_col + EMPIRICAL_OFFSETS["reported_sales"],
    )
    col_growth_rate = find_column(
        header_rows,
        first_col,
        token_sets=[("growth", "rate"), ("growth", "%")],
        fallback=anchor_col + EMPIRICAL_OFFSETS["growth_rate_pct"],
    )
    col_sales_captured = find_column(
        header_rows,
        first_col,
        token_sets=[("sales", "captured"), ("captured", "db")],
        fallback=anchor_col + EMPIRICAL_OFFSETS["sales_captured_in_db_pct"],
    )
    col_avg_pen = find_column(
        header_rows,
        first_col,
        token_sets=[("avg", "penetration"), ("penetration", "%")],
        fallback=anchor_col + EMPIRICAL_OFFSETS["avg_penetration_pct"],
    )
    col_forecast_value = find_column(
        header_rows,
        first_col,
        token_sets=[("estimated", "total", "sold"), ("forecast", "value")],
        fallback=anchor_col + EMPIRICAL_OFFSETS["forecast_value"],
    )
    col_actual_value = find_column(
        header_rows,
        first_col,
        token_sets=[("actual", "value"), ("reported", "sales")],
        fallback=anchor_col + EMPIRICAL_OFFSETS["actual_value"],
    )
    col_forecast_max = anchor_col + EMPIRICAL_OFFSETS["forecast_max"]
    col_forecast_min = anchor_col + EMPIRICAL_OFFSETS["forecast_min"]

    n_quarters = 10
    data_start_row = anchor_row + 1
    temp_avg_col = max(anchor_col + 20, 50)

    history_end = latest_numeric_row(sheet, col_avg_pen, anchor_row - 1)
    formula_rows: list[tuple[int, int]] = []
    if history_end is not None:
        for i in range(n_quarters):
            row = data_start_row + i
            requested_n = to_int(sheet.cells(row, col_num_quarters).value)
            effective_n = requested_n if requested_n and requested_n > 0 else (i + 1)
            start_row = clamp_min(history_end - effective_n + 1, 1)
            formula = f'=IFERROR(AVERAGE(R{start_row}C{col_avg_pen}:R{history_end}C{col_avg_pen}),"")'
            set_formula2(sheet.cells(row, temp_avg_col), formula)
            formula_rows.append((row, effective_n))
        if formula_rows:
            book.app.calculate()

    rows: list[dict[str, Any]] = []
    for i in range(n_quarters):
        row = data_start_row + i
        num_quarters_used = to_int(sheet.cells(row, col_num_quarters).value) or (i + 1)
        last_quarter_used = value_for_output(sheet.cells(row, col_last_quarter).value)
        quarterly_sales = to_float(sheet.cells(row, col_quarterly_sales).value)
        reported_sales = to_float(sheet.cells(row, col_reported_sales).value)
        growth_rate_pct = to_float(sheet.cells(row, col_growth_rate).value)
        sales_captured_pct = to_float(sheet.cells(row, col_sales_captured).value)

        avg_penetration_pct = to_float(sheet.cells(row, temp_avg_col).value)
        if avg_penetration_pct is None:
            avg_penetration_pct = to_float(sheet.cells(row, col_avg_pen).value)

        forecast_value = to_float(sheet.cells(row, col_forecast_value).value)
        actual_value = to_float(sheet.cells(row, col_actual_value).value)
        if actual_value is None:
            actual_value = reported_sales
        if forecast_value is None and reported_sales is not None and avg_penetration_pct:
            forecast_value = reported_sales / avg_penetration_pct

        forecast_max = to_float(sheet.cells(row, col_forecast_max).value)
        forecast_min = to_float(sheet.cells(row, col_forecast_min).value)
        range_width = (
            forecast_max - forecast_min
            if forecast_max is not None and forecast_min is not None
            else None
        )

        has_data = any(
            value is not None
            for value in [
                forecast_value,
                actual_value,
                forecast_max,
                forecast_min,
                avg_penetration_pct,
            ]
        )
        if not has_data:
            continue

        rows.append(
            {
                "model": meta.model,
                "ticker": meta.ticker,
                "model_period": meta.model_period,
                "model_date": meta.model_date,
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
                "sales_captured_in_db_pct": sales_captured_pct,
                "source_file": source_file,
            }
        )

    return rows


def extract_regression_rows(book: xw.Book, meta: FileMeta, source_file: str) -> list[dict[str, Any]]:
    sheet = get_sheet(book, "Regression Model")
    if sheet is None:
        print(f"Skipped Regression Model in {source_file}: sheet not found")
        return []

    cache = sheet_cache(sheet)
    anchor = find_max_anchor(cache)
    if anchor is None:
        print(f"Skipped Regression Model in {source_file}: max anchor not found")
        return []

    anchor_row, anchor_col = anchor
    header_rows = [
        row_values_from_cache(cache, anchor_row),
        row_values_from_cache(cache, anchor_row - 1),
    ]
    first_col = cache.first_col

    col_num_quarters = find_column(
        header_rows,
        first_col,
        token_sets=[("num", "quarter"), ("quarters", "used"), ("n", "qtrs")],
        fallback=anchor_col + REGRESSION_OFFSETS["num_quarters_used"],
    )
    col_forecast_value = find_column(
        header_rows,
        first_col,
        token_sets=[("tot", "fcst", "w", "o", "sa"), ("forecast", "without", "sa")],
        fallback=anchor_col + REGRESSION_OFFSETS["forecast_value"],
    )
    col_actual_value = find_column(
        header_rows,
        first_col,
        token_sets=[("actual", "value"), ("reported", "sales")],
        fallback=anchor_col + REGRESSION_OFFSETS["actual_value"],
    )
    col_forecast_max = anchor_col + REGRESSION_OFFSETS["forecast_max"]
    col_forecast_min = anchor_col + REGRESSION_OFFSETS["forecast_min"]

    # Explicit anchor-based offsets required by the spec.
    y_col = anchor_col - 7
    x_col = anchor_col - 11

    n_quarters = 10
    data_start_row = anchor_row + 1
    temp_intercept_col = max(anchor_col + 20, 50)
    temp_slope_col = temp_intercept_col + 1

    history_end = latest_numeric_row(sheet, y_col, anchor_row - 1)
    formula_rows: list[tuple[int, int]] = []
    if history_end is not None:
        for i in range(n_quarters):
            row = data_start_row + i
            requested_n = to_int(sheet.cells(row, col_num_quarters).value)
            effective_n = requested_n if requested_n and requested_n > 1 else max(2, i + 1)
            start_row = clamp_min(history_end - effective_n + 1, 1)
            intercept_formula = (
                f'=IFERROR(INTERCEPT(R{start_row}C{y_col}:R{history_end}C{y_col},'
                f'R{start_row}C{x_col}:R{history_end}C{x_col}),"")'
            )
            slope_formula = (
                f'=IFERROR(SLOPE(R{start_row}C{y_col}:R{history_end}C{y_col},'
                f'R{start_row}C{x_col}:R{history_end}C{x_col}),"")'
            )
            set_formula2(sheet.cells(row, temp_intercept_col), intercept_formula)
            set_formula2(sheet.cells(row, temp_slope_col), slope_formula)
            formula_rows.append((row, effective_n))
        if formula_rows:
            book.app.calculate()

    rows: list[dict[str, Any]] = []
    previous_signature: tuple[Any, ...] | None = None
    for i in range(n_quarters):
        row = data_start_row + i
        num_quarters_used = to_int(sheet.cells(row, col_num_quarters).value) or (i + 1)
        intercept = to_float(sheet.cells(row, temp_intercept_col).value)
        slope = to_float(sheet.cells(row, temp_slope_col).value)
        forecast_value = to_float(sheet.cells(row, col_forecast_value).value)
        forecast_max = to_float(sheet.cells(row, col_forecast_max).value)
        forecast_min = to_float(sheet.cells(row, col_forecast_min).value)
        actual_value = to_float(sheet.cells(row, col_actual_value).value)
        range_width = (
            forecast_max - forecast_min
            if forecast_max is not None and forecast_min is not None
            else None
        )

        has_data = any(
            value is not None
            for value in [intercept, slope, forecast_value, forecast_max, forecast_min]
        )
        if not has_data:
            continue

        signature = (
            num_quarters_used,
            round(intercept, 10) if intercept is not None else None,
            round(slope, 10) if slope is not None else None,
            round(forecast_value, 10) if forecast_value is not None else None,
            round(forecast_max, 10) if forecast_max is not None else None,
            round(forecast_min, 10) if forecast_min is not None else None,
        )
        if previous_signature == signature:
            continue
        previous_signature = signature

        rows.append(
            {
                "model": meta.model,
                "ticker": meta.ticker,
                "model_period": meta.model_period,
                "model_date": meta.model_date,
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


def write_sheet(ws: Any, columns: list[str], rows: list[dict[str, Any]]) -> None:
    ws.append(columns)
    for col_idx in range(1, len(columns) + 1):
        ws.cell(row=1, column=col_idx).font = Font(bold=True)

    for row_data in rows:
        ws.append([value_for_output(row_data.get(col)) for col in columns])

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions

    for col_idx, col_name in enumerate(columns, start=1):
        max_len = len(col_name)
        for row_data in rows:
            value = row_data.get(col_name)
            if value is None:
                continue
            max_len = max(max_len, len(str(value)))
        ws.column_dimensions[get_column_letter(col_idx)].width = min(max(max_len + 2, 12), 50)


def write_output(path: Path, empirical_rows: list[dict[str, Any]], regression_rows: list[dict[str, Any]]) -> None:
    workbook = Workbook()
    default_sheet = workbook.active
    workbook.remove(default_sheet)

    ws_empirical = workbook.create_sheet("empirical_candidates")
    write_sheet(ws_empirical, EMPIRICAL_COLUMNS, empirical_rows)

    ws_regression = workbook.create_sheet("regression_candidates")
    write_sheet(ws_regression, REGRESSION_COLUMNS, regression_rows)

    workbook.save(path)


def main() -> int:
    files_processed = 0
    empirical_all: list[dict[str, Any]] = []
    regression_all: list[dict[str, Any]] = []

    files = list(iter_source_files(input_dir))
    if not files and not input_dir.exists():
        return 1

    app = xw.App(visible=False, add_book=False)
    app.display_alerts = False
    app.screen_updating = False

    try:
        for file_path in files:
            if not file_path.is_file():
                print(f"Skipped {file_path.name}: not a file")
                continue
            if file_path.name.startswith("~"):
                print(f"Skipped {file_path.name}: temp file")
                continue
            if file_path.suffix.lower() != ".xlsx":
                print(f"Skipped {file_path.name}: not .xlsx")
                continue

            print(f"Processed file: {file_path.name}")
            wb: xw.Book | None = None
            try:
                wb = app.books.open(str(file_path), update_links=False)
                meta = parse_file_meta(file_path)
                empirical_rows = extract_empirical_rows(wb, meta, file_path.name)
                regression_rows = extract_regression_rows(wb, meta, file_path.name)
                empirical_all.extend(empirical_rows)
                regression_all.extend(regression_rows)
                files_processed += 1
            except Exception as exc:
                print(f"Skipped {file_path.name}: processing error ({exc})")
            finally:
                if wb is not None:
                    try:
                        safe_close_no_save(wb)
                    except Exception as close_exc:
                        print(f"Warning: failed to close {file_path.name} cleanly ({close_exc})")

        out_path = output_path_for_run(input_dir, output_dir)
        write_output(out_path, empirical_all, regression_all)
        print(f"Output path: {out_path}")
        print(f"Number of files processed: {files_processed}")
        print(f"Number of empirical rows: {len(empirical_all)}")
        print(f"Number of regression rows: {len(regression_all)}")
        return 0
    finally:
        try:
            app.quit()
        except Exception:
            pass


if __name__ == "__main__":
    raise SystemExit(main())
