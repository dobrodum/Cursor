#!/usr/bin/env python3
"""Extract empirical/regression model candidates from a folder of .xlsx files.

This script opens each source workbook exactly once (hidden Excel app), extracts
rows from both "Empirical Model" and "Regression Model" sheets, and writes one
output workbook with:
  - empirical_candidates
  - regression_candidates
"""

from __future__ import annotations

from dataclasses import dataclass
from datetime import date
from pathlib import Path
import re
from typing import Any, Iterable

import xlwings as xw
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter


# -----------------------------
# User-configurable paths
# -----------------------------
input_dir = Path("/path/to/input")
output_dir = Path("/path/to/output")


N_QUARTERS = 10

EMPIRICAL_HEADERS = [
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

REGRESSION_HEADERS = [
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

MONTH_MAP = {
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

PERIOD_DAY = {
    "early": 5,
    "mid": 15,
    "late": 25,
}

TEXT_NORMALIZER = re.compile(r"[^a-z0-9]+")
PERIOD_RE = re.compile(
    r"(?P<phase>Early|Mid|Late)(?P<month>[A-Za-z]{3,9})(?P<year>\d{4})",
    re.IGNORECASE,
)


@dataclass
class FileMeta:
    model: str
    ticker: str
    model_period: str
    model_date: str


@dataclass
class Anchor:
    row: int
    col: int


@dataclass
class SheetCache:
    first_row: int
    first_col: int
    values: list[list[Any]]

    @property
    def last_row(self) -> int:
        return self.first_row + len(self.values) - 1

    @property
    def last_col(self) -> int:
        if not self.values:
            return self.first_col
        return self.first_col + len(self.values[0]) - 1


def normalize_text(value: Any) -> str:
    if value is None:
        return ""
    text = str(value).strip().lower()
    return TEXT_NORMALIZER.sub("_", text).strip("_")


def to_2d(values: Any) -> list[list[Any]]:
    if values is None:
        return []
    if isinstance(values, (list, tuple)):
        if not values:
            return []
        first = values[0]
        if isinstance(first, (list, tuple)):
            return [list(row) for row in values]
        return [list(values)]
    return [[values]]


def build_sheet_cache(sheet: xw.Sheet) -> SheetCache:
    used = sheet.used_range
    values = to_2d(used.value)
    return SheetCache(first_row=used.row, first_col=used.column, values=values)


def cache_get(cache: SheetCache, row: int, col: int) -> Any:
    r_idx = row - cache.first_row
    c_idx = col - cache.first_col
    if r_idx < 0 or c_idx < 0:
        return None
    if r_idx >= len(cache.values):
        return None
    row_values = cache.values[r_idx]
    if c_idx >= len(row_values):
        return None
    return row_values[c_idx]


def is_number(value: Any) -> bool:
    return isinstance(value, (int, float)) and not isinstance(value, bool)


def as_float(value: Any) -> float | None:
    if is_number(value):
        return float(value)
    if isinstance(value, str):
        text = value.strip().replace(",", "")
        if not text:
            return None
        is_pct = text.endswith("%")
        if is_pct:
            text = text[:-1]
        try:
            parsed = float(text)
        except ValueError:
            return None
        return parsed / 100.0 if is_pct else parsed
    return None


def safe_subtract(a: Any, b: Any) -> float | None:
    a_f = as_float(a)
    b_f = as_float(b)
    if a_f is None or b_f is None:
        return None
    return a_f - b_f


def parse_file_meta(path: Path) -> FileMeta:
    stem = path.stem
    parts = [p.strip() for p in stem.split(" - ")]

    ticker = "UNKNOWN"
    if len(parts) >= 2:
        ticker_candidate = re.sub(r"[^A-Za-z0-9]", "", parts[1]).upper()
        if ticker_candidate:
            ticker = ticker_candidate

    period_match = PERIOD_RE.search(parts[2] if len(parts) >= 3 else stem)
    if period_match is None:
        period_match = PERIOD_RE.search(stem)

    model_period = "UNKNOWN_PERIOD"
    model_date = ""
    if period_match:
        phase = period_match.group("phase").title()
        month_key = period_match.group("month")[:3].lower()
        year = int(period_match.group("year"))
        month_num = MONTH_MAP.get(month_key)
        if month_num:
            model_period = f"{phase}{month_key.title()}_{year}"
            day = PERIOD_DAY[phase.lower()]
            model_date = date(year, month_num, day).isoformat()
        else:
            model_period = f"{phase}{period_match.group('month')}_{year}"

    model = f"{ticker}_{model_period}"
    return FileMeta(model=model, ticker=ticker, model_period=model_period, model_date=model_date)


def next_output_path(in_dir: Path, out_dir: Path) -> Path:
    base_name = f"{in_dir.name}_PARAM.xlsx"
    candidate = out_dir / base_name
    if not candidate.exists():
        return candidate

    idx = 1
    while True:
        candidate = out_dir / f"{in_dir.name}_PARAM.{idx}.xlsx"
        if not candidate.exists():
            return candidate
        idx += 1


def find_anchor(cache: SheetCache, token: str = "max") -> Anchor | None:
    norm_token = normalize_text(token)
    matches: list[tuple[int, int, int]] = []

    for r_idx, row_values in enumerate(cache.values):
        for c_idx, value in enumerate(row_values):
            if normalize_text(value) != norm_token:
                continue
            row = cache.first_row + r_idx
            col = cache.first_col + c_idx
            score = 0
            if normalize_text(cache_get(cache, row, col + 1)) == "min":
                score += 2
            if normalize_text(cache_get(cache, row, col - 1)) in {"tot_fcst_wo_sa", "tot_fcst_w_o_sa"}:
                score += 1
            matches.append((score, row, col))

    if not matches:
        return None

    matches.sort(key=lambda t: (-t[0], t[1], t[2]))
    _, row, col = matches[0]
    return Anchor(row=row, col=col)


def build_header_entries(
    cache: SheetCache,
    header_row: int,
    start_col: int,
    end_col: int,
) -> list[tuple[int, str]]:
    entries: list[tuple[int, str]] = []
    for col in range(start_col, end_col + 1):
        norm = normalize_text(cache_get(cache, header_row, col))
        if norm:
            entries.append((col, norm))
    return entries


def find_col(
    entries: list[tuple[int, str]],
    includes: Iterable[str],
    excludes: Iterable[str] = (),
) -> int | None:
    include_list = list(includes)
    exclude_list = list(excludes)
    for col, norm in entries:
        if all(token in norm for token in include_list) and not any(
            token in norm for token in exclude_list
        ):
            return col
    return None


def read_block(sheet: xw.Sheet, row1: int, row2: int, col1: int, col2: int) -> list[list[Any]]:
    values = sheet.range((row1, col1), (row2, col2)).value
    return to_2d(values)


def block_get(
    block: list[list[Any]],
    row_offset: int,
    col_abs: int | None,
    block_col_start: int,
) -> Any:
    if col_abs is None:
        return None
    c_idx = col_abs - block_col_start
    if row_offset < 0 or row_offset >= len(block):
        return None
    if c_idx < 0 or c_idx >= len(block[row_offset]):
        return None
    return block[row_offset][c_idx]


def find_last_numeric_row(cache: SheetCache, col: int, upper_bound_row: int) -> int | None:
    for row in range(upper_bound_row, cache.first_row - 1, -1):
        if is_number(cache_get(cache, row, col)):
            return row
    return None


def find_last_numeric_pair_row(
    cache: SheetCache,
    x_col: int,
    y_col: int,
    upper_bound_row: int,
) -> int | None:
    for row in range(upper_bound_row, cache.first_row - 1, -1):
        x_val = cache_get(cache, row, x_col)
        y_val = cache_get(cache, row, y_col)
        if is_number(x_val) and is_number(y_val):
            return row
    return None


def signature_value(value: Any) -> Any:
    val = as_float(value)
    if val is None:
        return value
    return round(val, 10)


def safe_close_workbook(workbook: xw.Book) -> None:
    try:
        workbook.close(save=False)
        return
    except TypeError:
        pass
    except Exception:
        pass

    try:
        workbook.close(False)
        return
    except Exception:
        pass

    try:
        workbook.api.Close(SaveChanges=False)
    except Exception:
        workbook.api.Close(False)


def extract_empirical_rows(sheet: xw.Sheet, meta: FileMeta, source_file: str) -> list[list[Any]]:
    cache = build_sheet_cache(sheet)
    anchor = find_anchor(cache, "max")
    if anchor is None:
        print(f"  skipped empirical extraction: 'max' anchor not found in {sheet.name}")
        return []

    header_entries = build_header_entries(
        cache=cache,
        header_row=anchor.row,
        start_col=max(1, anchor.col - 30),
        end_col=anchor.col + 8,
    )

    forecast_max_col = anchor.col
    forecast_min_col = find_col(header_entries, ["min"]) or (anchor.col + 1)
    forecast_value_col = (
        find_col(header_entries, ["estimated", "total", "sold"])
        or find_col(header_entries, ["tot", "fcst"])
        or find_col(header_entries, ["forecast"], ["max", "min"])
        or (anchor.col - 1)
    )
    num_quarters_col = (
        find_col(header_entries, ["num", "quarters"])
        or find_col(header_entries, ["quarters", "used"])
        or (anchor.col - 6)
    )
    last_quarter_col = find_col(header_entries, ["last", "quarter"]) or (anchor.col - 5)
    reported_sales_col = (
        find_col(header_entries, ["reported", "sales"])
        or find_col(header_entries, ["actual", "sales"])
        or (anchor.col - 2)
    )
    quarterly_sales_col = find_col(header_entries, ["quarterly", "sales"]) or (anchor.col - 3)
    growth_rate_col = find_col(header_entries, ["growth"])
    sales_captured_col = find_col(header_entries, ["captured"])
    avg_penetration_col = find_col(header_entries, ["avg", "penetration"])
    raw_penetration_col = (
        find_col(header_entries, ["penetration"], ["avg"])
        or find_col(header_entries, ["penetration"])
        or (anchor.col - 8)
    )

    cols = [
        forecast_max_col,
        forecast_min_col,
        forecast_value_col,
        num_quarters_col,
        last_quarter_col,
        reported_sales_col,
        quarterly_sales_col,
    ]
    if growth_rate_col is not None:
        cols.append(growth_rate_col)
    if sales_captured_col is not None:
        cols.append(sales_captured_col)
    if avg_penetration_col is not None:
        cols.append(avg_penetration_col)

    row_start = anchor.row + 1
    row_end = anchor.row + N_QUARTERS
    col_start = min(cols)
    col_end = max(cols)
    data_block = read_block(sheet, row_start, row_end, col_start, col_end)

    helper_col = col_end + 3
    helper_row = anchor.row
    helper_avg_cell = sheet.range((helper_row, helper_col))
    penetration_last_row = find_last_numeric_row(cache, raw_penetration_col, anchor.row - 1)

    rows: list[list[Any]] = []
    for i in range(N_QUARTERS):
        num_quarters_used = block_get(data_block, i, num_quarters_col, col_start)
        if not is_number(num_quarters_used):
            num_quarters_used = i + 1

        avg_penetration_pct = block_get(data_block, i, avg_penetration_col, col_start)
        if penetration_last_row is not None:
            start_row = penetration_last_row - int(i + 1) + 1
            if start_row >= cache.first_row:
                helper_avg_cell.formula2 = (
                    f"=AVERAGE(R{start_row}C{raw_penetration_col}:"
                    f"R{penetration_last_row}C{raw_penetration_col})"
                )
                sheet.book.app.calculate()
                avg_from_formula = helper_avg_cell.value
                if avg_from_formula is not None:
                    avg_penetration_pct = avg_from_formula

        forecast_value = block_get(data_block, i, forecast_value_col, col_start)
        reported_sales = block_get(data_block, i, reported_sales_col, col_start)
        quarterly_sales = block_get(data_block, i, quarterly_sales_col, col_start)

        if forecast_value is None:
            avg_val = as_float(avg_penetration_pct)
            sales_val = as_float(quarterly_sales)
            if avg_val is not None and sales_val is not None:
                avg_factor = avg_val / 100.0 if avg_val > 1 else avg_val
                forecast_value = sales_val * avg_factor

        forecast_max = block_get(data_block, i, forecast_max_col, col_start)
        forecast_min = block_get(data_block, i, forecast_min_col, col_start)
        growth_rate_pct = block_get(data_block, i, growth_rate_col, col_start)
        sales_captured_in_db_pct = block_get(data_block, i, sales_captured_col, col_start)
        last_quarter_used = block_get(data_block, i, last_quarter_col, col_start)

        if (
            forecast_value is None
            and forecast_max is None
            and forecast_min is None
            and avg_penetration_pct is None
        ):
            continue

        rows.append(
            [
                meta.model,
                meta.ticker,
                meta.model_period,
                meta.model_date,
                "empirical",
                "avg_penetration_pct",
                avg_penetration_pct,
                num_quarters_used,
                last_quarter_used,
                forecast_value,
                reported_sales,
                forecast_max,
                forecast_min,
                safe_subtract(forecast_max, forecast_min),
                avg_penetration_pct,
                quarterly_sales,
                reported_sales,
                growth_rate_pct,
                sales_captured_in_db_pct,
                source_file,
            ]
        )

    helper_avg_cell.value = None
    return rows


def extract_regression_rows(sheet: xw.Sheet, meta: FileMeta, source_file: str) -> list[list[Any]]:
    cache = build_sheet_cache(sheet)
    anchor = find_anchor(cache, "max")
    if anchor is None:
        print(f"  skipped regression extraction: 'max' anchor not found in {sheet.name}")
        return []

    header_entries = build_header_entries(
        cache=cache,
        header_row=anchor.row,
        start_col=max(1, anchor.col - 30),
        end_col=anchor.col + 8,
    )

    forecast_max_col = anchor.col
    forecast_min_col = find_col(header_entries, ["min"]) or (anchor.col + 1)
    forecast_value_col = (
        find_col(header_entries, ["tot", "fcst", "wo", "sa"])
        or find_col(header_entries, ["tot", "fcst"])
        or find_col(header_entries, ["forecast"], ["max", "min"])
        or (anchor.col - 1)
    )
    num_quarters_col = (
        find_col(header_entries, ["num", "quarters"])
        or find_col(header_entries, ["quarters", "used"])
        or (anchor.col - 6)
    )
    actual_value_col = find_col(header_entries, ["actual"])

    x_col = anchor.col - 11
    y_col = anchor.col - 7
    pair_last_row = find_last_numeric_pair_row(cache, x_col=x_col, y_col=y_col, upper_bound_row=anchor.row - 1)

    cols = [forecast_max_col, forecast_min_col, forecast_value_col, num_quarters_col]
    if actual_value_col is not None:
        cols.append(actual_value_col)
    row_start = anchor.row + 1
    row_end = anchor.row + N_QUARTERS
    col_start = min(cols)
    col_end = max(cols)
    data_block = read_block(sheet, row_start, row_end, col_start, col_end)

    helper_col = col_end + 3
    helper_row = anchor.row
    intercept_cell = sheet.range((helper_row, helper_col))
    slope_cell = sheet.range((helper_row, helper_col + 1))

    rows: list[list[Any]] = []
    prev_signature: tuple[Any, ...] | None = None

    for i in range(N_QUARTERS):
        quarter_count = i + 1
        num_quarters_used = block_get(data_block, i, num_quarters_col, col_start)
        if not is_number(num_quarters_used):
            num_quarters_used = quarter_count

        intercept = None
        slope = None
        if pair_last_row is not None:
            start_row = pair_last_row - quarter_count + 1
            if start_row >= cache.first_row:
                intercept_cell.formula2 = (
                    f"=INTERCEPT(R{start_row}C{y_col}:R{pair_last_row}C{y_col},"
                    f"R{start_row}C{x_col}:R{pair_last_row}C{x_col})"
                )
                slope_cell.formula2 = (
                    f"=SLOPE(R{start_row}C{y_col}:R{pair_last_row}C{y_col},"
                    f"R{start_row}C{x_col}:R{pair_last_row}C{x_col})"
                )
                sheet.book.app.calculate()
                intercept = intercept_cell.value
                slope = slope_cell.value

        forecast_value = block_get(data_block, i, forecast_value_col, col_start)
        forecast_max = block_get(data_block, i, forecast_max_col, col_start)
        forecast_min = block_get(data_block, i, forecast_min_col, col_start)
        actual_value = block_get(data_block, i, actual_value_col, col_start)

        if forecast_value is None and forecast_max is None and forecast_min is None:
            continue

        sig = (
            signature_value(num_quarters_used),
            signature_value(forecast_value),
            signature_value(forecast_max),
            signature_value(forecast_min),
            signature_value(intercept),
            signature_value(slope),
        )
        if prev_signature is not None and sig == prev_signature:
            continue
        prev_signature = sig

        rows.append(
            [
                meta.model,
                meta.ticker,
                meta.model_period,
                meta.model_date,
                "regression",
                "num_quarters_used",
                num_quarters_used,
                num_quarters_used,
                forecast_value,
                actual_value,
                forecast_max,
                forecast_min,
                safe_subtract(forecast_max, forecast_min),
                intercept,
                slope,
                source_file,
            ]
        )

    intercept_cell.value = None
    slope_cell.value = None
    return rows


def write_output_workbook(
    output_path: Path,
    empirical_rows: list[list[Any]],
    regression_rows: list[list[Any]],
) -> None:
    wb = Workbook()
    ws_emp = wb.active
    ws_emp.title = "empirical_candidates"
    ws_reg = wb.create_sheet("regression_candidates")

    write_sheet(ws_emp, EMPIRICAL_HEADERS, empirical_rows)
    write_sheet(ws_reg, REGRESSION_HEADERS, regression_rows)

    wb.save(output_path)


def write_sheet(worksheet, headers: list[str], rows: list[list[Any]]) -> None:
    worksheet.append(headers)
    for row in rows:
        worksheet.append(row)

    header_font = Font(bold=True)
    for idx in range(1, len(headers) + 1):
        worksheet.cell(row=1, column=idx).font = header_font

    worksheet.freeze_panes = "A2"
    worksheet.auto_filter.ref = worksheet.dimensions

    for col_idx, header in enumerate(headers, start=1):
        max_len = len(header)
        for row in rows:
            value = row[col_idx - 1]
            as_text = "" if value is None else str(value)
            if len(as_text) > max_len:
                max_len = len(as_text)
        worksheet.column_dimensions[get_column_letter(col_idx)].width = min(max_len + 2, 48)


def collect_source_files(in_dir: Path) -> list[Path]:
    files: list[Path] = []
    for path in sorted(in_dir.iterdir()):
        if not path.is_file():
            print(f"skipped {path.name}: not a file")
            continue
        if path.name.startswith("~"):
            print(f"skipped {path.name}: temporary file")
            continue
        if path.suffix.lower() != ".xlsx":
            print(f"skipped {path.name}: not an .xlsx file")
            continue
        files.append(path)
    return files


def main() -> None:
    in_dir = Path(input_dir)
    out_dir = Path(output_dir)

    if not in_dir.exists() or not in_dir.is_dir():
        raise FileNotFoundError(f"input_dir does not exist or is not a directory: {in_dir}")

    out_dir.mkdir(parents=True, exist_ok=True)
    source_files = collect_source_files(in_dir)
    output_path = next_output_path(in_dir, out_dir)

    empirical_rows: list[list[Any]] = []
    regression_rows: list[list[Any]] = []
    files_processed = 0

    app = xw.App(visible=False, add_book=False)
    app.display_alerts = False
    app.screen_updating = False

    try:
        for file_path in source_files:
            print(f"processing {file_path.name}")
            workbook = None
            try:
                workbook = app.books.open(str(file_path), update_links=False)

                meta = parse_file_meta(file_path)

                empirical_sheet = None
                regression_sheet = None
                for sheet in workbook.sheets:
                    norm_name = normalize_text(sheet.name)
                    if norm_name == "empirical_model":
                        empirical_sheet = sheet
                    elif norm_name == "regression_model":
                        regression_sheet = sheet

                if empirical_sheet is None:
                    print(f"  skipped empirical for {file_path.name}: missing 'Empirical Model' sheet")
                else:
                    empirical_rows.extend(
                        extract_empirical_rows(empirical_sheet, meta=meta, source_file=file_path.name)
                    )

                if regression_sheet is None:
                    print(f"  skipped regression for {file_path.name}: missing 'Regression Model' sheet")
                else:
                    regression_rows.extend(
                        extract_regression_rows(regression_sheet, meta=meta, source_file=file_path.name)
                    )

                files_processed += 1
            except Exception as exc:
                print(f"skipped {file_path.name}: {exc}")
            finally:
                if workbook is not None:
                    safe_close_workbook(workbook)
    finally:
        app.quit()

    write_output_workbook(output_path, empirical_rows=empirical_rows, regression_rows=regression_rows)

    print(f"output_path: {output_path}")
    print(f"files_processed: {files_processed}")
    print(f"empirical_rows: {len(empirical_rows)}")
    print(f"regression_rows: {len(regression_rows)}")


if __name__ == "__main__":
    main()
