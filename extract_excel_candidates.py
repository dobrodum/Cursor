from __future__ import annotations

import datetime as dt
import re
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Iterable

import xlwings as xw

# User-configurable paths.
input_dir = Path("./input")
output_dir = Path("./output")


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


DAY_BY_PERIOD = {"early": 5, "mid": 15, "late": 25}


@dataclass
class FileMeta:
    model: str
    ticker: str
    model_period: str
    model_date: str


@dataclass
class SheetCache:
    first_row: int
    first_col: int
    grid: list[list[Any]]
    labels: dict[str, tuple[int, int]]

    @property
    def last_row(self) -> int:
        if not self.grid:
            return self.first_row
        return self.first_row + len(self.grid) - 1

    @property
    def last_col(self) -> int:
        if not self.grid:
            return self.first_col
        return self.first_col + len(self.grid[0]) - 1


def normalize_label(value: str) -> str:
    return re.sub(r"\s+", " ", value.strip().lower())


def to_grid(values: Any) -> list[list[Any]]:
    if values is None:
        return []
    if not isinstance(values, list):
        return [[values]]
    if not values:
        return []
    if isinstance(values[0], list):
        rows = [list(row) if isinstance(row, list) else [row] for row in values]
    else:
        rows = [list(values)]

    width = max(len(row) for row in rows)
    for row in rows:
        if len(row) < width:
            row.extend([None] * (width - len(row)))
    return rows


def build_sheet_cache(sheet: xw.Sheet) -> SheetCache:
    used = sheet.used_range
    first_row = used.row
    first_col = used.column
    grid = to_grid(used.value)
    labels: dict[str, tuple[int, int]] = {}

    for r_idx, row in enumerate(grid):
        for c_idx, value in enumerate(row):
            if isinstance(value, str):
                key = normalize_label(value)
                if key and key not in labels:
                    labels[key] = (first_row + r_idx, first_col + c_idx)

    return SheetCache(first_row=first_row, first_col=first_col, grid=grid, labels=labels)


def get_value(sheet: xw.Sheet, cache: SheetCache, row: int, col: int) -> Any:
    r_idx = row - cache.first_row
    c_idx = col - cache.first_col
    if 0 <= r_idx < len(cache.grid) and 0 <= c_idx < len(cache.grid[0]):
        return cache.grid[r_idx][c_idx]
    return sheet.range((row, col)).value


def to_number(value: Any) -> float | None:
    if value is None:
        return None
    if isinstance(value, bool):
        return float(value)
    if isinstance(value, (int, float)):
        return float(value)
    if isinstance(value, str):
        text = value.strip().replace(",", "")
        if not text:
            return None
        if text.endswith("%"):
            text = text[:-1].strip()
            try:
                return float(text) / 100.0
            except ValueError:
                return None
        try:
            return float(text)
        except ValueError:
            return None
    return None


def to_scalar_list(values: Any, expected_len: int) -> list[Any]:
    if expected_len == 0:
        return []
    if isinstance(values, list):
        if values and isinstance(values[0], list):
            flat: list[Any] = []
            for row in values:
                if isinstance(row, list):
                    flat.append(row[0] if row else None)
                else:
                    flat.append(row)
            return flat
        return values
    return [values]


def parse_file_metadata(file_path: Path) -> FileMeta | None:
    stem = file_path.stem
    parts = [part.strip() for part in stem.split("-")]
    if len(parts) < 3:
        return None

    ticker = parts[-2].upper()
    period_token = re.split(r"[_\s]", parts[-1], maxsplit=1)[0]
    match = re.match(r"^(Early|Mid|Late)([A-Za-z]{3})(\d{4})$", period_token, re.IGNORECASE)
    if not match:
        return None

    period_name = match.group(1).capitalize()
    month_abbr = match.group(2).title()
    year_text = match.group(3)

    day = DAY_BY_PERIOD[period_name.lower()]
    try:
        month = dt.datetime.strptime(month_abbr, "%b").month
    except ValueError:
        return None

    model_period = f"{period_name}{month_abbr}_{year_text}"
    model_date = dt.date(int(year_text), month, day).isoformat()
    model = f"{ticker}_{model_period}"
    return FileMeta(model=model, ticker=ticker, model_period=model_period, model_date=model_date)


def find_anchor(cache: SheetCache, anchor_text: str = "max") -> tuple[int, int] | None:
    anchor_key = normalize_label(anchor_text)
    if anchor_key in cache.labels:
        return cache.labels[anchor_key]

    for label, position in cache.labels.items():
        if label.startswith(anchor_key):
            return position
    return None


def find_label_value_right(
    sheet: xw.Sheet, cache: SheetCache, labels: Iterable[str]
) -> float | None:
    normalized = [normalize_label(label) for label in labels]

    for label in normalized:
        if label in cache.labels:
            row, col = cache.labels[label]
            number = to_number(get_value(sheet, cache, row, col + 1))
            if number is not None:
                return number

    for key, (row, col) in cache.labels.items():
        if any(label in key for label in normalized):
            number = to_number(get_value(sheet, cache, row, col + 1))
            if number is not None:
                return number

    return None


def close_workbook_safely(workbook: xw.Book) -> None:
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
    except Exception as exc:
        print(f"[WARN] Could not close workbook cleanly: {exc}")


def resolve_output_path(input_folder: Path, output_folder: Path) -> Path:
    base_name = f"{input_folder.name}_PARAM"
    candidate = output_folder / f"{base_name}.xlsx"
    if not candidate.exists():
        return candidate

    suffix = 1
    while True:
        candidate = output_folder / f"{base_name}.{suffix}.xlsx"
        if not candidate.exists():
            return candidate
        suffix += 1


def build_empirical_rows(
    workbook: xw.Book, file_meta: FileMeta, source_file: str
) -> list[list[Any]]:
    try:
        sheet = workbook.sheets["Empirical Model"]
    except Exception:
        print(f"[SKIP] {source_file}: missing sheet 'Empirical Model'")
        return []

    cache = build_sheet_cache(sheet)
    anchor = find_anchor(cache, "max")
    if not anchor:
        print(f"[SKIP] {source_file}: could not find 'max' anchor in 'Empirical Model'")
        return []

    anchor_row, anchor_col = anchor
    quarter_label_col = anchor_col - 12
    quarterly_sales_col = anchor_col - 11
    reported_sales_col = anchor_col - 7
    growth_rate_col = anchor_col - 6
    captured_pct_col = anchor_col - 5
    n_quarters = 10

    history_rows: list[tuple[int, float, float]] = []
    for row in range(cache.first_row, anchor_row):
        quarterly_sales = to_number(get_value(sheet, cache, row, quarterly_sales_col))
        reported_sales = to_number(get_value(sheet, cache, row, reported_sales_col))
        if quarterly_sales is None or reported_sales is None or reported_sales == 0:
            continue
        history_rows.append((row, quarterly_sales, reported_sales))

    if not history_rows:
        print(f"[SKIP] {source_file}: no usable empirical history found")
        return []

    max_lookback = min(n_quarters, len(history_rows))
    lookbacks = list(range(1, max_lookback + 1))

    temp_row_start = max(cache.last_row + 2, anchor_row + 2)
    temp_col = max(cache.last_col + 2, anchor_col + 2)

    for idx, lookback in enumerate(lookbacks):
        start_row = history_rows[-lookback][0]
        end_row = history_rows[-1][0]
        formula = (
            f"=AVERAGE(R{start_row}C{quarterly_sales_col}:R{end_row}C{quarterly_sales_col}/"
            f"R{start_row}C{reported_sales_col}:R{end_row}C{reported_sales_col})"
        )
        sheet.range((temp_row_start + idx, temp_col)).formula2 = formula

    workbook.app.calculate()
    avg_values_raw = sheet.range(
        (temp_row_start, temp_col), (temp_row_start + len(lookbacks) - 1, temp_col)
    ).value
    avg_values = to_scalar_list(avg_values_raw, len(lookbacks))
    sheet.range(
        (temp_row_start, temp_col), (temp_row_start + len(lookbacks) - 1, temp_col)
    ).clear_contents()

    latest_row, latest_quarterly_sales, latest_reported_sales = history_rows[-1]
    forecast_max = to_number(get_value(sheet, cache, anchor_row, anchor_col + 1))
    forecast_min = find_label_value_right(sheet, cache, ["min", "forecast min", "minimum"])
    if forecast_min is None:
        maybe_min_label = normalize_label(str(get_value(sheet, cache, anchor_row + 1, anchor_col) or ""))
        if maybe_min_label == "min":
            forecast_min = to_number(get_value(sheet, cache, anchor_row + 1, anchor_col + 1))

    last_quarter_used = get_value(sheet, cache, latest_row, quarter_label_col)
    growth_rate_pct = to_number(get_value(sheet, cache, latest_row, growth_rate_col))
    sales_captured_in_db_pct = to_number(get_value(sheet, cache, latest_row, captured_pct_col))

    rows: list[list[Any]] = []
    for idx, lookback in enumerate(lookbacks):
        avg_penetration_pct = to_number(avg_values[idx] if idx < len(avg_values) else None)
        forecast_value = None
        if avg_penetration_pct not in (None, 0):
            forecast_value = latest_quarterly_sales / avg_penetration_pct

        range_width = None
        if forecast_max is not None and forecast_min is not None:
            range_width = forecast_max - forecast_min

        rows.append(
            [
                file_meta.model,
                file_meta.ticker,
                file_meta.model_period,
                file_meta.model_date,
                "empirical",
                "avg_penetration_pct",
                avg_penetration_pct,
                lookback,
                last_quarter_used,
                forecast_value,
                latest_reported_sales,
                forecast_max,
                forecast_min,
                range_width,
                avg_penetration_pct,
                latest_quarterly_sales,
                latest_reported_sales,
                growth_rate_pct,
                sales_captured_in_db_pct,
                source_file,
            ]
        )

    return rows


def build_regression_rows(
    workbook: xw.Book, file_meta: FileMeta, source_file: str
) -> list[list[Any]]:
    try:
        sheet = workbook.sheets["Regression Model"]
    except Exception:
        print(f"[SKIP] {source_file}: missing sheet 'Regression Model'")
        return []

    cache = build_sheet_cache(sheet)
    anchor = find_anchor(cache, "max")
    if not anchor:
        print(f"[SKIP] {source_file}: could not find 'max' anchor in 'Regression Model'")
        return []

    anchor_row, anchor_col = anchor
    y_col = anchor_col - 7
    x_col = anchor_col - 11
    n_quarters = 10

    history_rows: list[tuple[int, float, float]] = []
    for row in range(cache.first_row, anchor_row):
        x_val = to_number(get_value(sheet, cache, row, x_col))
        y_val = to_number(get_value(sheet, cache, row, y_col))
        if x_val is None or y_val is None:
            continue
        history_rows.append((row, x_val, y_val))

    if len(history_rows) < 2:
        print(f"[SKIP] {source_file}: not enough regression history rows")
        return []

    max_lookback = min(n_quarters, len(history_rows))
    lookbacks = list(range(2, max_lookback + 1))
    if not lookbacks:
        print(f"[SKIP] {source_file}: not enough rows for lookback regression")
        return []

    temp_row_start = max(cache.last_row + 2, anchor_row + 2)
    temp_intercept_col = max(cache.last_col + 2, anchor_col + 2)
    temp_slope_col = temp_intercept_col + 1

    for idx, lookback in enumerate(lookbacks):
        start_row = history_rows[-lookback][0]
        end_row = history_rows[-1][0]
        intercept_formula = (
            f"=INTERCEPT(R{start_row}C{y_col}:R{end_row}C{y_col},"
            f"R{start_row}C{x_col}:R{end_row}C{x_col})"
        )
        slope_formula = (
            f"=SLOPE(R{start_row}C{y_col}:R{end_row}C{y_col},"
            f"R{start_row}C{x_col}:R{end_row}C{x_col})"
        )
        sheet.range((temp_row_start + idx, temp_intercept_col)).formula2 = intercept_formula
        sheet.range((temp_row_start + idx, temp_slope_col)).formula2 = slope_formula

    workbook.app.calculate()
    intercepts_raw = sheet.range(
        (temp_row_start, temp_intercept_col),
        (temp_row_start + len(lookbacks) - 1, temp_intercept_col),
    ).value
    slopes_raw = sheet.range(
        (temp_row_start, temp_slope_col),
        (temp_row_start + len(lookbacks) - 1, temp_slope_col),
    ).value

    intercepts = to_scalar_list(intercepts_raw, len(lookbacks))
    slopes = to_scalar_list(slopes_raw, len(lookbacks))
    sheet.range(
        (temp_row_start, temp_intercept_col),
        (temp_row_start + len(lookbacks) - 1, temp_slope_col),
    ).clear_contents()

    forecast_input = to_number(get_value(sheet, cache, anchor_row, x_col))
    if forecast_input is None:
        forecast_input = history_rows[-1][1]

    forecast_max = to_number(get_value(sheet, cache, anchor_row, anchor_col + 1))
    forecast_min = find_label_value_right(sheet, cache, ["min", "forecast min", "minimum"])
    if forecast_min is None:
        maybe_min_label = normalize_label(str(get_value(sheet, cache, anchor_row + 1, anchor_col) or ""))
        if maybe_min_label == "min":
            forecast_min = to_number(get_value(sheet, cache, anchor_row + 1, anchor_col + 1))

    actual_value = find_label_value_right(
        sheet, cache, ["actual value", "actual", "reported sales", "actual sales"]
    )

    rows: list[list[Any]] = []
    previous_signature: tuple[Any, ...] | None = None
    for idx, lookback in enumerate(lookbacks):
        intercept = to_number(intercepts[idx] if idx < len(intercepts) else None)
        slope = to_number(slopes[idx] if idx < len(slopes) else None)
        if intercept is None or slope is None:
            continue

        forecast_total_without_sa = None
        if forecast_input is not None:
            forecast_total_without_sa = intercept + (slope * forecast_input)

        range_width = None
        if forecast_max is not None and forecast_min is not None:
            range_width = forecast_max - forecast_min

        signature = (
            lookback,
            round(intercept, 10),
            round(slope, 10),
            round(forecast_total_without_sa, 10) if forecast_total_without_sa is not None else None,
            round(forecast_max, 10) if forecast_max is not None else None,
            round(forecast_min, 10) if forecast_min is not None else None,
        )
        if signature == previous_signature:
            continue

        rows.append(
            [
                file_meta.model,
                file_meta.ticker,
                file_meta.model_period,
                file_meta.model_date,
                "regression",
                "num_quarters_used",
                lookback,
                lookback,
                forecast_total_without_sa,
                actual_value,
                forecast_max,
                forecast_min,
                range_width,
                intercept,
                slope,
                source_file,
            ]
        )
        previous_signature = signature

    return rows


def compute_widths(headers: list[str], rows: list[list[Any]], min_width: int = 12, max_width: int = 48) -> list[int]:
    widths: list[int] = []
    for col_idx, header in enumerate(headers):
        max_len = len(header)
        for row in rows:
            if col_idx >= len(row):
                continue
            value = row[col_idx]
            text = "" if value is None else str(value)
            max_len = max(max_len, len(text))
        widths.append(max(min_width, min(max_width, max_len + 2)))
    return widths


def write_sheet(sheet: xw.Sheet, headers: list[str], rows: list[list[Any]]) -> None:
    sheet.range((1, 1)).value = headers
    if rows:
        sheet.range((2, 1)).value = rows

    header_range = sheet.range((1, 1), (1, len(headers)))
    header_range.api.Font.Bold = True

    filter_last_row = max(2, len(rows) + 1)
    filter_range = sheet.range((1, 1), (filter_last_row, len(headers)))
    filter_range.api.AutoFilter()

    widths = compute_widths(headers, rows)
    for col_idx, width in enumerate(widths, start=1):
        sheet.range((1, col_idx), (1, col_idx)).column_width = width

    try:
        sheet.activate()
        window = sheet.book.app.api.ActiveWindow
        window.SplitRow = 1
        window.SplitColumn = 0
        window.FreezePanes = True
    except Exception:
        pass


def write_output_workbook(
    app: xw.App,
    output_path: Path,
    empirical_rows: list[list[Any]],
    regression_rows: list[list[Any]],
) -> None:
    out_wb = app.books.add()
    try:
        empirical_sheet = out_wb.sheets[0]
        empirical_sheet.name = "empirical_candidates"
        regression_sheet = out_wb.sheets.add("regression_candidates", after=empirical_sheet)

        write_sheet(empirical_sheet, EMPIRICAL_HEADERS, empirical_rows)
        write_sheet(regression_sheet, REGRESSION_HEADERS, regression_rows)

        out_wb.save(str(output_path))
    finally:
        close_workbook_safely(out_wb)


def main() -> None:
    if not input_dir.exists() or not input_dir.is_dir():
        print(f"[ERROR] input_dir does not exist or is not a directory: {input_dir}")
        return

    output_dir.mkdir(parents=True, exist_ok=True)
    output_path = resolve_output_path(input_dir, output_dir)

    files = sorted((path for path in input_dir.iterdir() if path.is_file()), key=lambda p: p.name.lower())

    app = xw.App(visible=False, add_book=False)
    app.display_alerts = False
    app.screen_updating = False
    try:
        app.calculation = "manual"
    except Exception:
        pass

    processed_files = 0
    empirical_rows: list[list[Any]] = []
    regression_rows: list[list[Any]] = []

    try:
        for file_path in files:
            if file_path.name.startswith("~"):
                print(f"[SKIP] {file_path.name}: temp file")
                continue
            if file_path.suffix.lower() != ".xlsx":
                print(f"[SKIP] {file_path.name}: not an .xlsx file")
                continue

            file_meta = parse_file_metadata(file_path)
            if file_meta is None:
                print(f"[SKIP] {file_path.name}: filename does not match expected model format")
                continue

            print(f"[PROCESS] {file_path.name}")
            workbook: xw.Book | None = None
            try:
                workbook = app.books.open(str(file_path), update_links=False)
                empirical_file_rows = build_empirical_rows(workbook, file_meta, file_path.name)
                regression_file_rows = build_regression_rows(workbook, file_meta, file_path.name)
                empirical_rows.extend(empirical_file_rows)
                regression_rows.extend(regression_file_rows)
                processed_files += 1
            except Exception as exc:
                print(f"[SKIP] {file_path.name}: processing error -> {exc}")
            finally:
                if workbook is not None:
                    close_workbook_safely(workbook)

        write_output_workbook(app, output_path, empirical_rows, regression_rows)
    finally:
        try:
            app.quit()
        except Exception:
            pass

    print(f"output_path: {output_path}")
    print(f"files_processed: {processed_files}")
    print(f"empirical_rows: {len(empirical_rows)}")
    print(f"regression_rows: {len(regression_rows)}")


if __name__ == "__main__":
    main()
