from __future__ import annotations

import re
from dataclasses import dataclass
from datetime import date
from pathlib import Path
from typing import Any

import xlwings as xw
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

# Configure these two paths before running.
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

WINDOW_DAY = {"early": 5, "mid": 15, "late": 25}
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
PERIOD_RE = re.compile(
    r"(?P<window>early|mid|late)(?P<month>[A-Za-z]{3,9})(?P<year>\d{4})", re.IGNORECASE
)
OUTPUT_FILE_RE = re.compile(r"_PARAM(?:\.\d+)?\.xlsx$", re.IGNORECASE)


@dataclass(frozen=True)
class FileMeta:
    model: str
    ticker: str
    model_period: str
    model_date: str


def ensure_2d(values: Any) -> list[list[Any]]:
    if values is None:
        return []
    if not isinstance(values, list):
        return [[values]]
    if values and not isinstance(values[0], list):
        return [values]
    return values


def as_float(value: Any) -> float | None:
    if value is None or value == "":
        return None
    if isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        return float(value)
    if isinstance(value, str):
        cleaned = value.strip().replace(",", "")
        if not cleaned:
            return None
        if cleaned.endswith("%"):
            try:
                return float(cleaned[:-1]) / 100.0
            except ValueError:
                return None
        try:
            return float(cleaned)
        except ValueError:
            return None
    return None


def to_month_abbrev(month_number: int) -> str:
    return [
        "",
        "Jan",
        "Feb",
        "Mar",
        "Apr",
        "May",
        "Jun",
        "Jul",
        "Aug",
        "Sep",
        "Oct",
        "Nov",
        "Dec",
    ][month_number]


def parse_file_meta(file_name: str) -> FileMeta:
    stem = Path(file_name).stem
    parts = [part.strip() for part in stem.split(" - ")]
    ticker = parts[1] if len(parts) >= 2 else "UNKNOWN"

    period_token = ""
    if len(parts) >= 3:
        period_token = parts[2].split("_")[0].strip()

    match = PERIOD_RE.search(period_token)
    if match:
        window = match.group("window").lower()
        month_token = match.group("month")[:3].lower()
        year = int(match.group("year"))
        month = MONTH_MAP.get(month_token)
        if month is None:
            model_period = period_token or "unknown_period"
            model_date = ""
        else:
            window_title = window.capitalize()
            model_period = f"{window_title}{to_month_abbrev(month)}_{year}"
            model_date = date(year, month, WINDOW_DAY[window]).isoformat()
    else:
        model_period = period_token or "unknown_period"
        model_date = ""

    model = f"{ticker}_{model_period}"
    return FileMeta(model=model, ticker=ticker, model_period=model_period, model_date=model_date)


def next_output_path(input_folder: Path, destination: Path) -> Path:
    base = f"{input_folder.name}_PARAM.xlsx"
    candidate = destination / base
    if not candidate.exists():
        return candidate

    index = 1
    while True:
        candidate = destination / f"{input_folder.name}_PARAM.{index}.xlsx"
        if not candidate.exists():
            return candidate
        index += 1


def find_sheet_case_insensitive(wb: xw.Book, name: str) -> xw.Sheet | None:
    wanted = name.strip().casefold()
    for sheet in wb.sheets:
        if sheet.name.strip().casefold() == wanted:
            return sheet
    return None


def find_max_anchor(sheet: xw.Sheet) -> tuple[int, int] | None:
    used = sheet.used_range
    values = ensure_2d(used.value)
    start_row = used.row
    start_col = used.column

    for r_idx, row in enumerate(values):
        for c_idx, value in enumerate(row):
            if isinstance(value, str) and value.strip().casefold() == "max":
                return start_row + r_idx, start_col + c_idx
    return None


def set_formula2(target_range: xw.Range, formula: str) -> None:
    try:
        target_range.formula2 = formula
    except Exception:
        target_range.formula = formula


def safe_close_workbook(wb: xw.Book) -> None:
    try:
        wb.close(save=False)
        return
    except TypeError:
        pass
    except Exception:
        pass

    try:
        wb.api.Close(SaveChanges=False)
        return
    except Exception:
        pass

    try:
        wb.close()
    except Exception:
        pass


def read_column_block(
    sheet: xw.Sheet, start_row: int, end_row: int, columns: list[int]
) -> dict[int, list[Any]]:
    if end_row < start_row:
        return {col: [] for col in columns}
    left = min(columns)
    right = max(columns)
    block = ensure_2d(sheet.range((start_row, left), (end_row, right)).value)
    out: dict[int, list[Any]] = {}
    for col in columns:
        offset = col - left
        out[col] = [row[offset] if offset < len(row) else None for row in block]
    return out


def fallback_linear_regression(xs: list[float], ys: list[float]) -> tuple[float, float] | None:
    if len(xs) != len(ys) or len(xs) < 2:
        return None
    mean_x = sum(xs) / len(xs)
    mean_y = sum(ys) / len(ys)
    denom = sum((x - mean_x) ** 2 for x in xs)
    if denom == 0:
        return None
    slope = sum((x - mean_x) * (y - mean_y) for x, y in zip(xs, ys)) / denom
    intercept = mean_y - (slope * mean_x)
    return intercept, slope


def collect_empirical_rows(wb: xw.Book, meta: FileMeta, source_file: str) -> list[dict[str, Any]]:
    rows: list[dict[str, Any]] = []
    sheet = find_sheet_case_insensitive(wb, "Empirical Model")
    if sheet is None:
        print(f"Skipping empirical extraction for {source_file} (missing 'Empirical Model' sheet)")
        return rows

    anchor = find_max_anchor(sheet)
    if anchor is None:
        print(f"Skipping empirical extraction for {source_file} (missing 'max' anchor)")
        return rows

    anchor_row, anchor_col = anchor
    history_end_row = anchor_row - 1
    n_quarters = 10
    if history_end_row < 2:
        return rows

    # Anchor-based offsets from 'max'.
    quarter_col = anchor_col - 12
    quarterly_sales_col = anchor_col - 11
    reported_sales_col = anchor_col - 7
    growth_rate_col = anchor_col - 6
    sales_captured_col = anchor_col - 5
    penetration_col = anchor_col - 4

    history_start_row = max(1, history_end_row - n_quarters + 1)
    columns = [
        quarter_col,
        quarterly_sales_col,
        reported_sales_col,
        growth_rate_col,
        sales_captured_col,
        penetration_col,
    ]
    if min(columns) < 1:
        print(
            f"Skipping empirical extraction for {source_file} "
            f"(invalid anchor offsets from column {anchor_col})"
        )
        return rows
    data = read_column_block(sheet, history_start_row, history_end_row, columns)

    quarter_values = data[quarter_col]
    quarterly_sales_values = data[quarterly_sales_col]
    reported_sales_values = data[reported_sales_col]
    growth_values = data[growth_rate_col]
    captured_values = data[sales_captured_col]
    penetration_values = data[penetration_col]

    available = len(quarter_values)
    if available == 0:
        return rows
    max_n = min(n_quarters, available)

    helper_col = max(sheet.used_range.last_cell.column + 2, anchor_col + 2)
    helper_row = max(anchor_row, 2)
    avg_cell = sheet.range((helper_row, helper_col))

    max_cell = sheet.range((anchor_row, anchor_col + 1))
    min_cell = sheet.range((anchor_row + 1, anchor_col + 1))

    latest_quarterly_sales = as_float(quarterly_sales_values[-1])
    latest_reported_sales = as_float(reported_sales_values[-1])
    latest_growth = as_float(growth_values[-1])
    latest_captured = as_float(captured_values[-1])

    if (
        latest_captured is None
        and latest_reported_sales is not None
        and latest_quarterly_sales not in (None, 0)
    ):
        latest_captured = latest_reported_sales / latest_quarterly_sales

    for num_quarters_used in range(1, max_n + 1):
        window_start_row = history_end_row - num_quarters_used + 1
        formula = (
            f"=AVERAGE(R{window_start_row}C{penetration_col}:"
            f"R{history_end_row}C{penetration_col})"
        )
        set_formula2(avg_cell, formula)
        wb.app.calculate()
        avg_penetration_pct = as_float(avg_cell.value)

        if avg_penetration_pct is None:
            raw = [as_float(v) for v in penetration_values[-num_quarters_used:]]
            numeric = [v for v in raw if v is not None]
            if numeric:
                avg_penetration_pct = sum(numeric) / len(numeric)

        forecast_value = None
        if avg_penetration_pct is not None and latest_quarterly_sales is not None:
            forecast_value = avg_penetration_pct * latest_quarterly_sales

        forecast_max = as_float(max_cell.value)
        forecast_min = as_float(min_cell.value)
        if forecast_max is None and forecast_value is not None:
            forecast_max = forecast_value * 1.05
        if forecast_min is None and forecast_value is not None:
            forecast_min = forecast_value * 0.95

        range_width = None
        if forecast_max is not None and forecast_min is not None:
            range_width = forecast_max - forecast_min

        last_quarter_used = quarter_values[-num_quarters_used]

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
                "actual_value": latest_reported_sales,
                "forecast_max": forecast_max,
                "forecast_min": forecast_min,
                "range_width": range_width,
                "avg_penetration_pct": avg_penetration_pct,
                "quarterly_sales": latest_quarterly_sales,
                "reported_sales": latest_reported_sales,
                "growth_rate_pct": latest_growth,
                "sales_captured_in_db_pct": latest_captured,
                "source_file": source_file,
            }
        )

    return rows


def collect_regression_rows(wb: xw.Book, meta: FileMeta, source_file: str) -> list[dict[str, Any]]:
    rows: list[dict[str, Any]] = []
    sheet = find_sheet_case_insensitive(wb, "Regression Model")
    if sheet is None:
        print(f"Skipping regression extraction for {source_file} (missing 'Regression Model' sheet)")
        return rows

    anchor = find_max_anchor(sheet)
    if anchor is None:
        print(f"Skipping regression extraction for {source_file} (missing 'max' anchor)")
        return rows

    anchor_row, anchor_col = anchor
    history_end_row = anchor_row - 1
    if history_end_row < 2:
        return rows

    n_quarters = 10
    # Required offsets from anchor.
    y_col = anchor_col - 7
    x_col = anchor_col - 11
    if x_col < 1 or y_col < 1:
        print(
            f"Skipping regression extraction for {source_file} "
            f"(invalid anchor offsets from column {anchor_col})"
        )
        return rows

    history_start_row = max(1, history_end_row - n_quarters + 1)
    data = read_column_block(sheet, history_start_row, history_end_row, [x_col, y_col])
    x_values_raw = data[x_col]
    y_values_raw = data[y_col]

    available = min(len(x_values_raw), len(y_values_raw))
    if available < 2:
        return rows
    max_n = min(n_quarters, available)

    helper_col = max(sheet.used_range.last_cell.column + 2, anchor_col + 2)
    helper_row = max(anchor_row, 2)
    intercept_cell = sheet.range((helper_row, helper_col))
    slope_cell = sheet.range((helper_row + 1, helper_col))

    forecast_max_cell = sheet.range((anchor_row, anchor_col + 1))
    forecast_min_cell = sheet.range((anchor_row + 1, anchor_col + 1))

    next_x = as_float(sheet.range((history_end_row + 1, x_col)).value)
    if next_x is None:
        last_x = as_float(x_values_raw[-1])
        if last_x is not None:
            next_x = last_x + 1

    actual_value = as_float(sheet.range((history_end_row + 1, y_col)).value)

    previous_signature: tuple[Any, ...] | None = None
    for num_quarters_used in range(2, max_n + 1):
        window_start_row = history_end_row - num_quarters_used + 1
        intercept_formula = (
            f"=INTERCEPT(R{window_start_row}C{y_col}:R{history_end_row}C{y_col},"
            f"R{window_start_row}C{x_col}:R{history_end_row}C{x_col})"
        )
        slope_formula = (
            f"=SLOPE(R{window_start_row}C{y_col}:R{history_end_row}C{y_col},"
            f"R{window_start_row}C{x_col}:R{history_end_row}C{x_col})"
        )
        set_formula2(intercept_cell, intercept_formula)
        set_formula2(slope_cell, slope_formula)
        wb.app.calculate()

        intercept = as_float(intercept_cell.value)
        slope = as_float(slope_cell.value)

        if intercept is None or slope is None:
            x_slice = [as_float(v) for v in x_values_raw[-num_quarters_used:]]
            y_slice = [as_float(v) for v in y_values_raw[-num_quarters_used:]]
            pairs = [(x, y) for x, y in zip(x_slice, y_slice) if x is not None and y is not None]
            if len(pairs) >= 2:
                fallback = fallback_linear_regression(
                    [pair[0] for pair in pairs], [pair[1] for pair in pairs]
                )
                if fallback is not None:
                    intercept, slope = fallback

        forecast_total_without_sa = None
        if intercept is not None and slope is not None and next_x is not None:
            forecast_total_without_sa = intercept + (slope * next_x)

        forecast_max = as_float(forecast_max_cell.value)
        forecast_min = as_float(forecast_min_cell.value)
        if forecast_max is None and forecast_total_without_sa is not None:
            forecast_max = forecast_total_without_sa * 1.05
        if forecast_min is None and forecast_total_without_sa is not None:
            forecast_min = forecast_total_without_sa * 0.95

        range_width = None
        if forecast_max is not None and forecast_min is not None:
            range_width = forecast_max - forecast_min

        signature = (
            round(intercept, 10) if intercept is not None else None,
            round(slope, 10) if slope is not None else None,
            round(forecast_total_without_sa, 10)
            if forecast_total_without_sa is not None
            else None,
            round(forecast_max, 10) if forecast_max is not None else None,
            round(forecast_min, 10) if forecast_min is not None else None,
        )
        if signature == previous_signature:
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
                "forecast_value": forecast_total_without_sa,
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


def write_sheet(ws: Any, headers: list[str], rows: list[dict[str, Any]]) -> None:
    ws.append(headers)
    for row in rows:
        ws.append([row.get(column) for column in headers])

    for cell in ws[1]:
        cell.font = Font(bold=True)

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions

    for col_idx, header in enumerate(headers, start=1):
        max_len = len(header)
        for row_idx in range(2, ws.max_row + 1):
            value = ws.cell(row=row_idx, column=col_idx).value
            if value is None:
                continue
            value_len = len(str(value))
            if value_len > max_len:
                max_len = value_len
        ws.column_dimensions[get_column_letter(col_idx)].width = min(max_len + 2, 40)


def write_output_workbook(
    output_path: Path,
    empirical_rows: list[dict[str, Any]],
    regression_rows: list[dict[str, Any]],
) -> None:
    out_wb = Workbook()
    default = out_wb.active
    out_wb.remove(default)

    empirical_ws = out_wb.create_sheet("empirical_candidates")
    regression_ws = out_wb.create_sheet("regression_candidates")

    write_sheet(empirical_ws, EMPIRICAL_HEADERS, empirical_rows)
    write_sheet(regression_ws, REGRESSION_HEADERS, regression_rows)

    out_wb.save(output_path)


def should_skip_file(path: Path) -> tuple[bool, str]:
    if not path.is_file():
        return True, "not a file"
    if path.name.startswith("~"):
        return True, "temporary file"
    if path.suffix.lower() != ".xlsx":
        return True, "not .xlsx"
    if OUTPUT_FILE_RE.search(path.name):
        return True, "output workbook pattern"
    return False, ""


def process_all_workbooks(input_folder: Path, destination_folder: Path) -> tuple[Path, int, int, int]:
    output_path = next_output_path(input_folder, destination_folder)

    empirical_rows: list[dict[str, Any]] = []
    regression_rows: list[dict[str, Any]] = []
    files_processed = 0

    app = xw.App(visible=False, add_book=False)
    app.display_alerts = False
    try:
        app.screen_updating = False
    except Exception:
        pass
    try:
        app.calculation = "manual"
    except Exception:
        pass

    try:
        for file_path in sorted(input_folder.iterdir()):
            skip, reason = should_skip_file(file_path)
            if skip:
                print(f"Skipping file: {file_path.name} ({reason})")
                continue

            print(f"Processing file: {file_path.name}")
            wb: xw.Book | None = None
            try:
                wb = app.books.open(str(file_path), update_links=False)
                meta = parse_file_meta(file_path.name)
                empirical_rows.extend(collect_empirical_rows(wb, meta, file_path.name))
                regression_rows.extend(collect_regression_rows(wb, meta, file_path.name))
                files_processed += 1
            except Exception as exc:
                print(f"Skipping file: {file_path.name} (processing error: {exc})")
            finally:
                if wb is not None:
                    safe_close_workbook(wb)
    finally:
        try:
            app.quit()
        except Exception:
            pass

    write_output_workbook(output_path, empirical_rows, regression_rows)
    return output_path, files_processed, len(empirical_rows), len(regression_rows)


def main() -> None:
    input_folder = Path(input_dir).expanduser().resolve()
    destination_folder = Path(output_dir).expanduser().resolve()

    if not input_folder.exists() or not input_folder.is_dir():
        raise SystemExit(f"Input directory does not exist: {input_folder}")
    destination_folder.mkdir(parents=True, exist_ok=True)

    output_path, files_processed, empirical_count, regression_count = process_all_workbooks(
        input_folder, destination_folder
    )

    print(f"Output path: {output_path}")
    print(f"Number of files processed: {files_processed}")
    print(f"Number of empirical rows: {empirical_count}")
    print(f"Number of regression rows: {regression_count}")


if __name__ == "__main__":
    main()
