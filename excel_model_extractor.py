from __future__ import annotations

from dataclasses import dataclass
from datetime import date
from pathlib import Path
import re
import statistics
from typing import Any

import xlwings as xw
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter


# Set these two folders before running.
input_dir = Path("/path/to/input")
output_dir = Path("/path/to/output")


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


@dataclass(frozen=True)
class FileLabel:
    model: str
    ticker: str
    model_period: str
    model_date: str


def to_float(value: Any) -> float | None:
    if value is None or value == "":
        return None
    if isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        return float(value)
    if isinstance(value, str):
        stripped = value.replace(",", "").strip()
        if stripped == "":
            return None
        try:
            return float(stripped)
        except ValueError:
            return None
    return None


def normalize_matrix(values: Any) -> list[list[Any]]:
    if values is None:
        return []
    if isinstance(values, list):
        if not values:
            return []
        if isinstance(values[0], list):
            return values
        return [values]
    return [[values]]


def normalize_percent(value: float | None) -> float | None:
    if value is None:
        return None
    return value / 100.0 if value > 1 else value


def month_number(month_text: str) -> int | None:
    key = month_text.strip().lower()[:3]
    month_map = {
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
    return month_map.get(key)


def parse_file_label(file_name: str) -> FileLabel | None:
    stem = Path(file_name).stem
    parts = [part.strip() for part in stem.split(" - ")]
    if len(parts) < 3:
        return None

    ticker = parts[1].upper()
    period_token = parts[2].split("_")[0]
    match = re.search(r"(Early|Mid|Late)([A-Za-z]+)(\d{4})", period_token, re.IGNORECASE)
    if not match:
        return None

    period_prefix = match.group(1).title()
    month_text = match.group(2).title()
    year = int(match.group(3))
    month_num = month_number(month_text)
    if month_num is None:
        return None

    day_map = {"Early": 5, "Mid": 15, "Late": 25}
    day = day_map[period_prefix]
    model_period = f"{period_prefix}{month_text[:3]}_{year}"
    model_date = date(year, month_num, day).isoformat()
    model = f"{ticker}_{model_period}"

    return FileLabel(model=model, ticker=ticker, model_period=model_period, model_date=model_date)


def next_output_path(input_folder: Path, out_folder: Path) -> Path:
    base_name = f"{input_folder.name}_PARAM"
    candidate = out_folder / f"{base_name}.xlsx"
    if not candidate.exists():
        return candidate

    index = 1
    while True:
        candidate = out_folder / f"{base_name}.{index}.xlsx"
        if not candidate.exists():
            return candidate
        index += 1


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


def find_anchor(sheet: xw.Sheet, target: str = "max") -> tuple[int, int] | None:
    used = sheet.used_range
    matrix = normalize_matrix(used.options(ndim=2).value)
    if not matrix:
        return None

    start_row = used.row
    start_col = used.column
    target_lower = target.strip().lower()

    for r_idx, row in enumerate(matrix):
        for c_idx, value in enumerate(row):
            if isinstance(value, str) and value.strip().lower() == target_lower:
                return start_row + r_idx, start_col + c_idx
    return None


def read_column_values(sheet: xw.Sheet, row_start: int, row_end: int, col: int) -> list[Any]:
    if row_end < row_start:
        return []
    values = sheet.range((row_start, col), (row_end, col)).value
    if isinstance(values, list):
        return values
    return [values]


def match_header_columns(sheet: xw.Sheet, header_row: int, anchor_col: int) -> dict[str, int]:
    left_col = max(1, anchor_col - 20)
    right_col = anchor_col + 8
    headers = sheet.range((header_row, left_col), (header_row, right_col)).value
    if not isinstance(headers, list):
        headers = [headers]

    aliases: dict[str, tuple[str, ...]] = {
        "num_quarters": ("num quarter", "quarters used", "n quarter"),
        "last_quarter": ("last quarter", "quarter", "qtr"),
        "forecast_value": ("estimated total sold", "total sold", "forecast"),
        "actual_value": ("actual", "reported sales"),
        "reported_sales": ("reported sales", "reported"),
        "quarterly_sales": ("quarterly sales", "quarter sales", "qtr sales"),
        "growth_rate_pct": ("growth rate", "growth%"),
        "sales_captured_in_db_pct": ("captured in db", "sales captured"),
        "penetration_pct": ("penetration", "pen%"),
        "max": ("max",),
        "min": ("min",),
    }

    matched: dict[str, int] = {}
    for key, needles in aliases.items():
        for idx, raw_header in enumerate(headers):
            if not isinstance(raw_header, str):
                continue
            normalized = raw_header.strip().lower()
            if any(needle in normalized for needle in needles):
                matched[key] = left_col + idx
                break
    return matched


def extract_empirical_candidates(workbook: xw.Book, file_label: FileLabel, source_file: str) -> list[dict[str, Any]]:
    if "Empirical Model" not in [sheet.name for sheet in workbook.sheets]:
        return []

    sheet = workbook.sheets["Empirical Model"]
    anchor = find_anchor(sheet, target="max")
    if anchor is None:
        return []

    anchor_row, anchor_col = anchor
    header_map = match_header_columns(sheet, anchor_row, anchor_col)

    penetration_col = header_map.get("penetration_pct", anchor_col - 6)
    quarter_col = header_map.get("last_quarter", anchor_col - 12)
    quarterly_sales_col = header_map.get("quarterly_sales", anchor_col - 10)
    reported_sales_col = header_map.get("reported_sales", anchor_col - 9)
    growth_rate_col = header_map.get("growth_rate_pct", anchor_col - 8)
    captured_col = header_map.get("sales_captured_in_db_pct", anchor_col - 7)
    forecast_col = header_map.get("forecast_value", anchor_col - 3)
    min_col = header_map.get("min", anchor_col + 1)
    max_col = header_map.get("max", anchor_col)
    n_quarters_col = header_map.get("num_quarters", anchor_col - 13)

    lookback_start = max(1, anchor_row - 200)
    penetration_values = read_column_values(sheet, lookback_start, anchor_row - 1, penetration_col)

    pen_rows: list[tuple[int, float]] = []
    for idx, value in enumerate(penetration_values, start=lookback_start):
        number = to_float(value)
        if number is not None:
            pen_rows.append((idx, number))

    if not pen_rows:
        return []

    max_n = min(10, len(pen_rows))
    temp_col = anchor_col + 30
    temp_start_row = anchor_row + 2
    contexts: list[dict[str, Any]] = []

    for offset, n_quarters in enumerate(range(1, max_n + 1)):
        selected = pen_rows[-n_quarters:]
        first_row = selected[0][0]
        last_row = selected[-1][0]
        target_row = temp_start_row + offset
        formula = f"=AVERAGE(R{first_row}C{penetration_col}:R{last_row}C{penetration_col})"
        sheet.cells(target_row, temp_col).formula2 = formula
        contexts.append(
            {
                "n_quarters": n_quarters,
                "first_row": first_row,
                "last_row": last_row,
                "selected_pen_values": [value for _, value in selected],
                "target_row": target_row,
            }
        )

    workbook.app.calculate()
    avg_values = read_column_values(sheet, temp_start_row, temp_start_row + len(contexts) - 1, temp_col)

    rows: list[dict[str, Any]] = []
    for idx, context in enumerate(contexts):
        n_quarters = context["n_quarters"]
        last_row = context["last_row"]
        output_row = anchor_row + n_quarters
        avg_pen_raw = to_float(avg_values[idx] if idx < len(avg_values) else None)
        avg_pen_fraction = normalize_percent(avg_pen_raw)

        quarterly_sales = to_float(sheet.cells(last_row, quarterly_sales_col).value)
        reported_sales = to_float(sheet.cells(last_row, reported_sales_col).value)
        growth_rate = to_float(sheet.cells(last_row, growth_rate_col).value)
        captured_pct = to_float(sheet.cells(last_row, captured_col).value)
        last_quarter_label = sheet.cells(last_row, quarter_col).value

        forecast_value = to_float(sheet.cells(output_row, forecast_col).value)
        forecast_max = to_float(sheet.cells(output_row, max_col).value)
        forecast_min = to_float(sheet.cells(output_row, min_col).value)

        selected_pen = [normalize_percent(v) for v in context["selected_pen_values"]]
        selected_pen = [v for v in selected_pen if v is not None and v > 0]

        if forecast_value is None and quarterly_sales is not None and avg_pen_fraction and avg_pen_fraction > 0:
            forecast_value = quarterly_sales / avg_pen_fraction

        if (forecast_max is None or forecast_min is None) and quarterly_sales is not None and selected_pen:
            min_pen = min(selected_pen)
            max_pen = max(selected_pen)
            if min_pen > 0:
                forecast_max = quarterly_sales / min_pen
            if max_pen > 0:
                forecast_min = quarterly_sales / max_pen

        range_width = None
        if forecast_max is not None and forecast_min is not None:
            range_width = forecast_max - forecast_min

        used_quarters = to_float(sheet.cells(output_row, n_quarters_col).value)
        if used_quarters is None:
            used_quarters = float(n_quarters)

        rows.append(
            {
                "model": file_label.model,
                "ticker": file_label.ticker,
                "model_period": file_label.model_period,
                "model_date": file_label.model_date,
                "method": "empirical",
                "parameter_name": "avg_penetration_pct",
                "parameter_value": avg_pen_raw,
                "num_quarters_used": int(used_quarters),
                "last_quarter_used": last_quarter_label,
                "forecast_value": forecast_value,
                "actual_value": reported_sales,
                "forecast_max": forecast_max,
                "forecast_min": forecast_min,
                "range_width": range_width,
                "avg_penetration_pct": avg_pen_raw,
                "quarterly_sales": quarterly_sales,
                "reported_sales": reported_sales,
                "growth_rate_pct": growth_rate,
                "sales_captured_in_db_pct": captured_pct,
                "source_file": source_file,
            }
        )

    return rows


def extract_regression_candidates(workbook: xw.Book, file_label: FileLabel, source_file: str) -> list[dict[str, Any]]:
    if "Regression Model" not in [sheet.name for sheet in workbook.sheets]:
        return []

    sheet = workbook.sheets["Regression Model"]
    anchor = find_anchor(sheet, target="max")
    if anchor is None:
        return []

    anchor_row, anchor_col = anchor
    y_col = anchor_col - 7
    x_col = anchor_col - 11

    lookback_start = max(1, anchor_row - 200)
    x_values = read_column_values(sheet, lookback_start, anchor_row - 1, x_col)
    y_values = read_column_values(sheet, lookback_start, anchor_row - 1, y_col)

    history: list[tuple[int, float, float]] = []
    for row_idx, (x_raw, y_raw) in enumerate(zip(x_values, y_values), start=lookback_start):
        x_num = to_float(x_raw)
        y_num = to_float(y_raw)
        if x_num is not None and y_num is not None:
            history.append((row_idx, x_num, y_num))

    if len(history) < 2:
        return []

    max_n = min(10, len(history))
    temp_col = anchor_col + 35
    temp_start_row = anchor_row + 2
    contexts: list[dict[str, Any]] = []

    for offset, n_quarters in enumerate(range(2, max_n + 1)):
        selected = history[-n_quarters:]
        first_row = selected[0][0]
        last_row = selected[-1][0]
        row = temp_start_row + offset

        intercept_formula = (
            f"=INTERCEPT(R{first_row}C{y_col}:R{last_row}C{y_col},R{first_row}C{x_col}:R{last_row}C{x_col})"
        )
        slope_formula = (
            f"=SLOPE(R{first_row}C{y_col}:R{last_row}C{y_col},R{first_row}C{x_col}:R{last_row}C{x_col})"
        )

        sheet.cells(row, temp_col).formula2 = intercept_formula
        sheet.cells(row, temp_col + 1).formula2 = slope_formula
        contexts.append({"n_quarters": n_quarters, "selected": selected, "target_row": row})

    workbook.app.calculate()

    intercept_values = sheet.range(
        (temp_start_row, temp_col), (temp_start_row + len(contexts) - 1, temp_col)
    ).value
    slope_values = sheet.range(
        (temp_start_row, temp_col + 1), (temp_start_row + len(contexts) - 1, temp_col + 1)
    ).value

    if not isinstance(intercept_values, list):
        intercept_values = [intercept_values]
    if not isinstance(slope_values, list):
        slope_values = [slope_values]

    rows: list[dict[str, Any]] = []
    last_signature: tuple[Any, ...] | None = None

    for idx, context in enumerate(contexts):
        n_quarters = context["n_quarters"]
        selected = context["selected"]
        intercept = to_float(intercept_values[idx] if idx < len(intercept_values) else None)
        slope = to_float(slope_values[idx] if idx < len(slope_values) else None)
        if intercept is None or slope is None:
            continue

        latest_x = selected[-1][1]
        forecast_fallback = intercept + slope * latest_x
        residuals = [y - (intercept + slope * x) for _, x, y in selected]
        std_dev = statistics.pstdev(residuals) if len(residuals) > 1 else 0.0
        max_fallback = forecast_fallback + std_dev
        min_fallback = forecast_fallback - std_dev

        output_row = anchor_row + n_quarters
        forecast_value = to_float(sheet.cells(output_row, anchor_col - 1).value)
        forecast_max = to_float(sheet.cells(output_row, anchor_col).value)
        forecast_min = to_float(sheet.cells(output_row, anchor_col + 1).value)
        actual_value = to_float(sheet.cells(output_row, anchor_col - 2).value)

        if forecast_value is None:
            forecast_value = forecast_fallback
        if forecast_max is None:
            forecast_max = max_fallback
        if forecast_min is None:
            forecast_min = min_fallback

        range_width = None
        if forecast_max is not None and forecast_min is not None:
            range_width = forecast_max - forecast_min

        signature = (
            round(intercept, 8),
            round(slope, 8),
            round(forecast_value, 8) if forecast_value is not None else None,
            round(forecast_max, 8) if forecast_max is not None else None,
            round(forecast_min, 8) if forecast_min is not None else None,
        )
        if signature == last_signature:
            continue
        last_signature = signature

        rows.append(
            {
                "model": file_label.model,
                "ticker": file_label.ticker,
                "model_period": file_label.model_period,
                "model_date": file_label.model_date,
                "method": "regression",
                "parameter_name": "num_quarters_used",
                "parameter_value": n_quarters,
                "num_quarters_used": n_quarters,
                "forecast_value": forecast_value,
                "actual_value": actual_value if actual_value is not None else "",
                "forecast_max": forecast_max,
                "forecast_min": forecast_min,
                "range_width": range_width,
                "intercept": intercept,
                "slope": slope,
                "source_file": source_file,
            }
        )

    return rows


def write_output_sheet(workbook: Workbook, sheet_name: str, headers: list[str], rows: list[dict[str, Any]]) -> None:
    sheet = workbook.create_sheet(title=sheet_name)
    sheet.append(headers)
    for cell in sheet[1]:
        cell.font = Font(bold=True)

    for row in rows:
        sheet.append([row.get(header, "") for header in headers])

    sheet.freeze_panes = "A2"
    sheet.auto_filter.ref = sheet.dimensions

    for col_idx, header in enumerate(headers, start=1):
        width = len(header) + 2
        for row_idx in range(2, sheet.max_row + 1):
            value = sheet.cell(row=row_idx, column=col_idx).value
            if value is None:
                continue
            width = max(width, len(str(value)) + 2)
        sheet.column_dimensions[get_column_letter(col_idx)].width = max(12, min(width, 42))


def ensure_directories() -> None:
    if not input_dir.exists() or not input_dir.is_dir():
        raise FileNotFoundError(f"input_dir does not exist or is not a directory: {input_dir}")
    output_dir.mkdir(parents=True, exist_ok=True)


def run() -> None:
    ensure_directories()

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

        for file_path in sorted(input_dir.iterdir()):
            if not file_path.is_file():
                continue
            file_name = file_path.name

            if file_name.startswith("~"):
                print(f"Skipped file: {file_name} (temporary file)")
                continue
            if file_path.suffix.lower() != ".xlsx":
                print(f"Skipped file: {file_name} (not .xlsx)")
                continue

            label = parse_file_label(file_name)
            if label is None:
                print(f"Skipped file: {file_name} (could not parse ticker/model period)")
                continue

            workbook = None
            try:
                workbook = app.books.open(str(file_path), update_links=False)
                file_empirical_rows = extract_empirical_candidates(workbook, label, file_name)
                file_regression_rows = extract_regression_candidates(workbook, label, file_name)
                empirical_rows.extend(file_empirical_rows)
                regression_rows.extend(file_regression_rows)
                files_processed += 1
                print(
                    f"Processed file: {file_name} "
                    f"(empirical_rows={len(file_empirical_rows)}, regression_rows={len(file_regression_rows)})"
                )
            except Exception as exc:
                print(f"Skipped file: {file_name} (processing error: {exc})")
            finally:
                if workbook is not None:
                    safe_close_workbook(workbook)

    finally:
        app.quit()

    output_path = next_output_path(input_dir, output_dir)
    out_book = Workbook()
    default_sheet = out_book.active
    out_book.remove(default_sheet)
    write_output_sheet(out_book, "empirical_candidates", EMPIRICAL_HEADERS, empirical_rows)
    write_output_sheet(out_book, "regression_candidates", REGRESSION_HEADERS, regression_rows)
    out_book.save(output_path)

    print(f"Output path: {output_path}")
    print(f"Number of files processed: {files_processed}")
    print(f"Number of empirical rows: {len(empirical_rows)}")
    print(f"Number of regression rows: {len(regression_rows)}")


if __name__ == "__main__":
    run()
