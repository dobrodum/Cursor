#!/usr/bin/env python3
from __future__ import annotations

import datetime as dt
import math
import re
import statistics
from dataclasses import dataclass
from pathlib import Path
from typing import Any

import xlwings as xw
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

# Configure these two paths before running.
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

MONTH_LOOKUP = {
    "jan": (1, "Jan"),
    "feb": (2, "Feb"),
    "mar": (3, "Mar"),
    "apr": (4, "Apr"),
    "may": (5, "May"),
    "jun": (6, "Jun"),
    "jul": (7, "Jul"),
    "aug": (8, "Aug"),
    "sep": (9, "Sep"),
    "oct": (10, "Oct"),
    "nov": (11, "Nov"),
    "dec": (12, "Dec"),
}

PHASE_DAY = {"early": 5, "mid": 15, "late": 25}


@dataclass(frozen=True)
class ModelMetadata:
    model: str
    ticker: str
    model_period: str
    model_date: str


@dataclass
class SheetSnapshot:
    values: list[list[Any]]
    start_row: int
    start_col: int
    last_row: int
    last_col: int
    labels: list[tuple[str, int, int]]


def normalize_text(value: Any) -> str:
    if value is None:
        return ""
    return " ".join(str(value).strip().lower().split())


def normalize_2d(values: Any) -> list[list[Any]]:
    if values is None:
        return []
    if isinstance(values, tuple):
        values = list(values)
    if not isinstance(values, list):
        return [[values]]
    if not values:
        return []
    if isinstance(values[0], list):
        output: list[list[Any]] = []
        for row in values:
            if isinstance(row, tuple):
                output.append(list(row))
            elif isinstance(row, list):
                output.append(row)
            else:
                output.append([row])
        return output
    return [list(values)]


def to_float(value: Any) -> float | None:
    if value is None or isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        if isinstance(value, float) and math.isnan(value):
            return None
        return float(value)
    if not isinstance(value, str):
        return None

    raw = value.strip()
    if not raw:
        return None
    percent = raw.endswith("%")
    cleaned = raw.replace(",", "").replace("%", "")
    try:
        number = float(cleaned)
    except ValueError:
        return None
    if percent:
        return number / 100.0
    return number


def safe_divide(numerator: float | None, denominator: float | None) -> float | None:
    if numerator is None or denominator in (None, 0):
        return None
    return numerator / denominator


def build_output_path(src_input_dir: Path, src_output_dir: Path) -> Path:
    base_stem = f"{src_input_dir.name}_PARAM"
    candidate = src_output_dir / f"{base_stem}.xlsx"
    suffix = 1
    while candidate.exists():
        candidate = src_output_dir / f"{base_stem}.{suffix}.xlsx"
        suffix += 1
    return candidate


def parse_file_metadata(file_path: Path) -> ModelMetadata:
    stem = file_path.stem
    parts = [part.strip() for part in stem.split(" - ")]

    ticker = "UNKNOWN"
    if len(parts) >= 2 and parts[1]:
        ticker = parts[1].upper()

    period_token = ""
    if len(parts) >= 3:
        period_token = re.sub(r"[_\s-]*send$", "", parts[2], flags=re.IGNORECASE).strip()

    model_period = "Unknown_0000"
    model_date = ""
    match = re.search(r"(Early|Mid|Late)([A-Za-z]{3,9})(\d{4})", period_token, flags=re.IGNORECASE)
    if match:
        phase_raw = match.group(1)
        month_raw = match.group(2)
        year = int(match.group(3))

        month_key = month_raw[:3].lower()
        month_info = MONTH_LOOKUP.get(month_key)
        if month_info:
            month_num, month_abbrev = month_info
            phase = phase_raw.capitalize()
            model_period = f"{phase}{month_abbrev}_{year}"
            day = PHASE_DAY[phase.lower()]
            model_date = dt.date(year, month_num, day).isoformat()
    elif period_token:
        normalized_period = re.sub(r"[^A-Za-z0-9]+", "", period_token)
        if normalized_period:
            model_period = normalized_period

    model = f"{ticker}_{model_period}"
    return ModelMetadata(model=model, ticker=ticker, model_period=model_period, model_date=model_date)


def take_snapshot(sheet: xw.Sheet) -> SheetSnapshot:
    used = sheet.used_range
    values = normalize_2d(used.value)
    start_row = used.row
    start_col = used.column
    last_row = start_row + len(values) - 1 if values else start_row
    width = max((len(row) for row in values), default=1)
    last_col = start_col + width - 1

    labels: list[tuple[str, int, int]] = []
    for row_idx, row in enumerate(values):
        for col_idx, value in enumerate(row):
            text = normalize_text(value)
            if text:
                labels.append((text, start_row + row_idx, start_col + col_idx))

    return SheetSnapshot(
        values=values,
        start_row=start_row,
        start_col=start_col,
        last_row=last_row,
        last_col=last_col,
        labels=labels,
    )


def snapshot_value(snapshot: SheetSnapshot, row: int, col: int) -> Any:
    row_idx = row - snapshot.start_row
    col_idx = col - snapshot.start_col
    if row_idx < 0 or col_idx < 0 or row_idx >= len(snapshot.values):
        return None
    row_values = snapshot.values[row_idx]
    if col_idx >= len(row_values):
        return None
    return row_values[col_idx]


def find_anchor_cell(snapshot: SheetSnapshot, target: str = "max") -> tuple[int | None, int | None]:
    wanted = normalize_text(target)
    best: tuple[int, int, int] | None = None

    for row_idx, row in enumerate(snapshot.values):
        for col_idx, value in enumerate(row):
            if normalize_text(value) != wanted:
                continue

            abs_row = snapshot.start_row + row_idx
            abs_col = snapshot.start_col + col_idx
            score = 0

            right = snapshot_value(snapshot, abs_row, abs_col + 1)
            if to_float(right) is not None:
                score += 2

            for jump in range(1, 7):
                below = snapshot_value(snapshot, abs_row + jump, abs_col)
                if normalize_text(below) == "min":
                    score += 1
                    break

            if best is None or score > best[0]:
                best = (score, abs_row, abs_col)

    if best is None:
        return None, None
    return best[1], best[2]


def find_column_from_keywords(
    snapshot: SheetSnapshot,
    anchor_row: int,
    anchor_col: int,
    keywords: list[str],
    fallback_col: int,
) -> int:
    best: tuple[float, int] | None = None
    for text, row, col in snapshot.labels:
        if any(keyword in text for keyword in keywords):
            score = abs(col - anchor_col) + (0.1 * abs(row - anchor_row))
            if best is None or score < best[0]:
                best = (score, col)
    return best[1] if best else fallback_col


def collect_numeric_rows(snapshot: SheetSnapshot, col: int, before_row: int) -> list[tuple[int, float]]:
    rows: list[tuple[int, float]] = []
    stop_row = min(before_row, snapshot.last_row + 1)
    for row in range(snapshot.start_row, stop_row):
        number = to_float(snapshot_value(snapshot, row, col))
        if number is not None:
            rows.append((row, number))
    return rows


def collect_xy_rows(snapshot: SheetSnapshot, x_col: int, y_col: int, before_row: int) -> list[tuple[int, float, float]]:
    rows: list[tuple[int, float, float]] = []
    stop_row = min(before_row, snapshot.last_row + 1)
    for row in range(snapshot.start_row, stop_row):
        x_val = to_float(snapshot_value(snapshot, row, x_col))
        y_val = to_float(snapshot_value(snapshot, row, y_col))
        if x_val is not None and y_val is not None:
            rows.append((row, x_val, y_val))
    return rows


def last_numeric_before(snapshot: SheetSnapshot, col: int, before_row: int) -> float | None:
    for row in range(min(before_row - 1, snapshot.last_row), snapshot.start_row - 1, -1):
        number = to_float(snapshot_value(snapshot, row, col))
        if number is not None:
            return number
    return None


def nearby_label_value(snapshot: SheetSnapshot, anchor_row: int, anchor_col: int, label: str) -> float | None:
    wanted = normalize_text(label)
    for text, row, col in snapshot.labels:
        if text == wanted and abs(col - anchor_col) <= 2 and abs(row - anchor_row) <= 8:
            value = to_float(snapshot_value(snapshot, row, col + 1))
            if value is not None:
                return value
    return None


def normalize_column_values(raw_values: Any, expected_length: int) -> list[Any]:
    if expected_length <= 0:
        return []
    if isinstance(raw_values, list):
        if raw_values and isinstance(raw_values[0], list):
            return [row[0] if row else None for row in raw_values]
        return raw_values
    return [raw_values]


def set_formula2(cell: xw.Range, formula: str) -> None:
    try:
        cell.formula2 = formula
    except Exception:
        cell.formula = formula


def close_workbook_without_saving(workbook: xw.Book) -> None:
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


def derive_forecast_band_from_penetration(
    quarterly_sales: float | None, min_penetration: float | None, max_penetration: float | None
) -> tuple[float | None, float | None]:
    if quarterly_sales is None:
        return None, None

    high_candidate = safe_divide(quarterly_sales, min_penetration)
    low_candidate = safe_divide(quarterly_sales, max_penetration)
    if high_candidate is None or low_candidate is None:
        return None, None
    return max(high_candidate, low_candidate), min(high_candidate, low_candidate)


def extract_empirical_rows(workbook: xw.Book, metadata: ModelMetadata, source_file: str) -> list[dict[str, Any]]:
    try:
        sheet = workbook.sheets["Empirical Model"]
    except Exception:
        print(f"skipped empirical in {source_file}: missing sheet 'Empirical Model'")
        return []

    snapshot = take_snapshot(sheet)
    anchor_row, anchor_col = find_anchor_cell(snapshot, target="max")
    if anchor_row is None or anchor_col is None:
        print(f"skipped empirical in {source_file}: could not find 'max' anchor")
        return []

    quarter_col = find_column_from_keywords(
        snapshot, anchor_row, anchor_col, ["quarter", "qtr"], fallback_col=anchor_col - 11
    )
    penetration_col = find_column_from_keywords(
        snapshot, anchor_row, anchor_col, ["penetration", "penet"], fallback_col=anchor_col - 3
    )
    quarterly_sales_col = find_column_from_keywords(
        snapshot,
        anchor_row,
        anchor_col,
        ["quarterly sales", "quarter sales", "sales"],
        fallback_col=anchor_col - 7,
    )
    reported_sales_col = find_column_from_keywords(
        snapshot,
        anchor_row,
        anchor_col,
        ["reported sales", "actual sales", "reported"],
        fallback_col=anchor_col - 6,
    )
    growth_rate_col = find_column_from_keywords(
        snapshot, anchor_row, anchor_col, ["growth rate", "growth"], fallback_col=anchor_col - 5
    )
    captured_pct_col = find_column_from_keywords(
        snapshot,
        anchor_row,
        anchor_col,
        ["captured in db", "captured", "database coverage"],
        fallback_col=anchor_col - 4,
    )

    penetration_history = collect_numeric_rows(snapshot, penetration_col, before_row=anchor_row)
    if not penetration_history:
        print(f"skipped empirical in {source_file}: no penetration history found")
        return []

    n_quarters = min(10, len(penetration_history))
    helper_start_row = anchor_row + 2
    helper_col = anchor_col + 20

    for idx, n_used in enumerate(range(1, n_quarters + 1)):
        start_row = penetration_history[-n_used][0]
        end_row = penetration_history[-1][0]
        formula = f"=AVERAGE(R{start_row}C{penetration_col}:R{end_row}C{penetration_col})"
        set_formula2(sheet.cells(helper_start_row + idx, helper_col), formula)

    workbook.app.calculate()

    avg_values_raw = sheet.range(
        (helper_start_row, helper_col),
        (helper_start_row + n_quarters - 1, helper_col),
    ).value
    avg_values = normalize_column_values(avg_values_raw, n_quarters)
    sheet.range(
        (helper_start_row, helper_col),
        (helper_start_row + n_quarters - 1, helper_col),
    ).clear_contents()

    quarterly_sales = last_numeric_before(snapshot, quarterly_sales_col, before_row=anchor_row)
    reported_sales = last_numeric_before(snapshot, reported_sales_col, before_row=anchor_row)
    growth_rate_pct = last_numeric_before(snapshot, growth_rate_col, before_row=anchor_row)
    sales_captured_pct = last_numeric_before(snapshot, captured_pct_col, before_row=anchor_row)

    max_penetration = to_float(snapshot_value(snapshot, anchor_row, anchor_col + 1))
    min_penetration = nearby_label_value(snapshot, anchor_row, anchor_col, "min")
    derived_max, derived_min = derive_forecast_band_from_penetration(
        quarterly_sales=quarterly_sales, min_penetration=min_penetration, max_penetration=max_penetration
    )

    rows: list[dict[str, Any]] = []
    for idx, n_used in enumerate(range(1, n_quarters + 1)):
        table_row = anchor_row + 1 + idx
        avg_penetration_pct = to_float(avg_values[idx]) if idx < len(avg_values) else None
        table_forecast = to_float(snapshot_value(snapshot, table_row, anchor_col - 2))
        table_actual = to_float(snapshot_value(snapshot, table_row, anchor_col - 3))
        table_max = to_float(snapshot_value(snapshot, table_row, anchor_col))
        table_min = to_float(snapshot_value(snapshot, table_row, anchor_col - 1))

        forecast_value = table_forecast
        if forecast_value is None:
            forecast_value = safe_divide(quarterly_sales, avg_penetration_pct)

        actual_value = table_actual if table_actual is not None else reported_sales
        forecast_max = table_max if table_max is not None else derived_max
        forecast_min = table_min if table_min is not None else derived_min
        range_width = (
            forecast_max - forecast_min
            if forecast_max is not None and forecast_min is not None
            else None
        )

        last_quarter_used = snapshot_value(snapshot, penetration_history[-n_used][0], quarter_col)
        row = {
            "model": metadata.model,
            "ticker": metadata.ticker,
            "model_period": metadata.model_period,
            "model_date": metadata.model_date,
            "method": "empirical",
            "parameter_name": "avg_penetration_pct",
            "parameter_value": avg_penetration_pct,
            "num_quarters_used": n_used,
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
        rows.append(row)

    return rows


def value_equal_or_close(left: Any, right: Any, tolerance: float = 1e-9) -> bool:
    if left is None and right is None:
        return True
    left_num = to_float(left)
    right_num = to_float(right)
    if left_num is not None and right_num is not None:
        return abs(left_num - right_num) <= tolerance
    return left == right


def is_duplicate_regression_row(previous: dict[str, Any], current: dict[str, Any]) -> bool:
    keys = [
        "num_quarters_used",
        "forecast_value",
        "forecast_max",
        "forecast_min",
        "intercept",
        "slope",
    ]
    return all(value_equal_or_close(previous.get(key), current.get(key)) for key in keys)


def extract_regression_rows(workbook: xw.Book, metadata: ModelMetadata, source_file: str) -> list[dict[str, Any]]:
    try:
        sheet = workbook.sheets["Regression Model"]
    except Exception:
        print(f"skipped regression in {source_file}: missing sheet 'Regression Model'")
        return []

    snapshot = take_snapshot(sheet)
    anchor_row, anchor_col = find_anchor_cell(snapshot, target="max")
    if anchor_row is None or anchor_col is None:
        print(f"skipped regression in {source_file}: could not find 'max' anchor")
        return []

    y_col = anchor_col - 7
    x_col = anchor_col - 11
    xy_history = collect_xy_rows(snapshot, x_col=x_col, y_col=y_col, before_row=anchor_row)
    if len(xy_history) < 2:
        print(f"skipped regression in {source_file}: not enough x/y rows")
        return []

    n_quarters = min(10, len(xy_history))
    helper_start_row = anchor_row + 2
    helper_col = anchor_col + 20

    for idx, n_used in enumerate(range(1, n_quarters + 1)):
        start_row = xy_history[-n_used][0]
        end_row = xy_history[-1][0]
        intercept_formula = (
            f"=INTERCEPT(R{start_row}C{y_col}:R{end_row}C{y_col},R{start_row}C{x_col}:R{end_row}C{x_col})"
        )
        slope_formula = (
            f"=SLOPE(R{start_row}C{y_col}:R{end_row}C{y_col},R{start_row}C{x_col}:R{end_row}C{x_col})"
        )
        set_formula2(sheet.cells(helper_start_row + idx, helper_col), intercept_formula)
        set_formula2(sheet.cells(helper_start_row + idx, helper_col + 1), slope_formula)

    workbook.app.calculate()

    intercept_raw = sheet.range(
        (helper_start_row, helper_col),
        (helper_start_row + n_quarters - 1, helper_col),
    ).value
    slope_raw = sheet.range(
        (helper_start_row, helper_col + 1),
        (helper_start_row + n_quarters - 1, helper_col + 1),
    ).value
    intercept_values = normalize_column_values(intercept_raw, n_quarters)
    slope_values = normalize_column_values(slope_raw, n_quarters)

    sheet.range(
        (helper_start_row, helper_col),
        (helper_start_row + n_quarters - 1, helper_col + 1),
    ).clear_contents()

    actual_col = find_column_from_keywords(
        snapshot,
        anchor_row,
        anchor_col,
        ["actual", "reported"],
        fallback_col=anchor_col - 6,
    )
    default_actual = last_numeric_before(snapshot, actual_col, before_row=anchor_row)

    rows: list[dict[str, Any]] = []
    for idx, n_used in enumerate(range(1, n_quarters + 1)):
        intercept = to_float(intercept_values[idx]) if idx < len(intercept_values) else None
        slope = to_float(slope_values[idx]) if idx < len(slope_values) else None
        if intercept is None or slope is None:
            continue

        table_row = anchor_row + 1 + idx
        table_num_quarters = to_float(snapshot_value(snapshot, table_row, anchor_col - 3))
        num_quarters_used = int(table_num_quarters) if table_num_quarters is not None else n_used

        window = xy_history[-n_used:]
        x_next = window[-1][1] + 1
        forecast_calc = intercept + (slope * x_next)
        table_forecast = to_float(snapshot_value(snapshot, table_row, anchor_col - 2))
        forecast_value = table_forecast if table_forecast is not None else forecast_calc

        forecast_max = to_float(snapshot_value(snapshot, table_row, anchor_col))
        forecast_min = to_float(snapshot_value(snapshot, table_row, anchor_col - 1))
        if forecast_value is not None and (forecast_max is None or forecast_min is None):
            residuals = [y - (intercept + slope * x) for _, x, y in window]
            if residuals:
                spread = statistics.pstdev(residuals)
            else:
                spread = 0.0
            if forecast_max is None:
                forecast_max = forecast_value + spread
            if forecast_min is None:
                forecast_min = forecast_value - spread

        range_width = (
            forecast_max - forecast_min
            if forecast_max is not None and forecast_min is not None
            else None
        )
        table_actual = to_float(snapshot_value(snapshot, table_row, anchor_col - 4))
        actual_value = table_actual if table_actual is not None else default_actual

        row = {
            "model": metadata.model,
            "ticker": metadata.ticker,
            "model_period": metadata.model_period,
            "model_date": metadata.model_date,
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

        if rows and is_duplicate_regression_row(rows[-1], row):
            continue
        rows.append(row)

    return rows


def append_rows_to_sheet(worksheet: Any, headers: list[str], rows: list[dict[str, Any]]) -> None:
    worksheet.append(headers)
    for row in rows:
        worksheet.append([row.get(header) for header in headers])

    for cell in worksheet[1]:
        cell.font = Font(bold=True)

    worksheet.freeze_panes = "A2"
    worksheet.auto_filter.ref = worksheet.dimensions

    for col_index, header in enumerate(headers, start=1):
        max_length = len(header)
        for row_index in range(2, worksheet.max_row + 1):
            value = worksheet.cell(row=row_index, column=col_index).value
            if value is None:
                continue
            max_length = max(max_length, len(str(value)))
        width = min(50, max(12, max_length + 2))
        worksheet.column_dimensions[get_column_letter(col_index)].width = width


def write_output_workbook(
    output_path: Path, empirical_rows: list[dict[str, Any]], regression_rows: list[dict[str, Any]]
) -> None:
    workbook = Workbook()
    empirical_sheet = workbook.active
    empirical_sheet.title = "empirical_candidates"
    regression_sheet = workbook.create_sheet("regression_candidates")

    append_rows_to_sheet(empirical_sheet, EMPIRICAL_COLUMNS, empirical_rows)
    append_rows_to_sheet(regression_sheet, REGRESSION_COLUMNS, regression_rows)
    workbook.save(output_path)


def run() -> None:
    if not input_dir.exists():
        raise FileNotFoundError(f"input_dir does not exist: {input_dir}")
    output_dir.mkdir(parents=True, exist_ok=True)

    empirical_rows: list[dict[str, Any]] = []
    regression_rows: list[dict[str, Any]] = []
    processed_files = 0

    app: xw.App | None = None
    try:
        app = xw.App(visible=False, add_book=False)
        app.display_alerts = False
        app.screen_updating = False

        for file_path in sorted(input_dir.iterdir()):
            if not file_path.is_file():
                continue
            if file_path.name.startswith("~"):
                print(f"skipped file {file_path.name}: temporary Excel file")
                continue
            if file_path.suffix.lower() != ".xlsx":
                print(f"skipped file {file_path.name}: not an .xlsx file")
                continue
            if re.match(rf"^{re.escape(input_dir.name)}_param(\.\d+)?$", file_path.stem, flags=re.IGNORECASE):
                print(f"skipped file {file_path.name}: prior generated output workbook")
                continue

            workbook: xw.Book | None = None
            try:
                workbook = app.books.open(str(file_path), update_links=False)
            except Exception as exc:
                print(f"skipped file {file_path.name}: unable to open ({exc})")
                continue

            try:
                metadata = parse_file_metadata(file_path)
                empirical_rows.extend(
                    extract_empirical_rows(workbook=workbook, metadata=metadata, source_file=file_path.name)
                )
                regression_rows.extend(
                    extract_regression_rows(workbook=workbook, metadata=metadata, source_file=file_path.name)
                )
                processed_files += 1
                print(f"processed file: {file_path.name}")
            except Exception as exc:
                print(f"skipped file {file_path.name}: processing error ({exc})")
            finally:
                if workbook is not None:
                    close_workbook_without_saving(workbook)
    finally:
        if app is not None:
            try:
                app.quit()
            except Exception:
                pass

    output_path = build_output_path(input_dir, output_dir)
    write_output_workbook(output_path, empirical_rows, regression_rows)

    print(f"output path: {output_path}")
    print(f"number of files processed: {processed_files}")
    print(f"number of empirical rows: {len(empirical_rows)}")
    print(f"number of regression rows: {len(regression_rows)}")


if __name__ == "__main__":
    run()
