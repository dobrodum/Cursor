#!/usr/bin/env python3
"""Extract empirical and regression model candidates from Excel workbooks."""

from __future__ import annotations

import re
from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path
from typing import Any, Iterable, Optional

import xlwings as xw
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

# -----------------------------
# User-configurable input/output
# -----------------------------
input_dir = r"./input"
output_dir = r"./output"

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

MONTH_LOOKUP = {
    "jan": 1,
    "january": 1,
    "feb": 2,
    "february": 2,
    "mar": 3,
    "march": 3,
    "apr": 4,
    "april": 4,
    "may": 5,
    "jun": 6,
    "june": 6,
    "jul": 7,
    "july": 7,
    "aug": 8,
    "august": 8,
    "sep": 9,
    "sept": 9,
    "september": 9,
    "oct": 10,
    "october": 10,
    "nov": 11,
    "november": 11,
    "dec": 12,
    "december": 12,
}

MONTH_ABBR = {
    1: "Jan",
    2: "Feb",
    3: "Mar",
    4: "Apr",
    5: "May",
    6: "Jun",
    7: "Jul",
    8: "Aug",
    9: "Sep",
    10: "Oct",
    11: "Nov",
    12: "Dec",
}


@dataclass(frozen=True)
class FileLabel:
    model: str
    ticker: str
    model_period: str
    model_date: str


def normalize_text(value: Any) -> str:
    if value is None:
        return ""
    text = str(value).strip().lower()
    return re.sub(r"[^a-z0-9]+", " ", text).strip()


def to_float(value: Any) -> Optional[float]:
    if value is None:
        return None
    if isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        return float(value)

    text = str(value).strip()
    if not text:
        return None
    text = text.replace(",", "")
    if text.endswith("%"):
        text = text[:-1]
        try:
            return float(text) / 100.0
        except ValueError:
            return None
    try:
        return float(text)
    except ValueError:
        return None


def to_int(value: Any) -> Optional[int]:
    numeric = to_float(value)
    if numeric is None:
        return None
    try:
        return int(round(numeric))
    except (TypeError, ValueError):
        return None


def clean_value(value: Any) -> Any:
    if value is None:
        return None
    if isinstance(value, str):
        trimmed = value.strip()
        return trimmed if trimmed else None
    if isinstance(value, datetime):
        return value.date().isoformat()
    if isinstance(value, date):
        return value.isoformat()
    return value


def safe_subtract(left: Optional[float], right: Optional[float]) -> Optional[float]:
    if left is None or right is None:
        return None
    return left - right


def output_path_for_run(in_dir: Path, out_dir: Path) -> Path:
    base_name = f"{in_dir.name}_PARAM"
    candidate = out_dir / f"{base_name}.xlsx"
    suffix = 1
    while candidate.exists():
        candidate = out_dir / f"{base_name}.{suffix}.xlsx"
        suffix += 1
    return candidate


def parse_file_label(file_path: Path) -> Optional[FileLabel]:
    parts = [part.strip() for part in file_path.stem.split(" - ")]
    if len(parts) < 3:
        return None

    ticker = parts[1].strip().upper()
    model_token = parts[2].split("_", 1)[0].strip()
    match = re.match(r"^(Early|Mid|Late)([A-Za-z]+)(\d{4})$", model_token, re.IGNORECASE)
    if not match:
        return None

    period_prefix = match.group(1).capitalize()
    month_token = match.group(2).lower().rstrip(".")
    year = int(match.group(3))

    month_num = MONTH_LOOKUP.get(month_token)
    if month_num is None and len(month_token) >= 3:
        month_num = MONTH_LOOKUP.get(month_token[:3])
    if month_num is None:
        return None

    day_map = {"Early": 5, "Mid": 15, "Late": 25}
    day = day_map[period_prefix]

    model_period = f"{period_prefix}{MONTH_ABBR[month_num]}_{year}"
    model_date = date(year, month_num, day).isoformat()
    model = f"{ticker}_{model_period}"

    return FileLabel(
        model=model,
        ticker=ticker,
        model_period=model_period,
        model_date=model_date,
    )


def read_column_values(sheet: xw.Sheet, start_row: int, end_row: int, col: Optional[int]) -> list[Any]:
    row_count = end_row - start_row + 1
    if col is None or col < 1:
        return [None] * row_count

    values = sheet.range((start_row, col), (end_row, col)).value
    if row_count == 1:
        return [values]
    if not isinstance(values, list):
        values = [values]

    if len(values) < row_count:
        values.extend([None] * (row_count - len(values)))
    return values[:row_count]


def find_anchor(sheet: xw.Sheet, target: str = "max") -> Optional[tuple[int, int]]:
    used = sheet.used_range
    values = used.value
    if values is None:
        return None

    if not isinstance(values, list):
        values = [[values]]
    elif values and not isinstance(values[0], list):
        values = [values]

    target_norm = target.strip().lower()
    start_row = used.row
    start_col = used.column

    for r_idx, row_values in enumerate(values):
        if not isinstance(row_values, list):
            row_values = [row_values]
        for c_idx, cell_value in enumerate(row_values):
            if isinstance(cell_value, str) and cell_value.strip().lower() == target_norm:
                return start_row + r_idx, start_col + c_idx
    return None


def scan_headers(sheet: xw.Sheet, header_row: int, anchor_col: int, window: int = 30) -> list[tuple[str, int]]:
    start_col = max(1, anchor_col - window)
    end_col = max(start_col, anchor_col + window)
    values = sheet.range((header_row, start_col), (header_row, end_col)).value
    if not isinstance(values, list):
        values = [values]
    headers: list[tuple[str, int]] = []
    for idx, value in enumerate(values):
        norm = normalize_text(value)
        if norm:
            headers.append((norm, start_col + idx))
    return headers


def find_column(headers: Iterable[tuple[str, int]], aliases: Iterable[str]) -> Optional[int]:
    for alias in aliases:
        alias_tokens = normalize_text(alias).split()
        if not alias_tokens:
            continue
        for header_text, col in headers:
            if all(token in header_text for token in alias_tokens):
                return col
    return None


def set_formula2(cell: xw.Range, formula: str) -> None:
    try:
        cell.formula2 = formula
    except Exception:
        cell.formula = formula


def close_without_saving(workbook: xw.Book) -> None:
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
        pass


def process_empirical_sheet(workbook: xw.Book, label: FileLabel, source_file: str) -> list[dict[str, Any]]:
    try:
        sheet = workbook.sheets["Empirical Model"]
    except Exception:
        print(f"skipped {source_file}: missing sheet 'Empirical Model'")
        return []

    anchor = find_anchor(sheet, "max")
    if not anchor:
        print(f"skipped {source_file}: 'max' anchor not found on Empirical Model")
        return []

    anchor_row, anchor_col = anchor
    header_pairs = scan_headers(sheet, anchor_row, anchor_col)
    start_row = anchor_row + 1
    end_row = start_row + N_QUARTERS - 1

    num_quarters_col = find_column(header_pairs, ["num quarters used", "quarters used", "num qtr", "n quarters"])
    last_quarter_col = find_column(header_pairs, ["last quarter used", "last qtr"])
    estimated_total_col = find_column(
        header_pairs,
        ["estimated total sold", "est total sold", "forecast value", "estimated total", "total sold"],
    )
    reported_sales_col = find_column(header_pairs, ["reported sales", "actual sales", "actual value", "actual"])
    quarterly_sales_col = find_column(header_pairs, ["quarterly sales", "qtr sales", "quarter sales"])
    growth_rate_col = find_column(header_pairs, ["growth rate", "growth pct", "growth percent"])
    sales_captured_col = find_column(
        header_pairs,
        ["sales captured in db", "captured in db", "db captured", "sales in db"],
    )
    avg_penetration_existing_col = find_column(header_pairs, ["avg penetration", "average penetration"])

    # Anchor-based defaults if headers are absent.
    num_quarters_col = num_quarters_col or max(1, anchor_col - 8)
    last_quarter_col = last_quarter_col or max(1, anchor_col - 9)
    estimated_total_col = estimated_total_col or max(1, anchor_col - 2)
    reported_sales_col = reported_sales_col or max(1, anchor_col - 3)
    quarterly_sales_col = quarterly_sales_col or max(1, anchor_col - 4)
    growth_rate_col = growth_rate_col or (anchor_col + 2)
    sales_captured_col = sales_captured_col or (anchor_col + 3)
    forecast_max_col = anchor_col
    forecast_min_col = find_column(header_pairs, ["forecast min", "min"]) or (anchor_col + 1)

    helper_col = anchor_col + 30
    penetration_start_col = max(1, anchor_col - 11)
    penetration_end_col = max(penetration_start_col, anchor_col - 2)

    formulas_written = False
    for row in range(start_row, end_row + 1):
        formula = f'=IFERROR(AVERAGE(RC{penetration_start_col}:RC{penetration_end_col}),"")'
        set_formula2(sheet.cells(row, helper_col), formula)
        formulas_written = True

    if formulas_written:
        workbook.app.calculate()

    num_quarters_vals = read_column_values(sheet, start_row, end_row, num_quarters_col)
    last_quarter_vals = read_column_values(sheet, start_row, end_row, last_quarter_col)
    estimated_total_vals = read_column_values(sheet, start_row, end_row, estimated_total_col)
    reported_sales_vals = read_column_values(sheet, start_row, end_row, reported_sales_col)
    quarterly_sales_vals = read_column_values(sheet, start_row, end_row, quarterly_sales_col)
    growth_rate_vals = read_column_values(sheet, start_row, end_row, growth_rate_col)
    sales_captured_vals = read_column_values(sheet, start_row, end_row, sales_captured_col)
    forecast_max_vals = read_column_values(sheet, start_row, end_row, forecast_max_col)
    forecast_min_vals = read_column_values(sheet, start_row, end_row, forecast_min_col)
    avg_penetration_formula_vals = read_column_values(sheet, start_row, end_row, helper_col)
    avg_penetration_existing_vals = read_column_values(sheet, start_row, end_row, avg_penetration_existing_col)

    rows: list[dict[str, Any]] = []
    for idx in range(N_QUARTERS):
        num_quarters = to_int(num_quarters_vals[idx]) or (idx + 1)
        forecast_value = to_float(estimated_total_vals[idx])
        actual_value = to_float(reported_sales_vals[idx])
        forecast_max = to_float(forecast_max_vals[idx])
        forecast_min = to_float(forecast_min_vals[idx])
        avg_pen_formula = to_float(avg_penetration_formula_vals[idx])
        avg_pen_existing = to_float(avg_penetration_existing_vals[idx])
        avg_penetration = avg_pen_formula if avg_pen_formula is not None else avg_pen_existing
        quarterly_sales = to_float(quarterly_sales_vals[idx])
        reported_sales = to_float(reported_sales_vals[idx])
        growth_rate = to_float(growth_rate_vals[idx])
        sales_captured = to_float(sales_captured_vals[idx])
        last_quarter_used = clean_value(last_quarter_vals[idx])

        if all(
            value is None
            for value in (
                forecast_value,
                actual_value,
                forecast_max,
                forecast_min,
                avg_penetration,
                quarterly_sales,
                reported_sales,
            )
        ):
            continue

        rows.append(
            {
                "model": label.model,
                "ticker": label.ticker,
                "model_period": label.model_period,
                "model_date": label.model_date,
                "method": "empirical",
                "parameter_name": "avg_penetration_pct",
                "parameter_value": avg_penetration,
                "num_quarters_used": num_quarters,
                "last_quarter_used": last_quarter_used,
                "forecast_value": forecast_value,
                "actual_value": actual_value,
                "forecast_max": forecast_max,
                "forecast_min": forecast_min,
                "range_width": safe_subtract(forecast_max, forecast_min),
                "avg_penetration_pct": avg_penetration,
                "quarterly_sales": quarterly_sales,
                "reported_sales": reported_sales,
                "growth_rate_pct": growth_rate,
                "sales_captured_in_db_pct": sales_captured,
                "source_file": source_file,
            }
        )
    return rows


def _comparable_number(value: Any) -> Any:
    numeric = to_float(value)
    if numeric is None:
        return clean_value(value)
    return round(numeric, 10)


def is_regression_duplicate(prev_row: dict[str, Any], current_row: dict[str, Any]) -> bool:
    keys = ("num_quarters_used", "forecast_value", "forecast_max", "forecast_min", "intercept", "slope")
    return all(_comparable_number(prev_row.get(key)) == _comparable_number(current_row.get(key)) for key in keys)


def process_regression_sheet(workbook: xw.Book, label: FileLabel, source_file: str) -> list[dict[str, Any]]:
    try:
        sheet = workbook.sheets["Regression Model"]
    except Exception:
        print(f"skipped {source_file}: missing sheet 'Regression Model'")
        return []

    anchor = find_anchor(sheet, "max")
    if not anchor:
        print(f"skipped {source_file}: 'max' anchor not found on Regression Model")
        return []

    anchor_row, anchor_col = anchor
    header_pairs = scan_headers(sheet, anchor_row, anchor_col)
    start_row = anchor_row + 1
    end_row = start_row + N_QUARTERS - 1

    y_col = anchor_col - 7
    x_col = anchor_col - 11

    num_quarters_col = find_column(header_pairs, ["num quarters used", "quarters used", "num qtr", "n quarters"])
    forecast_total_col = find_column(
        header_pairs,
        ["tot fcst w o sa", "total forecast without sa", "tot fcst wo sa", "tot fcst w sa"],
    )
    actual_value_col = find_column(header_pairs, ["actual value", "actual sales", "actual"])
    forecast_max_col = anchor_col
    forecast_min_col = find_column(header_pairs, ["forecast min", "min"]) or (anchor_col + 1)

    # Anchor-based defaults.
    num_quarters_col = num_quarters_col or max(1, anchor_col - 8)
    forecast_total_col = forecast_total_col or max(1, anchor_col - 2)

    helper_intercept_col = anchor_col + 30
    helper_slope_col = anchor_col + 31

    num_quarters_vals = read_column_values(sheet, start_row, end_row, num_quarters_col)
    forecast_total_vals = read_column_values(sheet, start_row, end_row, forecast_total_col)
    actual_value_vals = read_column_values(sheet, start_row, end_row, actual_value_col)
    forecast_max_vals = read_column_values(sheet, start_row, end_row, forecast_max_col)
    forecast_min_vals = read_column_values(sheet, start_row, end_row, forecast_min_col)

    formulas_written = False
    for idx in range(N_QUARTERS):
        row = start_row + idx
        num_quarters = to_int(num_quarters_vals[idx]) or (idx + 1)
        if num_quarters < 2:
            continue

        y_end = anchor_row - 1
        x_end = anchor_row - 1
        y_start = y_end - num_quarters + 1
        x_start = x_end - num_quarters + 1

        if min(y_start, x_start, x_col, y_col) < 1:
            continue

        intercept_formula = (
            f'=IFERROR(INTERCEPT(R{y_start}C{y_col}:R{y_end}C{y_col},'
            f'R{x_start}C{x_col}:R{x_end}C{x_col}),"")'
        )
        slope_formula = (
            f'=IFERROR(SLOPE(R{y_start}C{y_col}:R{y_end}C{y_col},'
            f'R{x_start}C{x_col}:R{x_end}C{x_col}),"")'
        )
        set_formula2(sheet.cells(row, helper_intercept_col), intercept_formula)
        set_formula2(sheet.cells(row, helper_slope_col), slope_formula)
        formulas_written = True

    if formulas_written:
        workbook.app.calculate()

    intercept_vals = read_column_values(sheet, start_row, end_row, helper_intercept_col)
    slope_vals = read_column_values(sheet, start_row, end_row, helper_slope_col)

    rows: list[dict[str, Any]] = []
    for idx in range(N_QUARTERS):
        num_quarters = to_int(num_quarters_vals[idx]) or (idx + 1)
        forecast_value = to_float(forecast_total_vals[idx])
        actual_value = to_float(actual_value_vals[idx])
        forecast_max = to_float(forecast_max_vals[idx])
        forecast_min = to_float(forecast_min_vals[idx])
        intercept = to_float(intercept_vals[idx])
        slope = to_float(slope_vals[idx])

        if all(value is None for value in (forecast_value, forecast_max, forecast_min, intercept, slope)):
            continue

        row_data = {
            "model": label.model,
            "ticker": label.ticker,
            "model_period": label.model_period,
            "model_date": label.model_date,
            "method": "regression",
            "parameter_name": "num_quarters_used",
            "parameter_value": num_quarters,
            "num_quarters_used": num_quarters,
            "forecast_value": forecast_value,
            "actual_value": actual_value,
            "forecast_max": forecast_max,
            "forecast_min": forecast_min,
            "range_width": safe_subtract(forecast_max, forecast_min),
            "intercept": intercept,
            "slope": slope,
            "source_file": source_file,
        }

        if rows and is_regression_duplicate(rows[-1], row_data):
            continue

        rows.append(row_data)

    return rows


def write_sheet(workbook: Workbook, name: str, columns: list[str], rows: list[dict[str, Any]]) -> None:
    ws = workbook.create_sheet(title=name)
    ws.append(columns)

    for record in rows:
        ws.append([clean_value(record.get(column)) for column in columns])

    for cell in ws[1]:
        cell.font = Font(bold=True)

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(columns))}{max(1, ws.max_row)}"

    for col_idx, header in enumerate(columns, start=1):
        width = len(header) + 2
        for row_idx in range(2, ws.max_row + 1):
            value = ws.cell(row=row_idx, column=col_idx).value
            if value is None:
                continue
            width = max(width, len(str(value)) + 2)
        ws.column_dimensions[get_column_letter(col_idx)].width = min(max(width, 12), 48)


def write_output_workbook(output_path: Path, empirical_rows: list[dict[str, Any]], regression_rows: list[dict[str, Any]]) -> None:
    workbook = Workbook()
    workbook.remove(workbook.active)
    write_sheet(workbook, "empirical_candidates", EMPIRICAL_COLUMNS, empirical_rows)
    write_sheet(workbook, "regression_candidates", REGRESSION_COLUMNS, regression_rows)
    workbook.save(output_path)


def process_all_files(in_dir: Path) -> tuple[int, list[dict[str, Any]], list[dict[str, Any]]]:
    empirical_rows: list[dict[str, Any]] = []
    regression_rows: list[dict[str, Any]] = []
    processed_files = 0

    app = xw.App(visible=False, add_book=False)
    app.display_alerts = False
    app.screen_updating = False

    try:
        for file_path in sorted(in_dir.iterdir(), key=lambda p: p.name.lower()):
            if not file_path.is_file():
                print(f"skipped {file_path.name}: not a file")
                continue
            if file_path.name.startswith("~"):
                print(f"skipped {file_path.name}: temp file")
                continue
            if file_path.suffix.lower() != ".xlsx":
                print(f"skipped {file_path.name}: not .xlsx")
                continue

            file_label = parse_file_label(file_path)
            if not file_label:
                print(f"skipped {file_path.name}: filename pattern not recognized")
                continue

            workbook: Optional[xw.Book] = None
            try:
                workbook = app.books.open(str(file_path), update_links=False)
                empirical_rows.extend(process_empirical_sheet(workbook, file_label, file_path.name))
                regression_rows.extend(process_regression_sheet(workbook, file_label, file_path.name))
                processed_files += 1
                print(f"processed {file_path.name}")
            except Exception as exc:
                print(f"skipped {file_path.name}: {exc}")
            finally:
                if workbook is not None:
                    close_without_saving(workbook)
    finally:
        app.quit()

    return processed_files, empirical_rows, regression_rows


def main() -> None:
    in_dir = Path(input_dir).expanduser().resolve()
    out_dir = Path(output_dir).expanduser().resolve()

    if not in_dir.exists() or not in_dir.is_dir():
        print(f"input_dir not found or not a directory: {in_dir}")
        return

    out_dir.mkdir(parents=True, exist_ok=True)
    output_path = output_path_for_run(in_dir, out_dir)

    processed_files, empirical_rows, regression_rows = process_all_files(in_dir)
    write_output_workbook(output_path, empirical_rows, regression_rows)

    print(f"output_path: {output_path}")
    print(f"files_processed: {processed_files}")
    print(f"empirical_rows: {len(empirical_rows)}")
    print(f"regression_rows: {len(regression_rows)}")


if __name__ == "__main__":
    main()
