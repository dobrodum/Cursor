#!/usr/bin/env python3
from __future__ import annotations

import re
from dataclasses import dataclass
from datetime import date
from pathlib import Path
from typing import Any, Dict, List, Optional, Sequence, Tuple

import xwings as xw
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

# Configure these two paths before running.
input_dir = Path("/path/to/input")
output_dir = Path("/path/to/output")

N_QUARTERS = 10
EMPIRICAL_MODEL_SHEET = "Empirical Model"
REGRESSION_MODEL_SHEET = "Regression Model"
EMPIRICAL_OUTPUT_SHEET = "empirical_candidates"
REGRESSION_OUTPUT_SHEET = "regression_candidates"

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

PERIOD_RE = re.compile(r"(Early|Mid|Late)\s*([A-Za-z]{3,9})\s*([12]\d{3})", re.IGNORECASE)

MatrixInfo = Tuple[int, int, List[List[Any]], int, int]


@dataclass(frozen=True)
class FileLabels:
    model: str
    ticker: str
    model_period: str
    model_date: str


def normalize_text(value: Any) -> str:
    if value is None:
        return ""
    text = str(value).strip().lower()
    text = re.sub(r"[^a-z0-9]+", " ", text)
    return " ".join(text.split())


def to_2d(values: Any) -> List[List[Any]]:
    if values is None:
        return []
    if not isinstance(values, list):
        return [[values]]
    if values and not isinstance(values[0], list):
        return [values]
    return values


def to_float(value: Any) -> Optional[float]:
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
            if not text:
                return None
            try:
                return float(text) / 100.0
            except ValueError:
                return None
        try:
            return float(text)
        except ValueError:
            return None
    return None


def to_int(value: Any, default: int) -> int:
    number = to_float(value)
    if number is None:
        return default
    return int(round(number))


def is_blank(value: Any) -> bool:
    return value is None or (isinstance(value, str) and value.strip() == "")


def compute_range_width(forecast_max: Any, forecast_min: Any) -> Optional[float]:
    max_value = to_float(forecast_max)
    min_value = to_float(forecast_min)
    if max_value is None or min_value is None:
        return None
    return max_value - min_value


def get_used_matrix(sheet: xw.Sheet) -> MatrixInfo:
    used = sheet.used_range
    base_row = used.row
    base_col = used.column
    values = to_2d(used.value)
    if not values:
        return base_row, base_col, [], base_row, base_col
    row_count = len(values)
    col_count = max((len(row) for row in values), default=1)
    end_row = base_row + row_count - 1
    end_col = base_col + col_count - 1
    return base_row, base_col, values, end_row, end_col


def matrix_value(matrix_info: MatrixInfo, row: int, col: int) -> Any:
    base_row, base_col, matrix, end_row, end_col = matrix_info
    if row < base_row or row > end_row or col < base_col or col > end_col:
        return None
    row_values = matrix[row - base_row]
    col_index = col - base_col
    if col_index < 0 or col_index >= len(row_values):
        return None
    return row_values[col_index]


def get_cell_value(sheet: xw.Sheet, matrix_info: MatrixInfo, row: int, col: int) -> Any:
    base_row, base_col, _, end_row, end_col = matrix_info
    if base_row <= row <= end_row and base_col <= col <= end_col:
        return matrix_value(matrix_info, row, col)
    return sheet.cells(row, col).value


def build_header_map(matrix_info: MatrixInfo, header_row: int, left_col: int, right_col: int) -> Dict[str, int]:
    headers: Dict[str, int] = {}
    for col in range(left_col, right_col + 1):
        header = normalize_text(matrix_value(matrix_info, header_row, col))
        if header and header not in headers:
            headers[header] = col
    return headers


def find_column(header_map: Dict[str, int], aliases: Sequence[str], fallback: Optional[int] = None) -> Optional[int]:
    normalized_aliases = [normalize_text(alias) for alias in aliases]

    for alias in normalized_aliases:
        if alias in header_map:
            return header_map[alias]

    for header, col in header_map.items():
        for alias in normalized_aliases:
            if alias and alias in header:
                return col

    return fallback


def find_max_anchor(matrix_info: MatrixInfo) -> Optional[Tuple[int, int]]:
    base_row, base_col, matrix, _, _ = matrix_info
    fallback: Optional[Tuple[int, int]] = None

    for row_index, row_values in enumerate(matrix):
        for col_index, value in enumerate(row_values):
            if normalize_text(value) != "max":
                continue

            row = base_row + row_index
            col = base_col + col_index

            if fallback is None:
                fallback = (row, col)

            right_header = normalize_text(matrix_value(matrix_info, row, col + 1))
            if right_header == "min":
                return row, col

    return fallback


def parse_file_labels(file_path: Path) -> FileLabels:
    stem = file_path.stem
    parts = [part.strip() for part in re.split(r"\s*-\s*", stem) if part.strip()]

    ticker = ""
    if len(parts) >= 2:
        ticker = re.sub(r"[^A-Za-z0-9]", "", parts[1]).upper()
    if not ticker:
        for part in parts:
            token = re.sub(r"[^A-Za-z0-9]", "", part).upper()
            if re.fullmatch(r"[A-Z]{2,8}", token):
                ticker = token
                break
    if not ticker:
        ticker = "UNKNOWN"

    period_source = parts[2] if len(parts) >= 3 else stem
    period_source = period_source.split("_")[0]
    match = PERIOD_RE.search(period_source) or PERIOD_RE.search(stem)

    model_period = "unknown_period"
    model_date = ""

    if match:
        phase_raw, month_raw, year_raw = match.groups()
        phase = phase_raw.capitalize()
        month_key = month_raw[:3].lower()
        month_number = MONTH_TO_NUMBER.get(month_key)
        if month_number is not None:
            day = {"Early": 5, "Mid": 15, "Late": 25}[phase]
            model_period = f"{phase}{month_key.capitalize()}_{year_raw}"
            model_date = date(int(year_raw), month_number, day).isoformat()

    model = f"{ticker}_{model_period}"
    return FileLabels(model=model, ticker=ticker, model_period=model_period, model_date=model_date)


def close_workbook_safe(workbook: xw.Book) -> None:
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
        # If all close methods fail there is nothing else to do.
        return


def build_empirical_rows(sheet: xw.Sheet, labels: FileLabels, source_file: str) -> List[Dict[str, Any]]:
    matrix_info = get_used_matrix(sheet)
    anchor = find_max_anchor(matrix_info)
    if anchor is None:
        print(f"skipped empirical in {source_file}: max anchor not found")
        return []

    anchor_row, anchor_col = anchor
    _, _, _, _, used_end_col = matrix_info

    header_map = build_header_map(
        matrix_info=matrix_info,
        header_row=anchor_row,
        left_col=max(1, anchor_col - 30),
        right_col=anchor_col + 30,
    )

    num_quarters_col = find_column(header_map, ["num_quarters_used", "num quarters used", "num quarters", "quarters"])
    last_quarter_col = find_column(header_map, ["last_quarter_used", "last quarter used", "last quarter"])
    forecast_value_col = find_column(
        header_map,
        ["estimated total sold", "forecast value", "tot fcst", "total forecast", "forecast total"],
        fallback=anchor_col - 1,
    )
    actual_value_col = find_column(header_map, ["actual value", "reported sales", "actual", "reported"])
    forecast_max_col = find_column(header_map, ["forecast max", "max"], fallback=anchor_col)
    forecast_min_col = find_column(header_map, ["forecast min", "min"], fallback=anchor_col + 1)
    quarterly_sales_col = find_column(
        header_map,
        ["quarterly sales", "quarterly_sale", "sales"],
        fallback=anchor_col - 11,
    )
    reported_sales_col = find_column(
        header_map,
        ["reported sales", "reported_sale", "reported"],
        fallback=anchor_col - 7,
    )
    growth_rate_col = find_column(header_map, ["growth rate pct", "growth rate", "growth pct", "growth"])
    sales_captured_col = find_column(
        header_map,
        ["sales captured in db pct", "sales captured in db", "captured in db", "db captured"],
    )

    quarterly_sales_col = max(1, quarterly_sales_col or 1)
    reported_sales_col = max(1, reported_sales_col or 1)
    temp_avg_penetration_col = used_end_col + 5

    wrote_formula = False
    for index in range(1, N_QUARTERS + 1):
        row = anchor_row + index
        num_quarters = to_int(
            get_cell_value(sheet, matrix_info, row, num_quarters_col) if num_quarters_col else index,
            default=index,
        )
        num_quarters = max(1, min(num_quarters, max(1, anchor_row - 1)))
        start_row = max(1, anchor_row - num_quarters)
        end_row = anchor_row - 1
        if start_row > end_row:
            continue

        formula = (
            "=IFERROR("
            "AVERAGE("
            f"IFERROR((R{start_row}C{reported_sales_col}:R{end_row}C{reported_sales_col})/"
            f"(R{start_row}C{quarterly_sales_col}:R{end_row}C{quarterly_sales_col}),\"\")),"
            "\"\")"
        )
        sheet.cells(row, temp_avg_penetration_col).formula2 = formula
        wrote_formula = True

    if wrote_formula:
        sheet.book.app.calculate()

    rows: List[Dict[str, Any]] = []
    for index in range(1, N_QUARTERS + 1):
        row = anchor_row + index
        num_quarters_used = to_int(
            get_cell_value(sheet, matrix_info, row, num_quarters_col) if num_quarters_col else index,
            default=index,
        )
        last_quarter_used = get_cell_value(sheet, matrix_info, row, last_quarter_col) if last_quarter_col else None
        forecast_value = get_cell_value(sheet, matrix_info, row, forecast_value_col) if forecast_value_col else None
        actual_value = get_cell_value(sheet, matrix_info, row, actual_value_col) if actual_value_col else None
        forecast_max = get_cell_value(sheet, matrix_info, row, forecast_max_col) if forecast_max_col else None
        forecast_min = get_cell_value(sheet, matrix_info, row, forecast_min_col) if forecast_min_col else None
        avg_penetration_pct = sheet.cells(row, temp_avg_penetration_col).value
        quarterly_sales = get_cell_value(sheet, matrix_info, row, quarterly_sales_col)
        reported_sales = get_cell_value(sheet, matrix_info, row, reported_sales_col)
        growth_rate_pct = get_cell_value(sheet, matrix_info, row, growth_rate_col) if growth_rate_col else None
        sales_captured_in_db_pct = (
            get_cell_value(sheet, matrix_info, row, sales_captured_col) if sales_captured_col else None
        )

        if all(
            is_blank(value)
            for value in (
                forecast_value,
                actual_value,
                forecast_max,
                forecast_min,
                quarterly_sales,
                reported_sales,
            )
        ):
            continue

        rows.append(
            {
                "model": labels.model,
                "ticker": labels.ticker,
                "model_period": labels.model_period,
                "model_date": labels.model_date,
                "method": "empirical",
                "parameter_name": "avg_penetration_pct",
                "parameter_value": avg_penetration_pct,
                "num_quarters_used": num_quarters_used,
                "last_quarter_used": last_quarter_used,
                "forecast_value": forecast_value,
                "actual_value": actual_value,
                "forecast_max": forecast_max,
                "forecast_min": forecast_min,
                "range_width": compute_range_width(forecast_max, forecast_min),
                "avg_penetration_pct": avg_penetration_pct,
                "quarterly_sales": quarterly_sales,
                "reported_sales": reported_sales,
                "growth_rate_pct": growth_rate_pct,
                "sales_captured_in_db_pct": sales_captured_in_db_pct,
                "source_file": source_file,
            }
        )

    return rows


def rows_are_equivalent(previous: Dict[str, Any], current: Dict[str, Any]) -> bool:
    keys = ("num_quarters_used", "forecast_value", "forecast_max", "forecast_min", "intercept", "slope")
    for key in keys:
        previous_value = previous.get(key)
        current_value = current.get(key)

        previous_number = to_float(previous_value)
        current_number = to_float(current_value)

        if previous_number is not None and current_number is not None:
            if abs(previous_number - current_number) > 1e-9:
                return False
            continue

        if normalize_text(previous_value) != normalize_text(current_value):
            return False
    return True


def build_regression_rows(sheet: xw.Sheet, labels: FileLabels, source_file: str) -> List[Dict[str, Any]]:
    matrix_info = get_used_matrix(sheet)
    anchor = find_max_anchor(matrix_info)
    if anchor is None:
        print(f"skipped regression in {source_file}: max anchor not found")
        return []

    anchor_row, anchor_col = anchor
    _, _, _, _, used_end_col = matrix_info

    header_map = build_header_map(
        matrix_info=matrix_info,
        header_row=anchor_row,
        left_col=max(1, anchor_col - 30),
        right_col=anchor_col + 30,
    )

    num_quarters_col = find_column(header_map, ["num_quarters_used", "num quarters used", "num quarters", "quarters"])
    forecast_total_col = find_column(
        header_map,
        ["tot fcst w/o sa", "tot fcst wo sa", "forecast total without sa", "total forecast without sa", "tot fcst"],
        fallback=anchor_col - 1,
    )
    actual_value_col = find_column(header_map, ["actual value", "actual", "reported sales", "reported"])
    forecast_max_col = find_column(header_map, ["forecast max", "max"], fallback=anchor_col)
    forecast_min_col = find_column(header_map, ["forecast min", "min"], fallback=anchor_col + 1)

    x_col = max(1, anchor_col - 11)
    y_col = max(1, anchor_col - 7)
    temp_intercept_col = used_end_col + 5
    temp_slope_col = used_end_col + 6

    wrote_formula = False
    for index in range(1, N_QUARTERS + 1):
        row = anchor_row + index
        num_quarters = to_int(
            get_cell_value(sheet, matrix_info, row, num_quarters_col) if num_quarters_col else index,
            default=index,
        )
        num_quarters = max(1, min(num_quarters, max(1, anchor_row - 1)))
        start_row = max(1, anchor_row - num_quarters)
        end_row = anchor_row - 1

        if start_row > end_row:
            continue

        intercept_formula = (
            f"=IFERROR(INTERCEPT(R{start_row}C{y_col}:R{end_row}C{y_col},"
            f"R{start_row}C{x_col}:R{end_row}C{x_col}),\"\")"
        )
        slope_formula = (
            f"=IFERROR(SLOPE(R{start_row}C{y_col}:R{end_row}C{y_col},"
            f"R{start_row}C{x_col}:R{end_row}C{x_col}),\"\")"
        )

        sheet.cells(row, temp_intercept_col).formula2 = intercept_formula
        sheet.cells(row, temp_slope_col).formula2 = slope_formula
        wrote_formula = True

    if wrote_formula:
        sheet.book.app.calculate()

    rows: List[Dict[str, Any]] = []
    for index in range(1, N_QUARTERS + 1):
        row = anchor_row + index
        num_quarters_used = to_int(
            get_cell_value(sheet, matrix_info, row, num_quarters_col) if num_quarters_col else index,
            default=index,
        )
        forecast_value = get_cell_value(sheet, matrix_info, row, forecast_total_col) if forecast_total_col else None
        actual_value = get_cell_value(sheet, matrix_info, row, actual_value_col) if actual_value_col else None
        forecast_max = get_cell_value(sheet, matrix_info, row, forecast_max_col) if forecast_max_col else None
        forecast_min = get_cell_value(sheet, matrix_info, row, forecast_min_col) if forecast_min_col else None
        intercept = sheet.cells(row, temp_intercept_col).value
        slope = sheet.cells(row, temp_slope_col).value

        if all(is_blank(value) for value in (forecast_value, forecast_max, forecast_min, intercept, slope)):
            continue

        row_data: Dict[str, Any] = {
            "model": labels.model,
            "ticker": labels.ticker,
            "model_period": labels.model_period,
            "model_date": labels.model_date,
            "method": "regression",
            "parameter_name": "num_quarters_used",
            "parameter_value": num_quarters_used,
            "num_quarters_used": num_quarters_used,
            "forecast_value": forecast_value,
            "actual_value": actual_value,
            "forecast_max": forecast_max,
            "forecast_min": forecast_min,
            "range_width": compute_range_width(forecast_max, forecast_min),
            "intercept": intercept,
            "slope": slope,
            "source_file": source_file,
        }

        if rows and rows_are_equivalent(rows[-1], row_data):
            continue
        rows.append(row_data)

    return rows


def add_sheet_with_rows(workbook: Workbook, name: str, columns: Sequence[str], rows: Sequence[Dict[str, Any]]) -> None:
    worksheet = workbook.create_sheet(name)
    worksheet.append(list(columns))

    for row in rows:
        worksheet.append([row.get(column, "") for column in columns])

    for cell in worksheet[1]:
        cell.font = Font(bold=True)

    worksheet.freeze_panes = "A2"
    worksheet.auto_filter.ref = worksheet.dimensions

    for col_index, column_name in enumerate(columns, start=1):
        values = [column_name]
        for row in rows:
            value = row.get(column_name)
            values.append("" if value is None else str(value))
        max_width = max(len(value) for value in values)
        worksheet.column_dimensions[get_column_letter(col_index)].width = min(max(12, max_width + 2), 48)


def build_output_path(input_path: Path, output_path: Path) -> Path:
    folder_name = input_path.resolve().name
    base_name = f"{folder_name}_PARAM"
    candidate = output_path / f"{base_name}.xlsx"
    suffix = 1
    while candidate.exists():
        candidate = output_path / f"{base_name}.{suffix}.xlsx"
        suffix += 1
    return candidate


def write_output_workbook(output_path: Path, empirical_rows: Sequence[Dict[str, Any]], regression_rows: Sequence[Dict[str, Any]]) -> None:
    workbook = Workbook()
    workbook.remove(workbook.active)

    add_sheet_with_rows(workbook, EMPIRICAL_OUTPUT_SHEET, EMPIRICAL_COLUMNS, empirical_rows)
    add_sheet_with_rows(workbook, REGRESSION_OUTPUT_SHEET, REGRESSION_COLUMNS, regression_rows)

    output_path.parent.mkdir(parents=True, exist_ok=True)
    workbook.save(output_path)


def main() -> None:
    input_path = Path(input_dir)
    output_path = Path(output_dir)

    if not input_path.exists() or not input_path.is_dir():
        raise FileNotFoundError(f"input_dir does not exist or is not a folder: {input_path}")

    output_path.mkdir(parents=True, exist_ok=True)
    output_file = build_output_path(input_path, output_path)

    empirical_rows: List[Dict[str, Any]] = []
    regression_rows: List[Dict[str, Any]] = []
    processed_files = 0

    app = xw.App(visible=False, add_book=False)
    app.display_alerts = False
    app.screen_updating = False

    try:
        for file_path in sorted(input_path.iterdir(), key=lambda p: p.name.lower()):
            if not file_path.is_file():
                print(f"skipped {file_path.name}: not a file")
                continue

            if file_path.name.startswith("~"):
                print(f"skipped {file_path.name}: temporary file")
                continue

            if file_path.suffix.lower() != ".xlsx":
                print(f"skipped {file_path.name}: not an .xlsx file")
                continue

            print(f"processing {file_path.name}")
            workbook: Optional[xw.Book] = None
            try:
                workbook = app.books.open(str(file_path), update_links=False)
                labels = parse_file_labels(file_path)

                try:
                    empirical_sheet = workbook.sheets[EMPIRICAL_MODEL_SHEET]
                    empirical_rows.extend(build_empirical_rows(empirical_sheet, labels, file_path.name))
                except Exception as exc:
                    print(f"skipped empirical in {file_path.name}: {exc}")

                try:
                    regression_sheet = workbook.sheets[REGRESSION_MODEL_SHEET]
                    regression_rows.extend(build_regression_rows(regression_sheet, labels, file_path.name))
                except Exception as exc:
                    print(f"skipped regression in {file_path.name}: {exc}")

                processed_files += 1
            except Exception as exc:
                print(f"skipped {file_path.name}: open/process failure ({exc})")
            finally:
                if workbook is not None:
                    close_workbook_safe(workbook)
    finally:
        app.quit()

    write_output_workbook(output_file, empirical_rows, regression_rows)

    print(f"output path: {output_file}")
    print(f"number of files processed: {processed_files}")
    print(f"number of empirical rows: {len(empirical_rows)}")
    print(f"number of regression rows: {len(regression_rows)}")


if __name__ == "__main__":
    main()
