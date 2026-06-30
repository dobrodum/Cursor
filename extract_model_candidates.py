#!/usr/bin/env python3
from __future__ import annotations

import datetime as dt
import re
from pathlib import Path
from typing import Any

import xlwings as xw
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

# Configure these two paths for your environment.
input_dir = Path("/path/to/input")
output_dir = Path("/path/to/output")

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

FILE_LABEL_PATTERN = re.compile(
    r"-\s*([A-Za-z0-9]+)\s*-\s*(Early|Mid|Late)\s*([A-Za-z]{3,9})\s*(\d{4})",
    re.IGNORECASE,
)


def normalize_label(value: Any) -> str:
    if value is None:
        return ""
    return re.sub(r"[^a-z0-9]+", " ", str(value).strip().lower()).strip()


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
    if not isinstance(values, list):
        return [values]
    if not values:
        return []
    if isinstance(values[0], list):
        return [row[0] if row else None for row in values]
    return values


def is_blank(value: Any) -> bool:
    if value is None:
        return True
    if isinstance(value, str):
        stripped = value.strip()
        return stripped == "" or stripped.startswith("#")
    return False


def to_float(value: Any) -> float | None:
    if value is None:
        return None
    if isinstance(value, (int, float)):
        return float(value)
    if isinstance(value, str):
        stripped = value.strip()
        if stripped == "" or stripped.startswith("#"):
            return None
        is_percent = "%" in stripped
        cleaned = stripped.replace(",", "").replace("%", "")
        try:
            numeric = float(cleaned)
        except ValueError:
            return None
        return numeric / 100.0 if is_percent else numeric
    return None


def subtract_if_numeric(max_value: Any, min_value: Any) -> float | None:
    max_float = to_float(max_value)
    min_float = to_float(min_value)
    if max_float is None or min_float is None:
        return None
    return max_float - min_float


def safe_close_workbook(wb: xw.Book) -> None:
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


def get_sheet(wb: xw.Book, sheet_name: str) -> xw.Sheet | None:
    try:
        return wb.sheets[sheet_name]
    except Exception:
        return None


def matrix_value(
    matrix: list[list[Any]],
    matrix_start_row: int,
    matrix_start_col: int,
    row: int,
    col: int | None,
) -> Any:
    if col is None:
        return None
    row_idx = row - matrix_start_row
    col_idx = col - matrix_start_col
    if row_idx < 0 or col_idx < 0 or row_idx >= len(matrix):
        return None
    row_values = matrix[row_idx]
    if col_idx >= len(row_values):
        return None
    return row_values[col_idx]


def find_max_anchor(
    matrix: list[list[Any]], matrix_start_row: int, matrix_start_col: int
) -> tuple[int, int] | None:
    candidates: list[tuple[int, int, int]] = []
    for row_idx, row_values in enumerate(matrix):
        for col_idx, cell_value in enumerate(row_values):
            if normalize_label(cell_value) != "max":
                continue
            near_values = row_values[max(0, col_idx - 4) : col_idx + 5]
            has_nearby_min = any(normalize_label(v) == "min" for v in near_values)
            score = 1 if has_nearby_min else 0
            candidates.append(
                (score, matrix_start_row + row_idx, matrix_start_col + col_idx)
            )
    if not candidates:
        return None
    candidates.sort(key=lambda item: (-item[0], item[1], item[2]))
    return candidates[0][1], candidates[0][2]


def find_col_from_header(
    header_row: list[Any],
    header_start_col: int,
    includes: list[str],
    excludes: list[str] | None = None,
) -> int | None:
    excludes = excludes or []
    include_tokens = [normalize_label(token) for token in includes]
    exclude_tokens = [normalize_label(token) for token in excludes]
    for idx, raw_value in enumerate(header_row):
        label = normalize_label(raw_value)
        if not label:
            continue
        if any(token in label for token in include_tokens):
            if any(token in label for token in exclude_tokens):
                continue
            return header_start_col + idx
    return None


def next_output_path(input_path: Path, output_path: Path) -> Path:
    stem = f"{input_path.name}_PARAM"
    candidate = output_path / f"{stem}.xlsx"
    if not candidate.exists():
        return candidate
    counter = 1
    while True:
        candidate = output_path / f"{stem}.{counter}.xlsx"
        if not candidate.exists():
            return candidate
        counter += 1


def parse_file_labels(file_name: str) -> dict[str, str]:
    match = FILE_LABEL_PATTERN.search(Path(file_name).stem)
    if match is None:
        raise ValueError("filename does not match expected ticker/period pattern")

    ticker = match.group(1).upper()
    cadence = match.group(2).title()
    month_token = match.group(3)[:3].title()
    year = int(match.group(4))
    month_num = dt.datetime.strptime(month_token, "%b").month
    day_map = {"Early": 5, "Mid": 15, "Late": 25}
    model_day = day_map[cadence]

    model_period = f"{cadence}{month_token}_{year}"
    model_date = dt.date(year, month_num, model_day).isoformat()
    model = f"{ticker}_{model_period}"

    return {
        "ticker": ticker,
        "model_period": model_period,
        "model_date": model_date,
        "model": model,
    }


def convert_error_to_none(value: Any) -> Any:
    if isinstance(value, str) and value.strip().startswith("#"):
        return None
    return value


def extract_empirical_rows(
    wb: xw.Book, labels: dict[str, str], source_file: str
) -> list[dict[str, Any]]:
    sheet = get_sheet(wb, "Empirical Model")
    if sheet is None:
        print("  - skipped empirical: sheet 'Empirical Model' not found")
        return []

    used = sheet.used_range
    matrix = normalize_2d(used.value)
    if not matrix:
        return []

    start_row = used.row
    start_col = used.column
    end_row = used.last_cell.row
    end_col = used.last_cell.column

    anchor = find_max_anchor(matrix, start_row, start_col)
    if anchor is None:
        print("  - skipped empirical: 'max' anchor not found")
        return []

    anchor_row, anchor_col = anchor
    header_idx = anchor_row - start_row
    if header_idx < 0 or header_idx >= len(matrix):
        return []
    header_row = matrix[header_idx]

    num_quarters_col = find_col_from_header(
        header_row,
        start_col,
        ["num quarters used", "num quarters", "quarters used", "n qtr", "n quarters"],
    )
    last_quarter_col = find_col_from_header(
        header_row,
        start_col,
        ["last quarter used", "last quarter", "last qtr"],
    )
    forecast_value_col = find_col_from_header(
        header_row,
        start_col,
        ["estimated total sold", "est total sold", "forecast value", "tot fcst"],
    )
    actual_value_col = find_col_from_header(
        header_row,
        start_col,
        ["actual value", "actual sales", "reported sales", "actual"],
    )
    forecast_min_col = find_col_from_header(header_row, start_col, ["min"])
    avg_pen_col = find_col_from_header(
        header_row,
        start_col,
        ["avg penetration", "average penetration", "avg pen"],
    )
    quarterly_sales_col = find_col_from_header(
        header_row,
        start_col,
        ["quarterly sales", "quarter sales", "qtr sales"],
    )
    reported_sales_col = find_col_from_header(
        header_row,
        start_col,
        ["reported sales", "reported"],
    )
    growth_rate_col = find_col_from_header(
        header_row,
        start_col,
        ["growth rate", "growth %"],
    )
    captured_col = find_col_from_header(
        header_row,
        start_col,
        ["sales captured in db", "captured in db", "sales captured", "db %"],
    )
    penetration_history_col = find_col_from_header(
        header_row,
        start_col,
        ["penetration"],
        excludes=["avg", "average"],
    )
    if penetration_history_col is None:
        penetration_history_col = avg_pen_col if avg_pen_col is not None else anchor_col

    max_rows = min(10, max(0, end_row - anchor_row))
    if max_rows == 0:
        return []

    helper_col = end_col + 2
    for idx in range(max_rows):
        out_row = anchor_row + 1 + idx
        data_start_row = anchor_row + 1
        data_end_row = anchor_row + 1 + idx
        sheet.range((out_row, helper_col)).formula2 = (
            f"=AVERAGE(R{data_start_row}C{penetration_history_col}:"
            f"R{data_end_row}C{penetration_history_col})"
        )
    wb.app.calculate()
    avg_pen_values = normalize_1d(
        sheet.range((anchor_row + 1, helper_col), (anchor_row + max_rows, helper_col)).value
    )

    rows: list[dict[str, Any]] = []
    for idx in range(max_rows):
        row_num = anchor_row + 1 + idx
        num_quarters_used = matrix_value(
            matrix, start_row, start_col, row_num, num_quarters_col
        )
        if is_blank(num_quarters_used):
            num_quarters_used = idx + 1

        forecast_value = matrix_value(
            matrix, start_row, start_col, row_num, forecast_value_col
        )
        actual_value = matrix_value(matrix, start_row, start_col, row_num, actual_value_col)
        forecast_max = matrix_value(matrix, start_row, start_col, row_num, anchor_col)
        forecast_min = matrix_value(matrix, start_row, start_col, row_num, forecast_min_col)

        if all(is_blank(v) for v in (forecast_value, forecast_max, forecast_min)):
            continue

        avg_penetration_pct = (
            convert_error_to_none(avg_pen_values[idx])
            if idx < len(avg_pen_values)
            else matrix_value(matrix, start_row, start_col, row_num, avg_pen_col)
        )
        if avg_penetration_pct is None:
            avg_penetration_pct = matrix_value(
                matrix, start_row, start_col, row_num, avg_pen_col
            )

        reported_sales = matrix_value(
            matrix, start_row, start_col, row_num, reported_sales_col
        )
        if is_blank(actual_value):
            actual_value = reported_sales

        row = {
            "model": labels["model"],
            "ticker": labels["ticker"],
            "model_period": labels["model_period"],
            "model_date": labels["model_date"],
            "method": "empirical",
            "parameter_name": "avg_penetration_pct",
            "parameter_value": avg_penetration_pct,
            "num_quarters_used": num_quarters_used,
            "last_quarter_used": matrix_value(
                matrix, start_row, start_col, row_num, last_quarter_col
            ),
            "forecast_value": forecast_value,
            "actual_value": actual_value,
            "forecast_max": forecast_max,
            "forecast_min": forecast_min,
            "range_width": subtract_if_numeric(forecast_max, forecast_min),
            "avg_penetration_pct": avg_penetration_pct,
            "quarterly_sales": matrix_value(
                matrix, start_row, start_col, row_num, quarterly_sales_col
            ),
            "reported_sales": reported_sales,
            "growth_rate_pct": matrix_value(
                matrix, start_row, start_col, row_num, growth_rate_col
            ),
            "sales_captured_in_db_pct": matrix_value(
                matrix, start_row, start_col, row_num, captured_col
            ),
            "source_file": source_file,
        }
        rows.append(row)

    return rows


def comparable(value: Any) -> Any:
    value = convert_error_to_none(value)
    numeric = to_float(value)
    if numeric is not None:
        return round(numeric, 10)
    if isinstance(value, str):
        return value.strip()
    return value


def extract_regression_rows(
    wb: xw.Book, labels: dict[str, str], source_file: str
) -> list[dict[str, Any]]:
    sheet = get_sheet(wb, "Regression Model")
    if sheet is None:
        print("  - skipped regression: sheet 'Regression Model' not found")
        return []

    used = sheet.used_range
    matrix = normalize_2d(used.value)
    if not matrix:
        return []

    start_row = used.row
    start_col = used.column
    end_row = used.last_cell.row
    end_col = used.last_cell.column

    anchor = find_max_anchor(matrix, start_row, start_col)
    if anchor is None:
        print("  - skipped regression: 'max' anchor not found")
        return []

    anchor_row, anchor_col = anchor
    header_idx = anchor_row - start_row
    if header_idx < 0 or header_idx >= len(matrix):
        return []
    header_row = matrix[header_idx]

    y_col = anchor_col - 7
    x_col = anchor_col - 11
    num_quarters_col = find_col_from_header(
        header_row,
        start_col,
        ["num quarters used", "num quarters", "quarters used", "n qtr", "n quarters"],
    )
    forecast_value_col = find_col_from_header(
        header_row,
        start_col,
        ["tot fcst w o sa", "tot fcst without sa", "forecast total without sa", "tot fcst"],
    )
    actual_value_col = find_col_from_header(
        header_row,
        start_col,
        ["actual value", "actual sales", "reported sales", "actual"],
    )
    forecast_min_col = find_col_from_header(header_row, start_col, ["min"])

    max_rows = min(10, max(0, end_row - anchor_row))
    if max_rows == 0:
        return []

    helper_intercept_col = end_col + 2
    helper_slope_col = end_col + 3
    for idx in range(max_rows):
        out_row = anchor_row + 1 + idx
        data_start_row = anchor_row + 1
        data_end_row = anchor_row + 1 + idx
        y_range = f"R{data_start_row}C{y_col}:R{data_end_row}C{y_col}"
        x_range = f"R{data_start_row}C{x_col}:R{data_end_row}C{x_col}"
        sheet.range((out_row, helper_intercept_col)).formula2 = f"=INTERCEPT({y_range},{x_range})"
        sheet.range((out_row, helper_slope_col)).formula2 = f"=SLOPE({y_range},{x_range})"
    wb.app.calculate()

    intercept_values = normalize_1d(
        sheet.range(
            (anchor_row + 1, helper_intercept_col),
            (anchor_row + max_rows, helper_intercept_col),
        ).value
    )
    slope_values = normalize_1d(
        sheet.range(
            (anchor_row + 1, helper_slope_col),
            (anchor_row + max_rows, helper_slope_col),
        ).value
    )

    rows: list[dict[str, Any]] = []
    previous_signature: tuple[Any, ...] | None = None

    for idx in range(max_rows):
        row_num = anchor_row + 1 + idx
        num_quarters_used = matrix_value(
            matrix, start_row, start_col, row_num, num_quarters_col
        )
        if is_blank(num_quarters_used):
            num_quarters_used = idx + 1

        forecast_value = matrix_value(
            matrix, start_row, start_col, row_num, forecast_value_col
        )
        actual_value = matrix_value(matrix, start_row, start_col, row_num, actual_value_col)
        forecast_max = matrix_value(matrix, start_row, start_col, row_num, anchor_col)
        forecast_min = matrix_value(matrix, start_row, start_col, row_num, forecast_min_col)
        intercept = convert_error_to_none(intercept_values[idx]) if idx < len(intercept_values) else None
        slope = convert_error_to_none(slope_values[idx]) if idx < len(slope_values) else None

        if all(is_blank(v) for v in (forecast_value, forecast_max, forecast_min)):
            continue

        signature = (
            comparable(num_quarters_used),
            comparable(forecast_value),
            comparable(forecast_max),
            comparable(forecast_min),
            comparable(intercept),
            comparable(slope),
        )
        if previous_signature is not None and signature == previous_signature:
            continue
        previous_signature = signature

        row = {
            "model": labels["model"],
            "ticker": labels["ticker"],
            "model_period": labels["model_period"],
            "model_date": labels["model_date"],
            "method": "regression",
            "parameter_name": "num_quarters_used",
            "parameter_value": num_quarters_used,
            "num_quarters_used": num_quarters_used,
            "forecast_value": forecast_value,
            "actual_value": actual_value if not is_blank(actual_value) else "",
            "forecast_max": forecast_max,
            "forecast_min": forecast_min,
            "range_width": subtract_if_numeric(forecast_max, forecast_min),
            "intercept": intercept,
            "slope": slope,
            "source_file": source_file,
        }
        rows.append(row)

    return rows


def write_output_sheet(
    workbook: Workbook, sheet_name: str, columns: list[str], rows: list[dict[str, Any]]
) -> None:
    sheet = workbook.create_sheet(title=sheet_name)
    sheet.append(columns)
    for row in rows:
        sheet.append([row.get(col, "") for col in columns])

    for cell in sheet[1]:
        cell.font = Font(bold=True)
    sheet.freeze_panes = "A2"
    sheet.auto_filter.ref = sheet.dimensions

    for col_idx, col_name in enumerate(columns, start=1):
        max_len = len(col_name)
        for values in sheet.iter_rows(
            min_row=2, max_row=sheet.max_row, min_col=col_idx, max_col=col_idx, values_only=True
        ):
            cell_value = values[0]
            text_value = "" if cell_value is None else str(cell_value)
            if len(text_value) > max_len:
                max_len = len(text_value)
        sheet.column_dimensions[get_column_letter(col_idx)].width = min(max(12, max_len + 2), 60)


def process_workbooks() -> None:
    input_path = Path(input_dir)
    output_path = Path(output_dir)
    output_path.mkdir(parents=True, exist_ok=True)

    if not input_path.exists():
        raise FileNotFoundError(f"input_dir does not exist: {input_path}")
    if not input_path.is_dir():
        raise NotADirectoryError(f"input_dir is not a directory: {input_path}")

    output_file_path = next_output_path(input_path, output_path)
    empirical_rows: list[dict[str, Any]] = []
    regression_rows: list[dict[str, Any]] = []
    files_processed = 0

    app: xw.App | None = None
    try:
        app = xw.App(visible=False, add_book=False)
        app.display_alerts = False
        app.screen_updating = False

        for file_path in sorted(input_path.iterdir()):
            if not file_path.is_file():
                continue
            if file_path.name.startswith("~"):
                print(f"skipped file: {file_path.name} (temporary file)")
                continue
            if file_path.suffix.lower() != ".xlsx":
                print(f"skipped file: {file_path.name} (not .xlsx)")
                continue

            print(f"processing file: {file_path.name}")
            try:
                labels = parse_file_labels(file_path.name)
            except ValueError as exc:
                print(f"skipped file: {file_path.name} ({exc})")
                continue

            wb: xw.Book | None = None
            try:
                wb = app.books.open(str(file_path), update_links=False)
                empirical_rows.extend(extract_empirical_rows(wb, labels, file_path.name))
                regression_rows.extend(extract_regression_rows(wb, labels, file_path.name))
                files_processed += 1
            except Exception as exc:
                print(f"skipped file: {file_path.name} (processing error: {exc})")
            finally:
                if wb is not None:
                    safe_close_workbook(wb)
    finally:
        if app is not None:
            app.quit()

    output_workbook = Workbook()
    default_sheet = output_workbook.active
    output_workbook.remove(default_sheet)
    write_output_sheet(
        output_workbook,
        "empirical_candidates",
        EMPIRICAL_COLUMNS,
        empirical_rows,
    )
    write_output_sheet(
        output_workbook,
        "regression_candidates",
        REGRESSION_COLUMNS,
        regression_rows,
    )
    output_workbook.save(output_file_path)

    print(f"output path: {output_file_path}")
    print(f"number of files processed: {files_processed}")
    print(f"number of empirical rows: {len(empirical_rows)}")
    print(f"number of regression rows: {len(regression_rows)}")


if __name__ == "__main__":
    process_workbooks()
