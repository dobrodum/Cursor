from __future__ import annotations

import re
from datetime import date
from pathlib import Path
from typing import Any

import xlwings as xw
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter


# Configure these two paths before running.
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

MONTH_TO_NUMBER = {
    "Jan": 1,
    "Feb": 2,
    "Mar": 3,
    "Apr": 4,
    "May": 5,
    "Jun": 6,
    "Jul": 7,
    "Aug": 8,
    "Sep": 9,
    "Oct": 10,
    "Nov": 11,
    "Dec": 12,
}

DAY_BY_PERIOD = {"Early": 5, "Mid": 15, "Late": 25}
PERIOD_RE = re.compile(
    r"\b(Early|Mid|Late)(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Sept|Oct|Nov|Dec)(\d{4})\b",
    flags=re.IGNORECASE,
)
OUTPUT_NAME_RE = re.compile(r"_PARAM(\.\d+)?\.xlsx$", flags=re.IGNORECASE)


def normalize_text(value: Any) -> str:
    if value is None:
        return ""
    text = str(value).strip().lower()
    text = re.sub(r"[^a-z0-9]+", " ", text)
    return re.sub(r"\s+", " ", text).strip()


def to_2d(values: Any) -> list[list[Any]]:
    if values is None:
        return []
    if not isinstance(values, list):
        return [[values]]
    if values and not isinstance(values[0], list):
        return [values]
    matrix: list[list[Any]] = [list(row) for row in values]
    if not matrix:
        return matrix
    max_len = max(len(row) for row in matrix)
    for row in matrix:
        if len(row) < max_len:
            row.extend([None] * (max_len - len(row)))
    return matrix


def flatten_column(values: Any, expected_len: int) -> list[Any]:
    matrix = to_2d(values)
    result: list[Any] = []
    for row in matrix:
        result.append(row[0] if row else None)
    if len(result) < expected_len:
        result.extend([None] * (expected_len - len(result)))
    return result[:expected_len]


def is_blank(value: Any) -> bool:
    return value is None or (isinstance(value, str) and value.strip() == "")


def as_float(value: Any) -> float | None:
    if value is None or isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        return float(value)
    if isinstance(value, str):
        text = value.strip().replace(",", "")
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
    return None


def to_int(value: Any) -> int | None:
    numeric = as_float(value)
    if numeric is None:
        return None
    return int(round(numeric))


def numeric_difference(a: Any, b: Any) -> float | None:
    left = as_float(a)
    right = as_float(b)
    if left is None or right is None:
        return None
    return left - right


def safe_matrix_get(
    matrix: list[list[Any]],
    start_row: int,
    start_col: int,
    row: int,
    col: int | None,
) -> Any:
    if col is None:
        return None
    r_idx = row - start_row
    c_idx = col - start_col
    if r_idx < 0 or c_idx < 0:
        return None
    if r_idx >= len(matrix):
        return None
    if c_idx >= len(matrix[r_idx]):
        return None
    return matrix[r_idx][c_idx]


def find_anchor_max(
    matrix: list[list[Any]], start_row: int, start_col: int
) -> tuple[int, int] | None:
    for r_idx, row in enumerate(matrix):
        for c_idx, value in enumerate(row):
            if normalize_text(value) == "max":
                return start_row + r_idx, start_col + c_idx
    return None


def find_column_by_aliases(
    header_row: list[Any],
    start_col: int,
    aliases: list[str],
    anchor_col: int | None = None,
    default: int | None = None,
) -> int | None:
    matches: list[tuple[int, int]] = []
    for idx, raw_value in enumerate(header_row):
        normalized = normalize_text(raw_value)
        if not normalized:
            continue
        if any(alias in normalized for alias in aliases):
            column = start_col + idx
            distance = 0 if anchor_col is None else abs(column - anchor_col)
            matches.append((distance, column))
    if matches:
        matches.sort(key=lambda item: (item[0], item[1]))
        return matches[0][1]
    return default


def parse_file_labels(file_name: str) -> dict[str, str]:
    stem = Path(file_name).stem
    split_parts = [part.strip() for part in stem.split(" - ")]

    ticker = "UNKNOWN"
    if len(split_parts) >= 2 and split_parts[1]:
        ticker = re.sub(r"[^A-Za-z0-9]", "", split_parts[1]).upper() or "UNKNOWN"

    model_period = "unknown_period"
    model_date = ""
    period_match = PERIOD_RE.search(stem)
    if period_match:
        period_word = period_match.group(1).title()
        month_word = period_match.group(2).title()
        if month_word == "Sept":
            month_word = "Sep"
        year_word = period_match.group(3)
        day = DAY_BY_PERIOD[period_word]
        month_number = MONTH_TO_NUMBER[month_word]
        model_period = f"{period_word}{month_word}_{year_word}"
        model_date = date(int(year_word), month_number, day).isoformat()

    model = f"{ticker}_{model_period}"
    return {
        "ticker": ticker,
        "model_period": model_period,
        "model_date": model_date,
        "model": model,
    }


def set_formula2(target_range: xw.Range, formula_values: list[list[str]]) -> None:
    try:
        target_range.formula2 = formula_values
    except Exception:
        target_range.formula = formula_values


def close_source_workbook(workbook: xw.Book) -> None:
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
        return
    except Exception:
        pass

    try:
        workbook.close()
    except Exception:
        pass


def build_output_path(input_path: Path, output_path: Path) -> Path:
    output_path.mkdir(parents=True, exist_ok=True)
    folder_name = input_path.resolve().name
    base = output_path / f"{folder_name}_PARAM.xlsx"
    if not base.exists():
        return base
    index = 1
    while True:
        candidate = output_path / f"{folder_name}_PARAM.{index}.xlsx"
        if not candidate.exists():
            return candidate
        index += 1


def extract_empirical_rows(
    workbook: xw.Book, labels: dict[str, str], source_file: str
) -> list[dict[str, Any]]:
    try:
        sheet = workbook.sheets["Empirical Model"]
    except Exception:
        return []

    used = sheet.used_range
    matrix = to_2d(used.value)
    if not matrix:
        return []

    start_row = used.row
    start_col = used.column
    used_last_col = start_col + len(matrix[0]) - 1
    anchor = find_anchor_max(matrix, start_row, start_col)
    if anchor is None:
        return []

    anchor_row, anchor_col = anchor
    header_row = matrix[anchor_row - start_row]

    min_col = find_column_by_aliases(
        header_row, start_col, ["min"], anchor_col=anchor_col, default=anchor_col + 1
    )
    forecast_col = find_column_by_aliases(
        header_row,
        start_col,
        ["estimated total sold", "forecast value", "forecast total", "tot fcst"],
        anchor_col=anchor_col,
        default=anchor_col - 1,
    )
    actual_col = find_column_by_aliases(
        header_row,
        start_col,
        ["reported sales", "actual sales", "actual value"],
        anchor_col=anchor_col,
    )
    num_quarters_col = find_column_by_aliases(
        header_row,
        start_col,
        ["num quarters used", "quarters used", "n quarters", "quarters"],
        anchor_col=anchor_col,
        default=anchor_col - 10,
    )
    last_quarter_col = find_column_by_aliases(
        header_row,
        start_col,
        ["last quarter used", "last quarter"],
        anchor_col=anchor_col,
    )
    quarterly_sales_col = find_column_by_aliases(
        header_row,
        start_col,
        ["quarterly sales"],
        anchor_col=anchor_col,
    )
    reported_sales_col = find_column_by_aliases(
        header_row,
        start_col,
        ["reported sales"],
        anchor_col=anchor_col,
        default=actual_col,
    )
    growth_rate_col = find_column_by_aliases(
        header_row,
        start_col,
        ["growth rate", "growth"],
        anchor_col=anchor_col,
    )
    sales_captured_col = find_column_by_aliases(
        header_row,
        start_col,
        ["sales captured in db", "captured in db", "captured"],
        anchor_col=anchor_col,
    )

    data_start_row = anchor_row + 1
    data_end_row = data_start_row + 10 - 1

    helper_col = max(used_last_col + 2, anchor_col + 2)
    avg_pen_formulas: list[list[str]] = []
    can_calculate_avg_pen = quarterly_sales_col is not None and reported_sales_col is not None
    for _ in range(data_start_row, data_end_row + 1):
        if can_calculate_avg_pen:
            q_offset = quarterly_sales_col - helper_col
            r_offset = reported_sales_col - helper_col
            avg_pen_formulas.append([f'=IFERROR(RC[{q_offset}]/RC[{r_offset}], "")'])
        else:
            avg_pen_formulas.append([""])

    helper_range = sheet.range((data_start_row, helper_col), (data_end_row, helper_col))
    set_formula2(helper_range, avg_pen_formulas)
    if can_calculate_avg_pen:
        workbook.app.calculate()
    avg_pen_values = flatten_column(helper_range.value, data_end_row - data_start_row + 1)
    helper_range.value = [[None] for _ in avg_pen_values]

    rows: list[dict[str, Any]] = []
    for idx, row_number in enumerate(range(data_start_row, data_end_row + 1), start=1):
        raw_num_quarters = safe_matrix_get(
            matrix, start_row, start_col, row_number, num_quarters_col
        )
        forecast_max = safe_matrix_get(matrix, start_row, start_col, row_number, anchor_col)
        forecast_min = safe_matrix_get(matrix, start_row, start_col, row_number, min_col)
        forecast_value = safe_matrix_get(
            matrix, start_row, start_col, row_number, forecast_col
        )
        actual_value = safe_matrix_get(matrix, start_row, start_col, row_number, actual_col)
        if all(
            is_blank(value)
            for value in [raw_num_quarters, forecast_max, forecast_min, forecast_value, actual_value]
        ):
            continue

        num_quarters_used = raw_num_quarters if not is_blank(raw_num_quarters) else idx
        last_quarter_used = safe_matrix_get(
            matrix, start_row, start_col, row_number, last_quarter_col
        )
        quarterly_sales = safe_matrix_get(
            matrix, start_row, start_col, row_number, quarterly_sales_col
        )
        reported_sales = safe_matrix_get(
            matrix, start_row, start_col, row_number, reported_sales_col
        )
        growth_rate_pct = safe_matrix_get(
            matrix, start_row, start_col, row_number, growth_rate_col
        )
        sales_captured_pct = safe_matrix_get(
            matrix, start_row, start_col, row_number, sales_captured_col
        )
        avg_penetration_pct = avg_pen_values[idx - 1] if idx - 1 < len(avg_pen_values) else None
        range_width = numeric_difference(forecast_max, forecast_min)

        rows.append(
            {
                "model": labels["model"],
                "ticker": labels["ticker"],
                "model_period": labels["model_period"],
                "model_date": labels["model_date"],
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


def comparable_value(value: Any) -> Any:
    numeric = as_float(value)
    if numeric is not None:
        return round(numeric, 12)
    if is_blank(value):
        return ""
    return str(value).strip()


def extract_regression_rows(
    workbook: xw.Book, labels: dict[str, str], source_file: str
) -> list[dict[str, Any]]:
    try:
        sheet = workbook.sheets["Regression Model"]
    except Exception:
        return []

    used = sheet.used_range
    matrix = to_2d(used.value)
    if not matrix:
        return []

    start_row = used.row
    start_col = used.column
    used_last_col = start_col + len(matrix[0]) - 1
    anchor = find_anchor_max(matrix, start_row, start_col)
    if anchor is None:
        return []

    anchor_row, anchor_col = anchor
    header_row = matrix[anchor_row - start_row]

    y_col = anchor_col - 7
    x_col = anchor_col - 11
    min_col = find_column_by_aliases(
        header_row, start_col, ["min"], anchor_col=anchor_col, default=anchor_col + 1
    )
    forecast_col = find_column_by_aliases(
        header_row,
        start_col,
        ["tot fcst w o sa", "tot fcst", "forecast total without sa", "forecast value"],
        anchor_col=anchor_col,
        default=anchor_col - 1,
    )
    actual_col = find_column_by_aliases(
        header_row,
        start_col,
        ["actual value", "actual sales", "reported sales"],
        anchor_col=anchor_col,
    )
    num_quarters_col = find_column_by_aliases(
        header_row,
        start_col,
        ["num quarters used", "quarters used", "n quarters", "quarters"],
        anchor_col=anchor_col,
        default=anchor_col - 10,
    )

    data_start_row = anchor_row + 1
    candidate_rows: list[int] = []
    empty_streak = 0
    for offset in range(0, 60):
        row_number = data_start_row + offset
        num_quarters_used = safe_matrix_get(
            matrix, start_row, start_col, row_number, num_quarters_col
        )
        forecast_value = safe_matrix_get(matrix, start_row, start_col, row_number, forecast_col)
        forecast_max = safe_matrix_get(matrix, start_row, start_col, row_number, anchor_col)
        forecast_min = safe_matrix_get(matrix, start_row, start_col, row_number, min_col)
        if all(
            is_blank(value)
            for value in [num_quarters_used, forecast_value, forecast_max, forecast_min]
        ):
            empty_streak += 1
            if empty_streak >= 3 and offset > 5:
                break
            continue
        empty_streak = 0
        candidate_rows.append(row_number)

    if not candidate_rows:
        return []

    first_candidate = min(candidate_rows)
    last_candidate = max(candidate_rows)
    helper_intercept_col = max(used_last_col + 2, anchor_col + 2)
    helper_slope_col = helper_intercept_col + 1

    intercept_formulas: list[list[str]] = []
    slope_formulas: list[list[str]] = []
    for row_number in range(first_candidate, last_candidate + 1):
        n_quarters = to_int(
            safe_matrix_get(matrix, start_row, start_col, row_number, num_quarters_col)
        )
        if n_quarters is None or n_quarters < 2:
            intercept_formulas.append([""])
            slope_formulas.append([""])
            continue

        start_data_row = max(data_start_row, row_number - n_quarters + 1)
        intercept_formulas.append(
            [
                (
                    f'=IFERROR(INTERCEPT(R{start_data_row}C{y_col}:R{row_number}C{y_col},'
                    f'R{start_data_row}C{x_col}:R{row_number}C{x_col}), "")'
                )
            ]
        )
        slope_formulas.append(
            [
                (
                    f'=IFERROR(SLOPE(R{start_data_row}C{y_col}:R{row_number}C{y_col},'
                    f'R{start_data_row}C{x_col}:R{row_number}C{x_col}), "")'
                )
            ]
        )

    intercept_range = sheet.range(
        (first_candidate, helper_intercept_col), (last_candidate, helper_intercept_col)
    )
    slope_range = sheet.range((first_candidate, helper_slope_col), (last_candidate, helper_slope_col))
    set_formula2(intercept_range, intercept_formulas)
    set_formula2(slope_range, slope_formulas)
    workbook.app.calculate()

    intercept_values = flatten_column(intercept_range.value, last_candidate - first_candidate + 1)
    slope_values = flatten_column(slope_range.value, last_candidate - first_candidate + 1)
    intercept_range.value = [[None] for _ in intercept_values]
    slope_range.value = [[None] for _ in slope_values]

    rows: list[dict[str, Any]] = []
    previous_signature: tuple[Any, ...] | None = None
    for row_number in candidate_rows:
        list_index = row_number - first_candidate
        num_quarters_used = safe_matrix_get(
            matrix, start_row, start_col, row_number, num_quarters_col
        )
        forecast_value = safe_matrix_get(matrix, start_row, start_col, row_number, forecast_col)
        actual_value = safe_matrix_get(matrix, start_row, start_col, row_number, actual_col)
        forecast_max = safe_matrix_get(matrix, start_row, start_col, row_number, anchor_col)
        forecast_min = safe_matrix_get(matrix, start_row, start_col, row_number, min_col)
        intercept = intercept_values[list_index] if list_index < len(intercept_values) else None
        slope = slope_values[list_index] if list_index < len(slope_values) else None

        signature = (
            comparable_value(num_quarters_used),
            comparable_value(forecast_value),
            comparable_value(forecast_max),
            comparable_value(forecast_min),
            comparable_value(intercept),
            comparable_value(slope),
        )
        if signature == previous_signature:
            continue
        previous_signature = signature

        rows.append(
            {
                "model": labels["model"],
                "ticker": labels["ticker"],
                "model_period": labels["model_period"],
                "model_date": labels["model_date"],
                "method": "regression",
                "parameter_name": "num_quarters_used",
                "parameter_value": num_quarters_used,
                "num_quarters_used": num_quarters_used,
                "forecast_value": forecast_value,
                "actual_value": actual_value,
                "forecast_max": forecast_max,
                "forecast_min": forecast_min,
                "range_width": numeric_difference(forecast_max, forecast_min),
                "intercept": intercept,
                "slope": slope,
                "source_file": source_file,
            }
        )

    return rows


def write_table_sheet(
    workbook: Workbook, sheet_name: str, columns: list[str], rows: list[dict[str, Any]]
) -> None:
    sheet = workbook.create_sheet(title=sheet_name)
    sheet.append(columns)
    for item in rows:
        sheet.append([item.get(column) for column in columns])

    for cell in sheet[1]:
        cell.font = Font(bold=True)

    sheet.freeze_panes = "A2"
    max_row = max(sheet.max_row, 1)
    max_col = len(columns)
    sheet.auto_filter.ref = f"A1:{get_column_letter(max_col)}{max_row}"

    for col_idx, name in enumerate(columns, start=1):
        width = max(12, min(45, len(name) + 2))
        for row_idx in range(2, sheet.max_row + 1):
            value = sheet.cell(row=row_idx, column=col_idx).value
            if value is None:
                continue
            width = max(width, min(45, len(str(value)) + 2))
        sheet.column_dimensions[get_column_letter(col_idx)].width = width


def write_output_workbook(
    output_path: Path, empirical_rows: list[dict[str, Any]], regression_rows: list[dict[str, Any]]
) -> None:
    workbook = Workbook()
    default_sheet = workbook.active
    workbook.remove(default_sheet)

    write_table_sheet(workbook, "empirical_candidates", EMPIRICAL_COLUMNS, empirical_rows)
    write_table_sheet(workbook, "regression_candidates", REGRESSION_COLUMNS, regression_rows)
    workbook.save(output_path)


def process_workbooks() -> None:
    source_dir = Path(input_dir).expanduser().resolve()
    destination_dir = Path(output_dir).expanduser().resolve()
    if not source_dir.exists() or not source_dir.is_dir():
        raise FileNotFoundError(f"Input directory not found: {source_dir}")

    output_path = build_output_path(source_dir, destination_dir)

    empirical_rows: list[dict[str, Any]] = []
    regression_rows: list[dict[str, Any]] = []
    processed_file_count = 0

    app = xw.App(visible=False, add_book=False)
    app.display_alerts = False
    app.screen_updating = False

    try:
        for file_path in sorted(source_dir.iterdir()):
            if file_path.is_dir():
                print(f"Skipped file: {file_path.name} (is a directory)")
                continue
            if file_path.name.startswith("~"):
                print(f"Skipped file: {file_path.name} (temporary file)")
                continue
            if file_path.suffix.lower() != ".xlsx":
                print(f"Skipped file: {file_path.name} (not an .xlsx file)")
                continue
            if OUTPUT_NAME_RE.search(file_path.name):
                print(f"Skipped file: {file_path.name} (looks like output workbook)")
                continue

            labels = parse_file_labels(file_path.name)
            workbook: xw.Book | None = None
            try:
                workbook = app.books.open(str(file_path), update_links=False)
                empirical_rows.extend(
                    extract_empirical_rows(workbook, labels=labels, source_file=file_path.name)
                )
                regression_rows.extend(
                    extract_regression_rows(workbook, labels=labels, source_file=file_path.name)
                )
                processed_file_count += 1
                print(f"Processed file: {file_path.name}")
            except Exception as exc:
                print(f"Skipped file: {file_path.name} (error: {exc})")
            finally:
                if workbook is not None:
                    close_source_workbook(workbook)
    finally:
        app.quit()

    write_output_workbook(output_path, empirical_rows, regression_rows)
    print(f"Output path: {output_path}")
    print(f"Number of files processed: {processed_file_count}")
    print(f"Number of empirical rows: {len(empirical_rows)}")
    print(f"Number of regression rows: {len(regression_rows)}")


if __name__ == "__main__":
    process_workbooks()
