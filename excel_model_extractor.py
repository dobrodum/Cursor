#!/usr/bin/env python3
"""Extract empirical and regression candidates from model workbooks."""

from __future__ import annotations

import datetime as dt
import re
from pathlib import Path
from typing import Any

import xlwings as xw
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

# -----------------------------
# User-configurable directories
# -----------------------------
input_dir = Path("input")
output_dir = Path("output")


N_QUARTERS = 10
EMPIRICAL_SHEET_NAME = "Empirical Model"
REGRESSION_SHEET_NAME = "Regression Model"

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

MODEL_DAY_BY_STAGE = {"EARLY": 5, "MID": 15, "LATE": 25}
MONTH_INDEX = {
    "JAN": 1,
    "FEB": 2,
    "MAR": 3,
    "APR": 4,
    "MAY": 5,
    "JUN": 6,
    "JUL": 7,
    "AUG": 8,
    "SEP": 9,
    "OCT": 10,
    "NOV": 11,
    "DEC": 12,
}


def normalize_header(value: Any) -> str:
    """Normalize header text for fuzzy matching."""
    if value is None:
        return ""
    text = str(value).strip().lower()
    if not text:
        return ""
    text = text.replace("\n", " ").replace("_", " ").replace("-", " ")
    text = re.sub(r"[^a-z0-9/% ]+", " ", text)
    text = re.sub(r"\s+", " ", text).strip()
    return text


def as_matrix(values: Any) -> list[list[Any]]:
    """Convert xlwings range values into a 2D list."""
    if values is None:
        return []
    if isinstance(values, list):
        if values and isinstance(values[0], list):
            return values
        return [values]
    return [[values]]


def is_blank(value: Any) -> bool:
    """Check if a value should be treated as empty."""
    if value is None:
        return True
    if isinstance(value, str):
        return value.strip() == ""
    return False


def to_float(value: Any) -> float | None:
    """Convert a value to float when possible."""
    if value is None or value == "":
        return None
    if isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        return float(value)
    text = str(value).strip().replace(",", "")
    if text.endswith("%"):
        text = text[:-1]
    try:
        return float(text)
    except ValueError:
        return None


def safe_difference(max_value: Any, min_value: Any) -> float | None:
    """Return max-min when both are numeric."""
    max_num = to_float(max_value)
    min_num = to_float(min_value)
    if max_num is None or min_num is None:
        return None
    return max_num - min_num


def to_positive_int(value: Any, fallback: int) -> int:
    """Convert value to positive integer with a fallback."""
    if value is None or value == "":
        return fallback
    try:
        parsed = int(float(value))
        if parsed > 0:
            return parsed
    except (TypeError, ValueError):
        pass
    return fallback


def parse_model_labels(file_name: str) -> dict[str, str]:
    """Parse ticker/model_period/model_date/model from source filename."""
    stem = Path(file_name).stem
    parts = [part.strip() for part in re.split(r"\s*-\s*", stem) if part.strip()]

    ticker = parts[1].upper() if len(parts) >= 2 else ""
    period_token = parts[2] if len(parts) >= 3 else ""
    period_token = re.sub(r"_?send$", "", period_token, flags=re.IGNORECASE).strip()

    model_period = period_token
    model_date = ""
    model = ""

    match = re.search(
        r"(?i)\b(early|mid|late)(jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)(\d{4})\b",
        period_token,
    )
    if match:
        stage = match.group(1).capitalize()
        month_abbr = match.group(2).capitalize()
        year = int(match.group(3))
        month_idx = MONTH_INDEX[month_abbr.upper()]
        day = MODEL_DAY_BY_STAGE[stage.upper()]

        model_period = f"{stage}{month_abbr}_{year}"
        model_date = dt.date(year, month_idx, day).isoformat()
        model = f"{ticker}_{model_period}" if ticker else model_period
    else:
        if ticker and model_period:
            model = f"{ticker}_{model_period}"
        else:
            model = ticker or model_period

    return {
        "ticker": ticker,
        "model_period": model_period,
        "model_date": model_date,
        "model": model,
    }


def next_output_path(source_dir: Path, destination_dir: Path) -> Path:
    """Create unique output path with .1/.2 suffix fallback."""
    base_name = f"{source_dir.name}_PARAM"
    candidate = destination_dir / f"{base_name}.xlsx"
    if not candidate.exists():
        return candidate

    suffix = 1
    while True:
        candidate = destination_dir / f"{base_name}.{suffix}.xlsx"
        if not candidate.exists():
            return candidate
        suffix += 1


def close_workbook_safely(workbook: xw.Book) -> None:
    """Close workbook without saving, with API fallback."""
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


def set_r1c1_formula2(cell: xw.Range, formula: str) -> None:
    """Set an R1C1 formula using Formula2 first, then safe fallbacks."""
    try:
        cell.api.Formula2R1C1 = formula
        return
    except Exception:
        pass

    try:
        cell.formula2 = formula
        return
    except Exception:
        pass

    try:
        cell.api.FormulaR1C1 = formula
        return
    except Exception:
        pass

    cell.formula = formula


def find_anchor(sheet: xw.Sheet, anchor_text: str = "max") -> tuple[int, int, int] | None:
    """Find first anchor cell match and return row, col, used_range_start_row."""
    used = sheet.used_range
    data = as_matrix(used.value)
    if not data:
        return None

    base_row = used.row
    base_col = used.column
    target = anchor_text.strip().lower()

    for row_idx, row_values in enumerate(data):
        for col_idx, value in enumerate(row_values):
            if isinstance(value, str) and value.strip().lower() == target:
                return base_row + row_idx, base_col + col_idx, base_row
    return None


def get_header_row_scan(
    sheet: xw.Sheet, header_row: int, anchor_col: int, span: int = 30
) -> list[tuple[int, str]]:
    """Read and normalize a slice of the header row around the anchor."""
    start_col = max(1, anchor_col - span)
    end_col = anchor_col + span
    values = sheet.range((header_row, start_col), (header_row, end_col)).value
    row_values = values if isinstance(values, list) else [values]

    output: list[tuple[int, str]] = []
    for idx, value in enumerate(row_values):
        normalized = normalize_header(value)
        if normalized:
            output.append((start_col + idx, normalized))
    return output


def column_from_patterns(
    headers: list[tuple[int, str]], patterns: list[str | tuple[str, ...]], default_col: int | None
) -> int | None:
    """Resolve a likely column index from normalized header patterns."""
    for col, header in headers:
        for pattern in patterns:
            if isinstance(pattern, str):
                if pattern in header:
                    return col
            else:
                if all(token in header for token in pattern):
                    return col
    return default_col


def read_column_values(sheet: xw.Sheet, start_row: int, end_row: int, col: int | None) -> list[Any]:
    """Read a vertical range as a python list."""
    row_count = max(0, end_row - start_row + 1)
    if row_count == 0 or col is None or col < 1:
        return [None] * row_count

    values = sheet.range((start_row, col), (end_row, col)).value
    if isinstance(values, list):
        return values
    return [values]


def extract_empirical_candidates(
    workbook: xw.Book, metadata: dict[str, str], source_file: str
) -> list[dict[str, Any]]:
    """Extract empirical candidates from the workbook."""
    try:
        sheet = workbook.sheets[EMPIRICAL_SHEET_NAME]
    except Exception:
        print(f"Skipped empirical extraction for {source_file}: sheet '{EMPIRICAL_SHEET_NAME}' not found")
        return []

    anchor_info = find_anchor(sheet, "max")
    if anchor_info is None:
        print(f"Skipped empirical extraction for {source_file}: 'max' anchor not found")
        return []

    anchor_row, anchor_col, used_start_row = anchor_info
    headers = get_header_row_scan(sheet, anchor_row, anchor_col)

    col_num_quarters = column_from_patterns(
        headers,
        [("num", "quarter"), ("quarter", "used"), ("n", "quarter"), ("#", "quarter")],
        anchor_col - 9,
    )
    col_last_quarter = column_from_patterns(
        headers, [("last", "quarter"), "last qtr", "last quarter used"], anchor_col - 8
    )
    col_forecast = column_from_patterns(
        headers,
        [
            ("estimated", "total", "sold"),
            ("forecast", "value"),
            ("forecast", "total"),
            ("tot", "fcst"),
        ],
        anchor_col - 2,
    )
    col_actual = column_from_patterns(
        headers, [("actual", "sales"), ("reported", "sales"), "actual value"], anchor_col - 3
    )
    col_min = column_from_patterns(headers, ["min", "minimum"], anchor_col + 1)
    col_penetration_source = column_from_patterns(
        headers,
        [("avg", "penetration"), ("average", "penetration"), ("penetration", "pct")],
        anchor_col - 7,
    )
    col_quarterly_sales = column_from_patterns(
        headers, [("quarterly", "sales"), ("sales", "quarter")], anchor_col - 4
    )
    col_reported_sales = column_from_patterns(
        headers, [("reported", "sales"), ("actual", "sales")], col_actual
    )
    col_growth = column_from_patterns(
        headers, [("growth", "rate"), ("growth", "pct"), "growth"], anchor_col - 5
    )
    col_sales_captured = column_from_patterns(
        headers, [("captured", "db"), ("sales", "captured"), ("in", "db")], anchor_col - 6
    )

    start_row = anchor_row + 1
    end_row = anchor_row + N_QUARTERS
    row_count = end_row - start_row + 1

    num_quarter_values = read_column_values(sheet, start_row, end_row, col_num_quarters)
    helper_col = anchor_col + 3
    formulas_written = False
    history_end = anchor_row - 1

    if history_end >= used_start_row and col_penetration_source and col_penetration_source > 0:
        for idx, num_value in enumerate(num_quarter_values):
            target_row = start_row + idx
            num_quarters = to_positive_int(num_value, idx + 1)
            history_start = max(used_start_row, history_end - num_quarters + 1)
            formula = (
                f'=IFERROR(AVERAGE(R{history_start}C{col_penetration_source}:'
                f'R{history_end}C{col_penetration_source}),"")'
            )
            set_r1c1_formula2(sheet.range((target_row, helper_col)), formula)
            formulas_written = True

    if formulas_written:
        workbook.app.calculate()

    last_quarter_values = read_column_values(sheet, start_row, end_row, col_last_quarter)
    forecast_values = read_column_values(sheet, start_row, end_row, col_forecast)
    actual_values = read_column_values(sheet, start_row, end_row, col_actual)
    max_values = read_column_values(sheet, start_row, end_row, anchor_col)
    min_values = read_column_values(sheet, start_row, end_row, col_min)
    avg_penetration_values = read_column_values(sheet, start_row, end_row, helper_col)
    quarterly_sales_values = read_column_values(sheet, start_row, end_row, col_quarterly_sales)
    reported_sales_values = read_column_values(sheet, start_row, end_row, col_reported_sales)
    growth_values = read_column_values(sheet, start_row, end_row, col_growth)
    captured_values = read_column_values(sheet, start_row, end_row, col_sales_captured)

    candidates: list[dict[str, Any]] = []
    for idx in range(row_count):
        num_quarters = to_positive_int(num_quarter_values[idx], idx + 1)
        forecast_value = forecast_values[idx]
        actual_value = actual_values[idx]
        forecast_max = max_values[idx]
        forecast_min = min_values[idx]
        avg_penetration = avg_penetration_values[idx]

        if (
            is_blank(num_quarter_values[idx])
            and is_blank(forecast_value)
            and is_blank(forecast_max)
            and is_blank(forecast_min)
            and is_blank(avg_penetration)
        ):
            continue

        candidates.append(
            {
                "model": metadata["model"],
                "ticker": metadata["ticker"],
                "model_period": metadata["model_period"],
                "model_date": metadata["model_date"],
                "method": "empirical",
                "parameter_name": "avg_penetration_pct",
                "parameter_value": avg_penetration,
                "num_quarters_used": num_quarters,
                "last_quarter_used": last_quarter_values[idx],
                "forecast_value": forecast_value,
                "actual_value": actual_value,
                "forecast_max": forecast_max,
                "forecast_min": forecast_min,
                "range_width": safe_difference(forecast_max, forecast_min),
                "avg_penetration_pct": avg_penetration,
                "quarterly_sales": quarterly_sales_values[idx],
                "reported_sales": reported_sales_values[idx],
                "growth_rate_pct": growth_values[idx],
                "sales_captured_in_db_pct": captured_values[idx],
                "source_file": source_file,
            }
        )

    return candidates


def dedupe_signature(values: list[Any]) -> tuple[Any, ...]:
    """Build stable signature used to skip duplicate final regression rows."""
    output: list[Any] = []
    for value in values:
        numeric = to_float(value)
        if numeric is None:
            output.append("" if value is None else str(value).strip())
        else:
            output.append(round(numeric, 10))
    return tuple(output)


def extract_regression_candidates(
    workbook: xw.Book, metadata: dict[str, str], source_file: str
) -> list[dict[str, Any]]:
    """Extract regression candidates from the workbook."""
    try:
        sheet = workbook.sheets[REGRESSION_SHEET_NAME]
    except Exception:
        print(f"Skipped regression extraction for {source_file}: sheet '{REGRESSION_SHEET_NAME}' not found")
        return []

    anchor_info = find_anchor(sheet, "max")
    if anchor_info is None:
        print(f"Skipped regression extraction for {source_file}: 'max' anchor not found")
        return []

    anchor_row, anchor_col, used_start_row = anchor_info
    y_col = anchor_col - 7
    x_col = anchor_col - 11
    headers = get_header_row_scan(sheet, anchor_row, anchor_col)

    col_num_quarters = column_from_patterns(
        headers,
        [("num", "quarter"), ("quarter", "used"), ("n", "quarter"), ("#", "quarter")],
        anchor_col - 12,
    )
    col_forecast = column_from_patterns(
        headers,
        [
            ("tot", "fcst", "w/o", "sa"),
            ("tot", "fcst", "wo", "sa"),
            ("forecast", "without", "sa"),
            ("forecast", "w/o", "sa"),
        ],
        anchor_col - 1,
    )
    col_actual = column_from_patterns(headers, [("actual", "sales"), "actual value"], None)
    col_min = column_from_patterns(headers, ["min", "minimum"], anchor_col + 1)

    start_row = anchor_row + 1
    end_row = anchor_row + N_QUARTERS
    row_count = end_row - start_row + 1
    num_quarter_values = read_column_values(sheet, start_row, end_row, col_num_quarters)

    intercept_col = anchor_col + 3
    slope_col = anchor_col + 4
    formulas_written = False
    history_end = anchor_row - 1

    if history_end >= used_start_row:
        for idx, num_value in enumerate(num_quarter_values):
            target_row = start_row + idx
            num_quarters = to_positive_int(num_value, idx + 1)
            history_start = max(used_start_row, history_end - num_quarters + 1)
            intercept_formula = (
                f'=IFERROR(INTERCEPT(R{history_start}C{y_col}:R{history_end}C{y_col},'
                f'R{history_start}C{x_col}:R{history_end}C{x_col}),"")'
            )
            slope_formula = (
                f'=IFERROR(SLOPE(R{history_start}C{y_col}:R{history_end}C{y_col},'
                f'R{history_start}C{x_col}:R{history_end}C{x_col}),"")'
            )
            set_r1c1_formula2(sheet.range((target_row, intercept_col)), intercept_formula)
            set_r1c1_formula2(sheet.range((target_row, slope_col)), slope_formula)
            formulas_written = True

    if formulas_written:
        workbook.app.calculate()

    forecast_values = read_column_values(sheet, start_row, end_row, col_forecast)
    actual_values = read_column_values(sheet, start_row, end_row, col_actual)
    max_values = read_column_values(sheet, start_row, end_row, anchor_col)
    min_values = read_column_values(sheet, start_row, end_row, col_min)
    intercept_values = read_column_values(sheet, start_row, end_row, intercept_col)
    slope_values = read_column_values(sheet, start_row, end_row, slope_col)

    candidates: list[dict[str, Any]] = []
    previous_signature: tuple[Any, ...] | None = None

    for idx in range(row_count):
        num_quarters = to_positive_int(num_quarter_values[idx], idx + 1)
        forecast_value = forecast_values[idx]
        actual_value = actual_values[idx] if idx < len(actual_values) else None
        forecast_max = max_values[idx]
        forecast_min = min_values[idx]
        intercept = intercept_values[idx]
        slope = slope_values[idx]

        if (
            is_blank(num_quarter_values[idx])
            and is_blank(forecast_value)
            and is_blank(forecast_max)
            and is_blank(forecast_min)
            and is_blank(intercept)
            and is_blank(slope)
        ):
            continue

        signature = dedupe_signature(
            [num_quarters, forecast_value, forecast_max, forecast_min, intercept, slope]
        )
        if previous_signature == signature:
            continue
        previous_signature = signature

        candidates.append(
            {
                "model": metadata["model"],
                "ticker": metadata["ticker"],
                "model_period": metadata["model_period"],
                "model_date": metadata["model_date"],
                "method": "regression",
                "parameter_name": "num_quarters_used",
                "parameter_value": num_quarters,
                "num_quarters_used": num_quarters,
                "forecast_value": forecast_value,
                "actual_value": actual_value,
                "forecast_max": forecast_max,
                "forecast_min": forecast_min,
                "range_width": safe_difference(forecast_max, forecast_min),
                "intercept": intercept,
                "slope": slope,
                "source_file": source_file,
            }
        )

    return candidates


def apply_sheet_formatting(worksheet, columns: list[str]) -> None:
    """Apply lightweight output formatting for readability."""
    for cell in worksheet[1]:
        cell.font = Font(bold=True)

    worksheet.freeze_panes = "A2"
    worksheet.auto_filter.ref = f"A1:{get_column_letter(len(columns))}{worksheet.max_row}"

    for idx, column_name in enumerate(columns, start=1):
        max_length = len(column_name)
        for row in worksheet.iter_rows(
            min_row=2, max_row=worksheet.max_row, min_col=idx, max_col=idx, values_only=True
        ):
            value = row[0]
            if value is not None:
                max_length = max(max_length, len(str(value)))
        worksheet.column_dimensions[get_column_letter(idx)].width = min(max(max_length + 2, 12), 42)


def write_output_workbook(
    output_path: Path,
    empirical_rows: list[dict[str, Any]],
    regression_rows: list[dict[str, Any]],
) -> None:
    """Write both candidate tables into one workbook."""
    workbook = Workbook()

    empirical_sheet = workbook.active
    empirical_sheet.title = "empirical_candidates"
    empirical_sheet.append(EMPIRICAL_COLUMNS)
    for row in empirical_rows:
        empirical_sheet.append([row.get(column) for column in EMPIRICAL_COLUMNS])
    apply_sheet_formatting(empirical_sheet, EMPIRICAL_COLUMNS)

    regression_sheet = workbook.create_sheet("regression_candidates")
    regression_sheet.append(REGRESSION_COLUMNS)
    for row in regression_rows:
        regression_sheet.append([row.get(column) for column in REGRESSION_COLUMNS])
    apply_sheet_formatting(regression_sheet, REGRESSION_COLUMNS)

    workbook.save(output_path)


def list_source_files(source_dir: Path) -> list[Path]:
    """Return all valid source .xlsx files and print skip reasons."""
    files: list[Path] = []
    for path in sorted(source_dir.iterdir()):
        if not path.is_file():
            print(f"Skipped file: {path.name} (not a regular file)")
            continue
        if path.name.startswith("~"):
            print(f"Skipped file: {path.name} (temporary workbook)")
            continue
        if path.suffix.lower() != ".xlsx":
            print(f"Skipped file: {path.name} (not .xlsx)")
            continue
        files.append(path)
    return files


def main() -> None:
    source_dir = input_dir.expanduser().resolve()
    destination_dir = output_dir.expanduser().resolve()
    destination_dir.mkdir(parents=True, exist_ok=True)

    if not source_dir.exists() or not source_dir.is_dir():
        raise FileNotFoundError(f"Input directory does not exist: {source_dir}")

    source_files = list_source_files(source_dir)
    output_path = next_output_path(source_dir, destination_dir)

    empirical_rows: list[dict[str, Any]] = []
    regression_rows: list[dict[str, Any]] = []
    processed_files = 0

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
        for file_path in source_files:
            metadata = parse_model_labels(file_path.name)
            workbook = None
            try:
                workbook = app.books.open(str(file_path), update_links=False)
                empirical_rows.extend(
                    extract_empirical_candidates(workbook, metadata, file_path.name)
                )
                regression_rows.extend(
                    extract_regression_candidates(workbook, metadata, file_path.name)
                )
                processed_files += 1
                print(f"Processed file: {file_path.name}")
            except Exception as exc:
                print(f"Skipped file: {file_path.name} ({exc})")
            finally:
                if workbook is not None:
                    close_workbook_safely(workbook)
    finally:
        app.quit()

    write_output_workbook(output_path, empirical_rows, regression_rows)

    print(f"Output path: {output_path}")
    print(f"Files processed: {processed_files}")
    print(f"Empirical rows: {len(empirical_rows)}")
    print(f"Regression rows: {len(regression_rows)}")


if __name__ == "__main__":
    main()
