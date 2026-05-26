#!/usr/bin/env python3
"""Extract empirical and regression model candidates from .xlsx workbooks."""

from __future__ import annotations

import calendar
import math
import re
from dataclasses import dataclass
from datetime import date
from pathlib import Path
from typing import Any

import xlwings as xw
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

# Configure paths here.
input_dir = Path("./input")
output_dir = Path("./output")

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

PERIOD_DAY_MAP = {"Early": 5, "Mid": 15, "Late": 25}
MONTH_LOOKUP = {
    **{name.lower(): idx for idx, name in enumerate(calendar.month_name) if idx},
    **{abbr.lower(): idx for idx, abbr in enumerate(calendar.month_abbr) if idx},
    "sept": 9,
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
    text = re.sub(r"\s+", " ", text)
    return text


def parse_month_token(month_token: str) -> int | None:
    token = month_token.strip().lower()
    if token in MONTH_LOOKUP:
        return MONTH_LOOKUP[token]
    short = token[:3]
    return MONTH_LOOKUP.get(short)


def parse_file_label(file_name: str) -> FileLabel:
    stem = Path(file_name).stem
    ticker = "UNKNOWN"
    period_token: str | None = None

    match = re.search(
        r"Model\s*-\s*(?P<ticker>[^-]+?)\s*-\s*(?P<period>[A-Za-z]+\d{4})",
        stem,
        flags=re.IGNORECASE,
    )
    if match:
        ticker = re.sub(r"\s+", "", match.group("ticker")).upper()
        period_token = match.group("period")
    else:
        parts = [part.strip() for part in stem.split("-")]
        if len(parts) >= 2 and parts[1]:
            ticker = re.sub(r"\s+", "", parts[1]).upper()
        for part in parts[2:]:
            token_match = re.search(
                r"(Early|Mid|Late)[A-Za-z]+\d{4}",
                part,
                flags=re.IGNORECASE,
            )
            if token_match:
                period_token = token_match.group(0)
                break

    model_period = "Unknown_0000"
    model_date = ""
    if period_token:
        period_match = re.fullmatch(
            r"(Early|Mid|Late)([A-Za-z]+)(\d{4})",
            period_token,
            flags=re.IGNORECASE,
        )
        if period_match:
            period_label = period_match.group(1).title()
            month_token = period_match.group(2)
            year = int(period_match.group(3))
            month_num = parse_month_token(month_token)
            if month_num:
                model_period = f"{period_label}{calendar.month_abbr[month_num]}_{year}"
                model_date = date(
                    year,
                    month_num,
                    PERIOD_DAY_MAP.get(period_label, 15),
                ).isoformat()

    model = f"{ticker}_{model_period}"
    return FileLabel(
        model=model,
        ticker=ticker,
        model_period=model_period,
        model_date=model_date,
    )


def pick_output_path(source_input_dir: Path, target_output_dir: Path) -> Path:
    target_output_dir.mkdir(parents=True, exist_ok=True)
    base_name = f"{source_input_dir.resolve().name}_PARAM"
    primary = target_output_dir / f"{base_name}.xlsx"
    if not primary.exists():
        return primary

    suffix = 1
    while True:
        candidate = target_output_dir / f"{base_name}.{suffix}.xlsx"
        if not candidate.exists():
            return candidate
        suffix += 1


def get_sheet_case_insensitive(workbook: xw.Book, sheet_name: str) -> xw.Sheet | None:
    target = sheet_name.strip().lower()
    for sheet in workbook.sheets:
        if sheet.name.strip().lower() == target:
            return sheet
    return None


def safe_close_source_workbook(workbook: xw.Book | None) -> None:
    if workbook is None:
        return
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
        workbook.api.Close(False)
        return
    except Exception:
        pass

    try:
        workbook.close()
    except Exception:
        pass


def normalize_used_range_values(values: Any, n_rows: int, n_cols: int) -> list[list[Any]]:
    if values is None:
        return []
    if n_rows == 1 and n_cols == 1:
        return [[values]]
    if n_rows == 1:
        return [values if isinstance(values, list) else [values]]
    if n_cols == 1:
        if isinstance(values, list):
            return [[item] for item in values]
        return [[values]]
    return values


def get_sheet_snapshot(
    sheet: xw.Sheet,
) -> tuple[list[list[Any]], int, int, int, int]:
    used = sheet.used_range
    start_row = used.row
    start_col = used.column
    n_rows = used.rows.count
    n_cols = used.columns.count
    values = normalize_used_range_values(used.value, n_rows, n_cols)
    end_row = start_row + max(n_rows, 1) - 1
    end_col = start_col + max(n_cols, 1) - 1
    return values, start_row, start_col, end_row, end_col


def find_max_anchor(
    matrix: list[list[Any]],
    start_row: int,
    start_col: int,
) -> tuple[int, int] | None:
    for row_idx, row_values in enumerate(matrix):
        for col_idx, cell_value in enumerate(row_values):
            if normalize_text(cell_value) == "max":
                return start_row + row_idx, start_col + col_idx
    return None


def find_nearby_column(
    matrix: list[list[Any]],
    start_row: int,
    start_col: int,
    anchor_row: int,
    anchor_col: int,
    keyword_groups: list[tuple[str, ...]],
    row_window: int = 6,
) -> int | None:
    best_col: int | None = None
    best_score: int | None = None

    for row_idx, row_values in enumerate(matrix):
        abs_row = start_row + row_idx
        if abs(abs_row - anchor_row) > row_window:
            continue
        for col_idx, cell_value in enumerate(row_values):
            label = normalize_text(cell_value)
            if not label:
                continue
            reduced = re.sub(r"[^a-z0-9]+", " ", label)
            if any(all(token in reduced for token in group) for group in keyword_groups):
                abs_col = start_col + col_idx
                score = abs(abs_row - anchor_row) * 100 + abs(abs_col - anchor_col)
                if best_score is None or score < best_score:
                    best_score = score
                    best_col = abs_col
    return best_col


def clamp_col(candidate_col: int | None, start_col: int, end_col: int) -> int | None:
    if candidate_col is None:
        return None
    if start_col <= candidate_col <= end_col:
        return candidate_col
    return None


def coalesce(*values: int | None) -> int | None:
    for value in values:
        if value is not None:
            return value
    return None


def read_column(sheet: xw.Sheet, col: int | None, start_row: int, end_row: int) -> list[Any]:
    count = end_row - start_row + 1
    if count <= 0:
        return []
    if col is None:
        return [None] * count

    values = sheet.range((start_row, col), (end_row, col)).value
    if count == 1:
        return [values]
    if isinstance(values, list):
        if values and isinstance(values[0], list):
            return [row[0] if row else None for row in values]
        return values
    return [values] * count


def is_blank(value: Any) -> bool:
    return value is None or (isinstance(value, str) and not value.strip())


def row_is_blank(values: list[Any]) -> bool:
    return all(is_blank(value) for value in values)


def to_numeric(value: Any) -> float | int | None:
    if value is None:
        return None
    if isinstance(value, bool):
        return int(value)
    if isinstance(value, (int, float)):
        if isinstance(value, float) and math.isnan(value):
            return None
        return value
    if isinstance(value, str):
        stripped = value.strip()
        if not stripped:
            return None
        cleaned = stripped.replace(",", "").replace("%", "")
        try:
            return float(cleaned)
        except ValueError:
            return None
    return None


def number_or_value(value: Any) -> Any:
    numeric = to_numeric(value)
    if numeric is not None:
        return numeric
    if isinstance(value, str):
        return value.strip()
    return value


def safe_range_width(max_value: Any, min_value: Any) -> float | None:
    max_num = to_numeric(max_value)
    min_num = to_numeric(min_value)
    if max_num is None or min_num is None:
        return None
    return float(max_num) - float(min_num)


def set_formula2_r1c1(cell: xw.Range, formula_r1c1: str) -> bool:
    try:
        cell.formula2 = formula_r1c1
        return True
    except Exception:
        pass
    try:
        cell.api.Formula2R1C1 = formula_r1c1
        return True
    except Exception:
        pass
    try:
        cell.formula = formula_r1c1
        return True
    except Exception:
        return False


def signature_value(value: Any) -> Any:
    numeric = to_numeric(value)
    if numeric is None:
        return number_or_value(value)
    return round(float(numeric), 10)


def extract_empirical_rows(
    workbook: xw.Book,
    sheet: xw.Sheet,
    label: FileLabel,
    source_file: str,
) -> list[dict[str, Any]]:
    matrix, start_row, start_col, end_row, end_col = get_sheet_snapshot(sheet)
    anchor = find_max_anchor(matrix, start_row, start_col)
    if anchor is None:
        print(f"Skipped empirical in {source_file}: missing 'max' anchor")
        return []

    anchor_row, anchor_col = anchor
    n_quarters = 10
    row_start = anchor_row + 1
    row_end = min(end_row, row_start + n_quarters - 1)
    if row_end < row_start:
        return []

    max_col = anchor_col
    min_col = coalesce(
        find_nearby_column(
            matrix,
            start_row,
            start_col,
            anchor_row,
            anchor_col,
            [("min",)],
        ),
        clamp_col(anchor_col + 1, start_col, end_col),
    )
    num_quarters_col = coalesce(
        find_nearby_column(
            matrix,
            start_row,
            start_col,
            anchor_row,
            anchor_col,
            [("num", "quarter"), ("quarter", "used"), ("n", "quarter")],
        ),
        clamp_col(anchor_col - 10, start_col, end_col),
    )
    last_quarter_col = coalesce(
        find_nearby_column(
            matrix,
            start_row,
            start_col,
            anchor_row,
            anchor_col,
            [("last", "quarter"), ("quarter", "last")],
        ),
        clamp_col(anchor_col - 9, start_col, end_col),
    )
    forecast_col = coalesce(
        find_nearby_column(
            matrix,
            start_row,
            start_col,
            anchor_row,
            anchor_col,
            [
                ("estimated", "total", "sold"),
                ("estimate", "total", "sold"),
                ("forecast", "value"),
                ("total", "sold"),
            ],
        ),
        clamp_col(anchor_col - 2, start_col, end_col),
    )
    actual_col = coalesce(
        find_nearby_column(
            matrix,
            start_row,
            start_col,
            anchor_row,
            anchor_col,
            [("reported", "sales"), ("actual", "sales"), ("actual", "value")],
        ),
        clamp_col(anchor_col - 3, start_col, end_col),
    )
    quarterly_sales_col = coalesce(
        find_nearby_column(
            matrix,
            start_row,
            start_col,
            anchor_row,
            anchor_col,
            [("quarterly", "sales"), ("quarter", "sales")],
        ),
        clamp_col(anchor_col - 6, start_col, end_col),
    )
    reported_sales_col = coalesce(
        find_nearby_column(
            matrix,
            start_row,
            start_col,
            anchor_row,
            anchor_col,
            [("reported", "sales"), ("sales", "reported")],
        ),
        actual_col,
    )
    growth_rate_col = coalesce(
        find_nearby_column(
            matrix,
            start_row,
            start_col,
            anchor_row,
            anchor_col,
            [("growth", "rate"), ("growth", "pct"), ("growth", "percent")],
        ),
        clamp_col(anchor_col - 5, start_col, end_col),
    )
    sales_captured_col = coalesce(
        find_nearby_column(
            matrix,
            start_row,
            start_col,
            anchor_row,
            anchor_col,
            [
                ("sales", "captured", "db"),
                ("captured", "db"),
                ("captured", "pct"),
                ("sales", "captured"),
            ],
        ),
        clamp_col(anchor_col - 4, start_col, end_col),
    )
    avg_penetration_col = coalesce(
        find_nearby_column(
            matrix,
            start_row,
            start_col,
            anchor_row,
            anchor_col,
            [
                ("avg", "penetration"),
                ("average", "penetration"),
                ("penetration", "avg"),
            ],
        ),
        sales_captured_col,
    )

    temp_avg_col = end_col + 2
    formula_rows: list[int] = []
    for row in range(row_start, row_end + 1):
        formula = ""
        if avg_penetration_col is not None:
            formula = (
                f'=IFERROR(AVERAGE(R{row_start}C{avg_penetration_col}:R{row}C'
                f"{avg_penetration_col}),\"\")"
            )
        elif reported_sales_col is not None and quarterly_sales_col is not None:
            formula = (
                f'=IFERROR(AVERAGE(R{row_start}C{reported_sales_col}:R{row}C'
                f"{reported_sales_col}/R{row_start}C{quarterly_sales_col}:R{row}C"
                f'{quarterly_sales_col}),"")'
            )
        if formula and set_formula2_r1c1(sheet.cells(row, temp_avg_col), formula):
            formula_rows.append(row)

    if formula_rows:
        workbook.app.calculate()

    max_values = read_column(sheet, max_col, row_start, row_end)
    min_values = read_column(sheet, min_col, row_start, row_end)
    num_quarters_values = read_column(sheet, num_quarters_col, row_start, row_end)
    last_quarter_values = read_column(sheet, last_quarter_col, row_start, row_end)
    forecast_values = read_column(sheet, forecast_col, row_start, row_end)
    actual_values = read_column(sheet, actual_col, row_start, row_end)
    quarterly_values = read_column(sheet, quarterly_sales_col, row_start, row_end)
    reported_values = read_column(sheet, reported_sales_col, row_start, row_end)
    growth_values = read_column(sheet, growth_rate_col, row_start, row_end)
    captured_values = read_column(sheet, sales_captured_col, row_start, row_end)
    avg_values = read_column(sheet, temp_avg_col, row_start, row_end)

    rows: list[dict[str, Any]] = []
    for idx in range(row_end - row_start + 1):
        guard_values = [
            num_quarters_values[idx],
            forecast_values[idx],
            actual_values[idx],
            max_values[idx],
            min_values[idx],
            quarterly_values[idx],
            reported_values[idx],
        ]
        if row_is_blank(guard_values):
            break

        num_quarters_raw = to_numeric(num_quarters_values[idx])
        num_quarters = int(num_quarters_raw) if num_quarters_raw is not None else idx + 1
        forecast_max = to_numeric(max_values[idx])
        forecast_min = to_numeric(min_values[idx])
        forecast_value = to_numeric(forecast_values[idx])
        actual_value = to_numeric(actual_values[idx])
        reported_sales = to_numeric(reported_values[idx])
        quarterly_sales = to_numeric(quarterly_values[idx])
        growth_rate = to_numeric(growth_values[idx])
        sales_captured = to_numeric(captured_values[idx])
        avg_penetration = to_numeric(avg_values[idx])

        if reported_sales is None and actual_value is not None:
            reported_sales = actual_value
        if actual_value is None and reported_sales is not None:
            actual_value = reported_sales

        if avg_penetration is None:
            avg_penetration = sales_captured

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
                "last_quarter_used": number_or_value(last_quarter_values[idx]),
                "forecast_value": forecast_value,
                "actual_value": actual_value,
                "forecast_max": forecast_max,
                "forecast_min": forecast_min,
                "range_width": safe_range_width(forecast_max, forecast_min),
                "avg_penetration_pct": avg_penetration,
                "quarterly_sales": quarterly_sales,
                "reported_sales": reported_sales,
                "growth_rate_pct": growth_rate,
                "sales_captured_in_db_pct": sales_captured,
                "source_file": source_file,
            }
        )

    return rows


def extract_regression_rows(
    workbook: xw.Book,
    sheet: xw.Sheet,
    label: FileLabel,
    source_file: str,
) -> list[dict[str, Any]]:
    matrix, start_row, start_col, end_row, end_col = get_sheet_snapshot(sheet)
    anchor = find_max_anchor(matrix, start_row, start_col)
    if anchor is None:
        print(f"Skipped regression in {source_file}: missing 'max' anchor")
        return []

    anchor_row, anchor_col = anchor
    row_start = anchor_row + 1
    row_end = min(end_row, row_start + 9)
    if row_end < row_start:
        return []

    x_col = clamp_col(anchor_col - 11, start_col, end_col)
    y_col = clamp_col(anchor_col - 7, start_col, end_col)
    if x_col is None or y_col is None:
        print(f"Skipped regression in {source_file}: invalid x/y anchor offsets")
        return []

    max_col = anchor_col
    min_col = coalesce(
        find_nearby_column(
            matrix,
            start_row,
            start_col,
            anchor_row,
            anchor_col,
            [("min",)],
        ),
        clamp_col(anchor_col + 1, start_col, end_col),
    )
    num_quarters_col = coalesce(
        find_nearby_column(
            matrix,
            start_row,
            start_col,
            anchor_row,
            anchor_col,
            [("num", "quarter"), ("quarter", "used"), ("n", "quarter")],
        ),
        clamp_col(x_col - 1, start_col, end_col),
    )
    forecast_col = coalesce(
        find_nearby_column(
            matrix,
            start_row,
            start_col,
            anchor_row,
            anchor_col,
            [
                ("tot", "fcst", "w", "o", "sa"),
                ("tot", "forecast", "without", "sa"),
                ("total", "forecast", "without", "sa"),
                ("fcst", "w", "o", "sa"),
            ],
        ),
        y_col,
    )
    actual_col = find_nearby_column(
        matrix,
        start_row,
        start_col,
        anchor_row,
        anchor_col,
        [("actual", "sales"), ("reported", "sales")],
    )

    temp_intercept_col = end_col + 3
    temp_slope_col = end_col + 4
    formula_rows: list[int] = []

    for row in range(row_start, row_end + 1):
        row_guard = [
            sheet.cells(row, x_col).value,
            sheet.cells(row, y_col).value,
            sheet.cells(row, max_col).value,
            sheet.cells(row, min_col).value if min_col is not None else None,
        ]
        if row_is_blank(row_guard):
            if formula_rows:
                break
            continue

        intercept_formula = (
            f'=IFERROR(INTERCEPT(R{row_start}C{y_col}:R{row}C{y_col},R{row_start}C'
            f'{x_col}:R{row}C{x_col}),"")'
        )
        slope_formula = (
            f'=IFERROR(SLOPE(R{row_start}C{y_col}:R{row}C{y_col},R{row_start}C'
            f'{x_col}:R{row}C{x_col}),"")'
        )
        ok_intercept = set_formula2_r1c1(
            sheet.cells(row, temp_intercept_col), intercept_formula
        )
        ok_slope = set_formula2_r1c1(sheet.cells(row, temp_slope_col), slope_formula)
        if ok_intercept and ok_slope:
            formula_rows.append(row)

    if not formula_rows:
        return []

    workbook.app.calculate()

    rows: list[dict[str, Any]] = []
    previous_signature: tuple[Any, ...] | None = None
    for idx, row in enumerate(formula_rows):
        num_quarters_raw = to_numeric(
            sheet.cells(row, num_quarters_col).value if num_quarters_col else None
        )
        num_quarters = int(num_quarters_raw) if num_quarters_raw is not None else idx + 1
        forecast_value = to_numeric(
            sheet.cells(row, forecast_col).value if forecast_col else None
        )
        actual_value = to_numeric(sheet.cells(row, actual_col).value if actual_col else None)
        forecast_max = to_numeric(sheet.cells(row, max_col).value)
        forecast_min = to_numeric(sheet.cells(row, min_col).value if min_col else None)
        intercept = to_numeric(sheet.cells(row, temp_intercept_col).value)
        slope = to_numeric(sheet.cells(row, temp_slope_col).value)

        current_signature = (
            signature_value(num_quarters),
            signature_value(forecast_value),
            signature_value(forecast_max),
            signature_value(forecast_min),
            signature_value(intercept),
            signature_value(slope),
        )
        if previous_signature is not None and current_signature == previous_signature:
            continue
        previous_signature = current_signature

        rows.append(
            {
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
                "range_width": safe_range_width(forecast_max, forecast_min),
                "intercept": intercept,
                "slope": slope,
                "source_file": source_file,
            }
        )

    return rows


def write_output_sheet(
    worksheet,
    columns: list[str],
    rows: list[dict[str, Any]],
) -> None:
    worksheet.append(columns)
    for row in rows:
        worksheet.append([row.get(column) for column in columns])

    for cell in worksheet[1]:
        cell.font = Font(bold=True)

    worksheet.freeze_panes = "A2"
    if worksheet.max_row >= 1 and worksheet.max_column >= 1:
        worksheet.auto_filter.ref = (
            f"A1:{get_column_letter(worksheet.max_column)}{worksheet.max_row}"
        )

    max_row_for_width_scan = min(worksheet.max_row, 5000)
    for col_idx, column_name in enumerate(columns, start=1):
        max_len = len(column_name)
        for row_idx in range(2, max_row_for_width_scan + 1):
            value = worksheet.cell(row=row_idx, column=col_idx).value
            if value is None:
                continue
            max_len = max(max_len, len(str(value)))
        worksheet.column_dimensions[get_column_letter(col_idx)].width = min(
            max(12, max_len + 2),
            48,
        )


def write_output_workbook(
    destination: Path,
    empirical_rows: list[dict[str, Any]],
    regression_rows: list[dict[str, Any]],
) -> None:
    workbook = Workbook()
    empirical_sheet = workbook.active
    empirical_sheet.title = "empirical_candidates"
    write_output_sheet(empirical_sheet, EMPIRICAL_COLUMNS, empirical_rows)

    regression_sheet = workbook.create_sheet("regression_candidates")
    write_output_sheet(regression_sheet, REGRESSION_COLUMNS, regression_rows)
    workbook.save(destination)


def main() -> None:
    source_dir = Path(input_dir).expanduser()
    target_dir = Path(output_dir).expanduser()

    if not source_dir.exists():
        print(f"Input directory not found: {source_dir}")
        return
    if not source_dir.is_dir():
        print(f"Input path is not a directory: {source_dir}")
        return

    output_path = pick_output_path(source_dir, target_dir)

    candidate_files: list[Path] = []
    for file_path in sorted(source_dir.iterdir()):
        if not file_path.is_file():
            continue
        if file_path.name.startswith("~"):
            print(f"Skipped file {file_path.name}: temporary file")
            continue
        if file_path.suffix.lower() != ".xlsx":
            print(f"Skipped file {file_path.name}: not an .xlsx file")
            continue
        candidate_files.append(file_path)

    empirical_rows: list[dict[str, Any]] = []
    regression_rows: list[dict[str, Any]] = []
    processed_count = 0

    app = xw.App(visible=False, add_book=False)
    app.display_alerts = False
    app.screen_updating = False
    try:
        try:
            app.api.Calculation = -4135  # xlCalculationManual
        except Exception:
            pass

        for file_path in candidate_files:
            print(f"Processing file {file_path.name}")
            workbook: xw.Book | None = None
            try:
                workbook = app.books.open(str(file_path), update_links=False)
                file_label = parse_file_label(file_path.name)

                empirical_sheet = get_sheet_case_insensitive(workbook, "Empirical Model")
                if empirical_sheet is None:
                    print(
                        f"Skipped empirical in {file_path.name}: "
                        "sheet 'Empirical Model' not found"
                    )
                    current_empirical = []
                else:
                    current_empirical = extract_empirical_rows(
                        workbook=workbook,
                        sheet=empirical_sheet,
                        label=file_label,
                        source_file=file_path.name,
                    )

                regression_sheet = get_sheet_case_insensitive(workbook, "Regression Model")
                if regression_sheet is None:
                    print(
                        f"Skipped regression in {file_path.name}: "
                        "sheet 'Regression Model' not found"
                    )
                    current_regression = []
                else:
                    current_regression = extract_regression_rows(
                        workbook=workbook,
                        sheet=regression_sheet,
                        label=file_label,
                        source_file=file_path.name,
                    )

                empirical_rows.extend(current_empirical)
                regression_rows.extend(current_regression)
                processed_count += 1
            except Exception as exc:
                print(f"Skipped file {file_path.name}: {exc}")
            finally:
                safe_close_source_workbook(workbook)
    finally:
        try:
            app.quit()
        except Exception:
            pass

    write_output_workbook(output_path, empirical_rows, regression_rows)
    print(f"Output path: {output_path}")
    print(f"Number of files processed: {processed_count}")
    print(f"Number of empirical rows: {len(empirical_rows)}")
    print(f"Number of regression rows: {len(regression_rows)}")


if __name__ == "__main__":
    main()
