#!/usr/bin/env python3
from __future__ import annotations

from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path
import re
from typing import Any, Callable

import xlwings as xw
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

# User-configurable paths
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

PERIOD_REGEX = re.compile(r"(Early|Mid|Late)\s*([A-Za-z]{3,9})\s*(\d{4})", re.IGNORECASE)
TICKER_REGEX = re.compile(r"\b[A-Z]{1,8}\b")

MONTH_ALIASES = {
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

DAY_BY_PHASE = {"early": 5, "mid": 15, "late": 25}

EMPIRICAL_HEADER_PATTERNS = {
    "num_quarters_used": [r"\bnum(?:ber)?\b.*\bquarter", r"\bn\b.*\bquarter"],
    "last_quarter_used": [r"\blast\b.*\bquarter"],
    "forecast_value": [r"estimated.*total.*sold", r"\bforecast\b"],
    "actual_value": [r"\breported\b.*\bsales\b", r"\bactual\b"],
    "forecast_max": [r"^max$"],
    "forecast_min": [r"^min$"],
    "avg_penetration_pct": [r"\bavg\b.*\bpenetration\b", r"\bpenetration\b"],
    "quarterly_sales": [r"\bquarterly\b.*\bsales\b"],
    "reported_sales": [r"\breported\b.*\bsales\b"],
    "growth_rate_pct": [r"\bgrowth\b"],
    "sales_captured_in_db_pct": [r"\bcaptured\b.*\bdb\b", r"\bin db\b"],
}

REGRESSION_HEADER_PATTERNS = {
    "num_quarters_used": [r"\bnum(?:ber)?\b.*\bquarter", r"\bn\b.*\bquarter"],
    "forecast_value": [r"\btot\b.*\bfcst\b.*\bwo\b.*\bsa\b", r"\bforecast\b"],
    "actual_value": [r"\breported\b.*\bsales\b", r"\bactual\b"],
    "forecast_max": [r"^max$"],
    "forecast_min": [r"^min$"],
}

# Offset fallback when a header cannot be detected.
EMPIRICAL_FALLBACK_OFFSETS = {
    "forecast_max": 0,
    "forecast_min": 1,
    "forecast_value": -1,
    "actual_value": -2,
    "reported_sales": -2,
    "last_quarter_used": -3,
    "quarterly_sales": -4,
    "growth_rate_pct": -5,
    "sales_captured_in_db_pct": -6,
    "avg_penetration_pct": -7,
}

REGRESSION_FALLBACK_OFFSETS = {
    "forecast_max": 0,
    "forecast_min": 1,
    "forecast_value": -1,
    "actual_value": -2,
}


@dataclass
class ModelMetadata:
    model: str
    ticker: str
    model_period: str
    model_date: str


def normalize_text(value: Any) -> str:
    if value is None:
        return ""
    text = str(value).strip()
    text = re.sub(r"\s+", " ", text)
    return text.lower()


def to_float(value: Any) -> float | None:
    if value is None or value == "":
        return None
    if isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        return float(value)

    text = str(value).strip()
    if not text:
        return None

    is_percent = text.endswith("%")
    if is_percent:
        text = text[:-1]

    text = text.replace(",", "")
    try:
        number = float(text)
        return number / 100 if is_percent else number
    except ValueError:
        return None


def to_matrix(values: Any) -> list[list[Any]]:
    if values is None:
        return []
    if isinstance(values, (list, tuple)):
        if not values:
            return []
        first = values[0]
        if isinstance(first, (list, tuple)):
            return [list(row) for row in values]
        return [list(values)]
    return [[values]]


def parse_month(month_text: str) -> int:
    token = re.sub(r"[^A-Za-z]", "", month_text).lower()
    if token in MONTH_ALIASES:
        return MONTH_ALIASES[token]

    for fmt in ("%b", "%B"):
        try:
            return datetime.strptime(month_text, fmt).month
        except ValueError:
            continue

    raise ValueError(f"Unable to parse month token '{month_text}'")


def parse_file_metadata(file_name: str) -> ModelMetadata:
    stem = Path(file_name).stem

    ticker = ""
    parts = [part.strip() for part in stem.split(" - ")]
    if len(parts) >= 2:
        candidate = re.sub(r"[^A-Za-z0-9]", "", parts[1]).upper()
        if re.fullmatch(r"[A-Z]{1,8}", candidate):
            ticker = candidate

    if not ticker:
        upper_stem = stem.upper()
        ticker_match = TICKER_REGEX.search(upper_stem)
        ticker = ticker_match.group(0) if ticker_match else "UNKNOWN"

    period_match = PERIOD_REGEX.search(stem)
    if not period_match:
        model_period = "Unknown_0000"
        model_date = ""
        model = f"{ticker}_{model_period}"
        return ModelMetadata(model=model, ticker=ticker, model_period=model_period, model_date=model_date)

    phase = period_match.group(1).title()
    phase_key = phase.lower()
    month_token = period_match.group(2)
    year = int(period_match.group(3))
    month_num = parse_month(month_token)
    day = DAY_BY_PHASE[phase_key]
    period_month = date(year, month_num, 1).strftime("%b")

    model_period = f"{phase}{period_month}_{year}"
    model_date = date(year, month_num, day).isoformat()
    model = f"{ticker}_{model_period}"

    return ModelMetadata(model=model, ticker=ticker, model_period=model_period, model_date=model_date)


def find_sheet_case_insensitive(workbook: xw.Book, name: str) -> xw.Sheet | None:
    target = name.lower()
    for sheet in workbook.sheets:
        if sheet.name.lower() == target:
            return sheet
    return None


def read_cell(sheet: xw.Sheet, row: int, col: int) -> Any:
    if row < 1 or col < 1:
        return None
    try:
        return sheet.range((row, col)).value
    except Exception:
        return None


def set_formula2_r1c1(cell: xw.Range, formula_r1c1: str) -> None:
    api = cell.api
    if hasattr(api, "Formula2R1C1"):
        api.Formula2R1C1 = formula_r1c1
        return
    if hasattr(api, "FormulaR1C1"):
        api.FormulaR1C1 = formula_r1c1
        return
    if hasattr(api, "Formula2"):
        api.Formula2 = formula_r1c1
        return
    cell.formula = formula_r1c1


def close_workbook_safe(workbook: xw.Book) -> None:
    attempts: list[Callable[[], None]] = [
        lambda: workbook.close(save=False),
        lambda: workbook.close(False),
        lambda: workbook.api.Close(SaveChanges=False),
        lambda: workbook.api.Close(False),
    ]

    last_error: Exception | None = None
    for attempt in attempts:
        try:
            attempt()
            return
        except Exception as error:
            last_error = error

    if last_error:
        raise last_error


def find_anchor_cell(sheet: xw.Sheet, anchor_label: str = "max") -> tuple[int, int] | None:
    used = sheet.used_range
    start_row = used.row
    start_col = used.column
    matrix = to_matrix(used.value)
    target = anchor_label.lower()

    for row_index, row_values in enumerate(matrix):
        for col_index, value in enumerate(row_values):
            if normalize_text(value) == target:
                return start_row + row_index, start_col + col_index
    return None


def map_header_columns(
    sheet: xw.Sheet,
    header_row: int,
    col_start: int,
    col_end: int,
    patterns: dict[str, list[str]],
) -> dict[str, int]:
    result: dict[str, int] = {}

    for col in range(max(1, col_start), max(1, col_end) + 1):
        header_text = normalize_text(read_cell(sheet, header_row, col))
        if not header_text:
            continue
        for key, key_patterns in patterns.items():
            if key in result:
                continue
            if any(re.search(pattern, header_text) for pattern in key_patterns):
                result[key] = col
    return result


def resolve_col(
    header_cols: dict[str, int],
    key: str,
    anchor_col: int,
    fallback_offsets: dict[str, int],
) -> int | None:
    if key in header_cols:
        return header_cols[key]
    fallback = fallback_offsets.get(key)
    if fallback is None:
        return None
    col = anchor_col + fallback
    return col if col > 0 else None


def infer_results_direction(sheet: xw.Sheet, anchor_row: int, max_col: int, min_col: int) -> int:
    down_score = 0
    up_score = 0
    for step in range(1, 11):
        down_row = anchor_row + step
        up_row = anchor_row - step

        down_has_number = to_float(read_cell(sheet, down_row, max_col)) is not None or to_float(
            read_cell(sheet, down_row, min_col)
        ) is not None
        up_has_number = to_float(read_cell(sheet, up_row, max_col)) is not None or to_float(
            read_cell(sheet, up_row, min_col)
        ) is not None

        if down_has_number:
            down_score += 1
        if up_has_number:
            up_score += 1

    return 1 if down_score >= up_score else -1


def round_or_none(value: float | None, digits: int = 10) -> float | None:
    if value is None:
        return None
    return round(value, digits)


def extract_empirical_candidates(
    workbook: xw.Book,
    sheet: xw.Sheet,
    metadata: ModelMetadata,
    source_file: str,
) -> list[dict[str, Any]]:
    anchor = find_anchor_cell(sheet, "max")
    if anchor is None:
        return []

    anchor_row, anchor_col = anchor
    used = sheet.used_range
    first_col = used.column
    last_col = used.column + used.columns.count - 1
    header_cols = map_header_columns(
        sheet=sheet,
        header_row=anchor_row,
        col_start=max(first_col, anchor_col - 25),
        col_end=min(last_col, anchor_col + 25),
        patterns=EMPIRICAL_HEADER_PATTERNS,
    )

    forecast_max_col = resolve_col(header_cols, "forecast_max", anchor_col, EMPIRICAL_FALLBACK_OFFSETS) or anchor_col
    forecast_min_col = (
        resolve_col(header_cols, "forecast_min", anchor_col, EMPIRICAL_FALLBACK_OFFSETS) or (anchor_col + 1)
    )
    forecast_value_col = resolve_col(header_cols, "forecast_value", anchor_col, EMPIRICAL_FALLBACK_OFFSETS)
    actual_value_col = resolve_col(header_cols, "actual_value", anchor_col, EMPIRICAL_FALLBACK_OFFSETS)
    reported_sales_col = resolve_col(header_cols, "reported_sales", anchor_col, EMPIRICAL_FALLBACK_OFFSETS)
    quarterly_sales_col = resolve_col(header_cols, "quarterly_sales", anchor_col, EMPIRICAL_FALLBACK_OFFSETS)
    growth_rate_col = resolve_col(header_cols, "growth_rate_pct", anchor_col, EMPIRICAL_FALLBACK_OFFSETS)
    sales_captured_col = resolve_col(
        header_cols, "sales_captured_in_db_pct", anchor_col, EMPIRICAL_FALLBACK_OFFSETS
    )
    last_quarter_col = resolve_col(header_cols, "last_quarter_used", anchor_col, EMPIRICAL_FALLBACK_OFFSETS)
    num_quarters_col = header_cols.get("num_quarters_used")
    penetration_col = resolve_col(header_cols, "avg_penetration_pct", anchor_col, EMPIRICAL_FALLBACK_OFFSETS)

    direction = infer_results_direction(sheet, anchor_row, forecast_max_col, forecast_min_col)

    temp_calc_cell = sheet.range((anchor_row, last_col + 5))
    rows: list[dict[str, Any]] = []

    for n_quarters in range(1, 11):
        result_row = anchor_row + direction * n_quarters
        if result_row < 1:
            continue

        avg_penetration = None
        hist_end = anchor_row - 1
        hist_start = hist_end - n_quarters + 1
        if penetration_col is not None and hist_start >= 1 and hist_end >= hist_start:
            avg_formula = (
                f'=IFERROR(AVERAGE(R{hist_start}C{penetration_col}:R{hist_end}C{penetration_col}),"")'
            )
            set_formula2_r1c1(temp_calc_cell, avg_formula)
            workbook.app.calculate()
            avg_penetration = to_float(temp_calc_cell.value)

        num_quarters_raw = read_cell(sheet, result_row, num_quarters_col) if num_quarters_col else None
        num_quarters_used = int(to_float(num_quarters_raw) or n_quarters)
        last_quarter_used = read_cell(sheet, result_row, last_quarter_col) if last_quarter_col else ""
        if isinstance(last_quarter_used, (date, datetime)):
            last_quarter_used = last_quarter_used.isoformat()

        forecast_max = to_float(read_cell(sheet, result_row, forecast_max_col))
        forecast_min = to_float(read_cell(sheet, result_row, forecast_min_col))
        forecast_value = to_float(read_cell(sheet, result_row, forecast_value_col)) if forecast_value_col else None
        actual_value = to_float(read_cell(sheet, result_row, actual_value_col)) if actual_value_col else None
        quarterly_sales = to_float(read_cell(sheet, result_row, quarterly_sales_col)) if quarterly_sales_col else None
        reported_sales = to_float(read_cell(sheet, result_row, reported_sales_col)) if reported_sales_col else None
        growth_rate_pct = to_float(read_cell(sheet, result_row, growth_rate_col)) if growth_rate_col else None
        sales_captured_pct = to_float(read_cell(sheet, result_row, sales_captured_col)) if sales_captured_col else None

        if forecast_value is None and forecast_max is not None and forecast_min is not None:
            forecast_value = (forecast_max + forecast_min) / 2
        if actual_value is None and reported_sales is not None:
            actual_value = reported_sales
        if reported_sales is None and actual_value is not None:
            reported_sales = actual_value

        range_width = None
        if forecast_max is not None and forecast_min is not None:
            range_width = forecast_max - forecast_min

        has_signal = any(
            value is not None
            for value in (
                forecast_value,
                actual_value,
                forecast_max,
                forecast_min,
                avg_penetration,
                quarterly_sales,
                reported_sales,
            )
        )
        if not has_signal:
            continue

        rows.append(
            {
                "model": metadata.model,
                "ticker": metadata.ticker,
                "model_period": metadata.model_period,
                "model_date": metadata.model_date,
                "method": "empirical",
                "parameter_name": "avg_penetration_pct",
                "parameter_value": avg_penetration,
                "num_quarters_used": num_quarters_used,
                "last_quarter_used": last_quarter_used,
                "forecast_value": forecast_value,
                "actual_value": actual_value,
                "forecast_max": forecast_max,
                "forecast_min": forecast_min,
                "range_width": range_width,
                "avg_penetration_pct": avg_penetration,
                "quarterly_sales": quarterly_sales,
                "reported_sales": reported_sales,
                "growth_rate_pct": growth_rate_pct,
                "sales_captured_in_db_pct": sales_captured_pct,
                "source_file": source_file,
            }
        )

    temp_calc_cell.clear_contents()
    return rows


def extract_regression_candidates(
    workbook: xw.Book,
    sheet: xw.Sheet,
    metadata: ModelMetadata,
    source_file: str,
) -> list[dict[str, Any]]:
    anchor = find_anchor_cell(sheet, "max")
    if anchor is None:
        return []

    anchor_row, anchor_col = anchor
    used = sheet.used_range
    first_col = used.column
    last_col = used.column + used.columns.count - 1

    header_cols = map_header_columns(
        sheet=sheet,
        header_row=anchor_row,
        col_start=max(first_col, anchor_col - 25),
        col_end=min(last_col, anchor_col + 25),
        patterns=REGRESSION_HEADER_PATTERNS,
    )

    forecast_max_col = resolve_col(header_cols, "forecast_max", anchor_col, REGRESSION_FALLBACK_OFFSETS) or anchor_col
    forecast_min_col = (
        resolve_col(header_cols, "forecast_min", anchor_col, REGRESSION_FALLBACK_OFFSETS) or (anchor_col + 1)
    )
    forecast_value_col = resolve_col(header_cols, "forecast_value", anchor_col, REGRESSION_FALLBACK_OFFSETS)
    actual_value_col = resolve_col(header_cols, "actual_value", anchor_col, REGRESSION_FALLBACK_OFFSETS)
    num_quarters_col = header_cols.get("num_quarters_used")

    y_col = anchor_col - 7
    x_col = anchor_col - 11
    if y_col < 1 or x_col < 1:
        return []

    direction = infer_results_direction(sheet, anchor_row, forecast_max_col, forecast_min_col)

    intercept_cell = sheet.range((anchor_row, last_col + 5))
    slope_cell = sheet.range((anchor_row, last_col + 6))

    data_end_row = anchor_row - 1
    if data_end_row < 1:
        return []

    rows: list[dict[str, Any]] = []
    previous_signature: tuple[Any, ...] | None = None

    for n_quarters in range(1, 11):
        data_start_row = data_end_row - n_quarters + 1
        if data_start_row < 1:
            continue

        intercept_formula = (
            f'=IFERROR(INTERCEPT(R{data_start_row}C{y_col}:R{data_end_row}C{y_col},'
            f'R{data_start_row}C{x_col}:R{data_end_row}C{x_col}),"")'
        )
        slope_formula = (
            f'=IFERROR(SLOPE(R{data_start_row}C{y_col}:R{data_end_row}C{y_col},'
            f'R{data_start_row}C{x_col}:R{data_end_row}C{x_col}),"")'
        )

        set_formula2_r1c1(intercept_cell, intercept_formula)
        set_formula2_r1c1(slope_cell, slope_formula)
        workbook.app.calculate()

        intercept = to_float(intercept_cell.value)
        slope = to_float(slope_cell.value)

        result_row = anchor_row + direction * n_quarters
        if result_row < 1:
            continue

        num_quarters_raw = read_cell(sheet, result_row, num_quarters_col) if num_quarters_col else None
        num_quarters_used = int(to_float(num_quarters_raw) or n_quarters)

        forecast_value = to_float(read_cell(sheet, result_row, forecast_value_col)) if forecast_value_col else None
        if forecast_value is None and intercept is not None and slope is not None:
            next_x = to_float(read_cell(sheet, data_end_row + 1, x_col))
            if next_x is None:
                last_x = to_float(read_cell(sheet, data_end_row, x_col))
                if last_x is not None:
                    next_x = last_x + 1
            if next_x is not None:
                forecast_value = intercept + slope * next_x

        actual_value = to_float(read_cell(sheet, result_row, actual_value_col)) if actual_value_col else None
        forecast_max = to_float(read_cell(sheet, result_row, forecast_max_col))
        forecast_min = to_float(read_cell(sheet, result_row, forecast_min_col))
        if forecast_max is None and forecast_value is not None:
            forecast_max = forecast_value
        if forecast_min is None and forecast_value is not None:
            forecast_min = forecast_value

        range_width = None
        if forecast_max is not None and forecast_min is not None:
            range_width = forecast_max - forecast_min

        signature = (
            num_quarters_used,
            round_or_none(forecast_value),
            round_or_none(intercept),
            round_or_none(slope),
            round_or_none(forecast_max),
            round_or_none(forecast_min),
        )
        if previous_signature == signature:
            continue
        previous_signature = signature

        has_signal = any(value is not None for value in (forecast_value, intercept, slope, forecast_max, forecast_min))
        if not has_signal:
            continue

        rows.append(
            {
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
        )

    intercept_cell.clear_contents()
    slope_cell.clear_contents()
    return rows


def build_output_path(source_input_dir: Path, source_output_dir: Path) -> Path:
    source_output_dir.mkdir(parents=True, exist_ok=True)
    base_name = f"{source_input_dir.name}_PARAM"
    candidate = source_output_dir / f"{base_name}.xlsx"
    counter = 1

    while candidate.exists():
        candidate = source_output_dir / f"{base_name}.{counter}.xlsx"
        counter += 1

    return candidate


def write_sheet(workbook: Workbook, sheet_name: str, columns: list[str], rows: list[dict[str, Any]]) -> None:
    if sheet_name in workbook.sheetnames:
        ws = workbook[sheet_name]
    else:
        ws = workbook.create_sheet(sheet_name)

    ws.append(columns)
    for record in rows:
        ws.append([record.get(column, "") for column in columns])

    for col_index, column_name in enumerate(columns, start=1):
        max_length = len(column_name)
        for row_index in range(2, len(rows) + 2):
            value = ws.cell(row=row_index, column=col_index).value
            if value is None:
                continue
            text = value.isoformat() if isinstance(value, (date, datetime)) else str(value)
            max_length = max(max_length, len(text))
        ws.column_dimensions[get_column_letter(col_index)].width = min(max(max_length + 2, 12), 42)

    for header_cell in ws[1]:
        header_cell.font = Font(bold=True)

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(columns))}{max(len(rows) + 1, 1)}"


def write_output_workbook(
    destination_path: Path,
    empirical_rows: list[dict[str, Any]],
    regression_rows: list[dict[str, Any]],
) -> None:
    out_wb = Workbook()
    default_sheet = out_wb.active
    out_wb.remove(default_sheet)

    write_sheet(out_wb, "empirical_candidates", EMPIRICAL_COLUMNS, empirical_rows)
    write_sheet(out_wb, "regression_candidates", REGRESSION_COLUMNS, regression_rows)
    out_wb.save(destination_path)


def main() -> None:
    source_input_dir = Path(input_dir).expanduser().resolve()
    source_output_dir = Path(output_dir).expanduser().resolve()

    if not source_input_dir.exists() or not source_input_dir.is_dir():
        raise FileNotFoundError(f"Input directory does not exist or is not a directory: {source_input_dir}")

    output_path = build_output_path(source_input_dir, source_output_dir)

    empirical_rows: list[dict[str, Any]] = []
    regression_rows: list[dict[str, Any]] = []
    processed_file_count = 0

    app = xw.App(visible=False, add_book=False)
    app.display_alerts = False
    app.screen_updating = False
    try:
        app.calculation = "manual"
    except Exception:
        pass

    try:
        for file_path in sorted(source_input_dir.iterdir(), key=lambda path: path.name.lower()):
            if not file_path.is_file():
                print(f"Skipped {file_path.name}: not a file")
                continue
            if file_path.name.startswith("~"):
                print(f"Skipped {file_path.name}: temp file")
                continue
            if file_path.suffix.lower() != ".xlsx":
                print(f"Skipped {file_path.name}: not an .xlsx file")
                continue

            print(f"Processed {file_path.name}")
            try:
                workbook = app.books.open(str(file_path), update_links=False)
            except Exception as error:
                print(f"Skipped {file_path.name}: failed to open ({error})")
                continue

            try:
                metadata = parse_file_metadata(file_path.name)

                empirical_sheet = find_sheet_case_insensitive(workbook, "Empirical Model")
                if empirical_sheet is None:
                    print(f"Skipped empirical extraction for {file_path.name}: sheet 'Empirical Model' not found")
                else:
                    empirical_rows.extend(
                        extract_empirical_candidates(
                            workbook=workbook,
                            sheet=empirical_sheet,
                            metadata=metadata,
                            source_file=file_path.name,
                        )
                    )

                regression_sheet = find_sheet_case_insensitive(workbook, "Regression Model")
                if regression_sheet is None:
                    print(f"Skipped regression extraction for {file_path.name}: sheet 'Regression Model' not found")
                else:
                    regression_rows.extend(
                        extract_regression_candidates(
                            workbook=workbook,
                            sheet=regression_sheet,
                            metadata=metadata,
                            source_file=file_path.name,
                        )
                    )

                processed_file_count += 1
            except Exception as error:
                print(f"Skipped {file_path.name}: processing error ({error})")
            finally:
                try:
                    close_workbook_safe(workbook)
                except Exception as close_error:
                    print(f"Skipped {file_path.name}: failed to close safely ({close_error})")
    finally:
        app.quit()

    write_output_workbook(output_path, empirical_rows, regression_rows)

    print(f"Output path: {output_path}")
    print(f"Number of files processed: {processed_file_count}")
    print(f"Number of empirical rows: {len(empirical_rows)}")
    print(f"Number of regression rows: {len(regression_rows)}")


if __name__ == "__main__":
    main()
