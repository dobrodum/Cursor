#!/usr/bin/env python3
"""Extract empirical and regression model candidates into one workbook."""

from __future__ import annotations

import re
from dataclasses import dataclass
from datetime import date
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Sequence, Tuple

import xlwings as xw
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

# ----------------------------
# User-configurable locations.
# ----------------------------
input_dir = "/path/to/input"
output_dir = "/path/to/output"

N_QUARTERS = 10
DAY_BY_STAGE = {"early": 5, "mid": 15, "late": 25}
MONTH_BY_ABBR = {
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


@dataclass
class FileMetadata:
    model: str
    ticker: str
    model_period: str
    model_date: str


def is_number(value: Any) -> bool:
    return isinstance(value, (int, float)) and not isinstance(value, bool)


def as_float(value: Any) -> Optional[float]:
    if is_number(value):
        return float(value)
    return None


def normalize_text(value: str) -> str:
    return re.sub(r"\s+", " ", re.sub(r"[^a-z0-9]+", " ", value.lower())).strip()


def ensure_2d(values: Any) -> List[List[Any]]:
    if values is None:
        return []
    if not isinstance(values, list):
        return [[values]]
    if values and not isinstance(values[0], list):
        return [values]
    return values


def matrix_value(
    matrix: List[List[Any]],
    start_row: int,
    start_col: int,
    row: Optional[int],
    col: Optional[int],
) -> Any:
    if row is None or col is None:
        return None
    r_idx = row - start_row
    c_idx = col - start_col
    if r_idx < 0 or c_idx < 0 or r_idx >= len(matrix):
        return None
    row_values = matrix[r_idx]
    if c_idx >= len(row_values):
        return None
    return row_values[c_idx]


def parse_file_metadata(file_name: str) -> FileMetadata:
    stem = Path(file_name).stem
    parts = [piece.strip() for piece in stem.split(" - ")]
    ticker = parts[1] if len(parts) > 1 and parts[1] else "UNKNOWN"

    period_source = parts[2] if len(parts) > 2 else stem
    period_source = period_source.split("_")[0]
    match = re.search(r"(?i)\b(early|mid|late)([a-z]{3,9})(\d{4})\b", period_source)

    model_period = "Unknown_Period"
    model_date = ""
    if match:
        stage = match.group(1).lower()
        month_token = match.group(2).lower()
        year = int(match.group(3))
        month = MONTH_BY_ABBR.get(month_token[:3])
        if month is not None:
            model_period = f"{stage.title()}{month_token[:3].title()}_{year}"
            model_date = date(year, month, DAY_BY_STAGE[stage]).isoformat()

    model = f"{ticker}_{model_period}"
    return FileMetadata(model=model, ticker=ticker, model_period=model_period, model_date=model_date)


def build_output_path(input_path: Path, output_path: Path) -> Path:
    output_path.mkdir(parents=True, exist_ok=True)
    folder_name = input_path.name or "input"
    base_name = f"{folder_name}_PARAM.xlsx"
    candidate = output_path / base_name
    if not candidate.exists():
        return candidate

    counter = 1
    while True:
        candidate = output_path / f"{folder_name}_PARAM.{counter}.xlsx"
        if not candidate.exists():
            return candidate
        counter += 1


def set_formula2(range_obj: Any, formula: str) -> None:
    try:
        range_obj.formula2 = formula
    except Exception:
        range_obj.formula = formula


def safe_close_workbook(wb: Any) -> None:
    close_attempts = (
        lambda: wb.close(save=False),
        lambda: wb.close(False),
        lambda: wb.api.Close(SaveChanges=False),
        lambda: wb.api.Close(False),
    )
    for close_fn in close_attempts:
        try:
            close_fn()
            return
        except Exception:
            continue


def find_anchor_max(
    matrix: List[List[Any]], start_row: int, start_col: int
) -> Optional[Tuple[int, int]]:
    for r_idx, row_values in enumerate(matrix):
        for c_idx, value in enumerate(row_values):
            if isinstance(value, str) and value.strip().lower() == "max":
                return start_row + r_idx, start_col + c_idx
    return None


def collect_text_cells(
    matrix: List[List[Any]], start_row: int, start_col: int
) -> List[Tuple[int, int, str]]:
    cells: List[Tuple[int, int, str]] = []
    for r_idx, row_values in enumerate(matrix):
        for c_idx, value in enumerate(row_values):
            if isinstance(value, str):
                normalized = normalize_text(value)
                if normalized:
                    cells.append((start_row + r_idx, start_col + c_idx, normalized))
    return cells


def find_row_by_keywords(
    text_cells: Sequence[Tuple[int, int, str]],
    include_tokens: Sequence[str],
    exclude_tokens: Sequence[str] = (),
    anchor_row: Optional[int] = None,
    anchor_col: Optional[int] = None,
) -> Optional[int]:
    best_row = None
    best_score = float("inf")

    include = [token.lower() for token in include_tokens]
    exclude = [token.lower() for token in exclude_tokens]

    for row, col, text in text_cells:
        words = text.split()
        if not all(any(word == token or word.startswith(token) for word in words) for token in include):
            continue
        if any(any(word == token or word.startswith(token) for word in words) for token in exclude):
            continue

        row_gap = abs(row - anchor_row) if anchor_row is not None else 0
        col_gap = abs(col - anchor_col) if anchor_col is not None else 0
        score = (row_gap * 5) + col_gap
        if score < best_score:
            best_score = score
            best_row = row

    return best_row


def infer_scenario_direction(
    matrix: List[List[Any]],
    start_row: int,
    start_col: int,
    anchor_row: int,
    anchor_col: int,
) -> int:
    right_hits = 0
    left_hits = 0
    for step in range(1, N_QUARTERS + 1):
        right_val = matrix_value(matrix, start_row, start_col, anchor_row, anchor_col + step)
        left_val = matrix_value(matrix, start_row, start_col, anchor_row, anchor_col - step)
        if is_number(right_val):
            right_hits += 1
        if is_number(left_val):
            left_hits += 1
    return -1 if left_hits > right_hits else 1


def build_scenario_columns(
    anchor_col: int, direction: int, min_col: int, max_col: int
) -> List[int]:
    columns: List[int] = []
    for step in range(1, N_QUARTERS + 1):
        col = anchor_col + (direction * step)
        if min_col <= col <= max_col:
            columns.append(col)
    return columns


def row_numeric_columns(
    matrix: List[List[Any]], start_row: int, start_col: int, row: Optional[int]
) -> List[int]:
    if row is None:
        return []
    r_idx = row - start_row
    if r_idx < 0 or r_idx >= len(matrix):
        return []

    numeric_cols: List[int] = []
    for c_idx, value in enumerate(matrix[r_idx]):
        if is_number(value):
            numeric_cols.append(start_col + c_idx)
    return numeric_cols


def calc_range_width(max_value: Any, min_value: Any) -> Optional[float]:
    f_max = as_float(max_value)
    f_min = as_float(min_value)
    if f_max is None or f_min is None:
        return None
    return f_max - f_min


def get_sheet_matrix(sheet: Any) -> Tuple[List[List[Any]], int, int, int, int]:
    used_range = sheet.used_range
    matrix = ensure_2d(used_range.value)
    start_row = used_range.row
    start_col = used_range.column
    max_row = start_row + len(matrix) - 1 if matrix else start_row
    max_col = start_col + max((len(r) for r in matrix), default=0) - 1
    return matrix, start_row, start_col, max_row, max_col


def extract_empirical_rows(wb: Any, metadata: FileMetadata, source_file: str) -> List[Dict[str, Any]]:
    try:
        sheet = wb.sheets["Empirical Model"]
    except Exception:
        print(f"Skipped empirical extraction for {source_file}: missing 'Empirical Model' sheet")
        return []

    matrix, start_row, start_col, max_row, max_col = get_sheet_matrix(sheet)
    if not matrix:
        print(f"Skipped empirical extraction for {source_file}: empty used range")
        return []

    anchor = find_anchor_max(matrix, start_row, start_col)
    if anchor is None:
        print(f"Skipped empirical extraction for {source_file}: could not find 'max' anchor")
        return []
    anchor_row, anchor_col = anchor

    direction = infer_scenario_direction(matrix, start_row, start_col, anchor_row, anchor_col)
    scenario_cols = build_scenario_columns(anchor_col, direction, start_col, max_col)
    if not scenario_cols:
        print(f"Skipped empirical extraction for {source_file}: no scenario columns near anchor")
        return []

    text_cells = collect_text_cells(matrix, start_row, start_col)
    min_row = find_row_by_keywords(
        text_cells,
        include_tokens=("min",),
        anchor_row=anchor_row,
        anchor_col=anchor_col,
    )
    estimated_total_row = (
        find_row_by_keywords(
            text_cells,
            include_tokens=("estimated", "total", "sold"),
            anchor_row=anchor_row,
            anchor_col=anchor_col,
        )
        or find_row_by_keywords(
            text_cells,
            include_tokens=("tot", "fcst"),
            anchor_row=anchor_row,
            anchor_col=anchor_col,
        )
    )
    reported_sales_row = find_row_by_keywords(
        text_cells,
        include_tokens=("reported", "sales"),
        anchor_row=anchor_row,
        anchor_col=anchor_col,
    )
    quarterly_sales_row = find_row_by_keywords(
        text_cells,
        include_tokens=("quarterly", "sales"),
        anchor_row=anchor_row,
        anchor_col=anchor_col,
    )
    growth_rate_row = find_row_by_keywords(
        text_cells,
        include_tokens=("growth", "rate"),
        anchor_row=anchor_row,
        anchor_col=anchor_col,
    )
    sales_captured_row = (
        find_row_by_keywords(
            text_cells,
            include_tokens=("sales", "captured", "db"),
            anchor_row=anchor_row,
            anchor_col=anchor_col,
        )
        or find_row_by_keywords(
            text_cells,
            include_tokens=("captured", "db"),
            anchor_row=anchor_row,
            anchor_col=anchor_col,
        )
    )
    penetration_history_row = (
        find_row_by_keywords(
            text_cells,
            include_tokens=("penetration",),
            exclude_tokens=("avg",),
            anchor_row=anchor_row,
            anchor_col=anchor_col,
        )
        or find_row_by_keywords(
            text_cells,
            include_tokens=("penetration",),
            anchor_row=anchor_row,
            anchor_col=anchor_col,
        )
    )

    historical_cols = row_numeric_columns(matrix, start_row, start_col, penetration_history_row)
    if direction == 1:
        historical_cols = [col for col in historical_cols if col < anchor_col]
    else:
        historical_cols = [col for col in historical_cols if col > anchor_col]
    historical_cols = sorted(historical_cols, key=lambda col: abs(col - anchor_col))

    temp_row = max_row + 2
    temp_col = max_col + 2
    avg_cell = sheet.cells(temp_row, temp_col)

    rows: List[Dict[str, Any]] = []
    for index, scenario_col in enumerate(scenario_cols, start=1):
        if index > N_QUARTERS:
            break

        avg_penetration = None
        last_quarter_used = None
        if penetration_history_row is not None and len(historical_cols) >= index:
            selected = historical_cols[:index]
            start_hist_col = min(selected)
            end_hist_col = max(selected)
            avg_formula = (
                f"=AVERAGE(R{penetration_history_row}C{start_hist_col}:"
                f"R{penetration_history_row}C{end_hist_col})"
            )
            set_formula2(avg_cell, avg_formula)
            wb.app.calculate()
            avg_penetration = avg_cell.value

            header_row = penetration_history_row - 1
            boundary_col = selected[-1]
            last_quarter_used = matrix_value(matrix, start_row, start_col, header_row, boundary_col)

        forecast_max = matrix_value(matrix, start_row, start_col, anchor_row, scenario_col)
        forecast_min = matrix_value(
            matrix, start_row, start_col, min_row if min_row is not None else anchor_row + 1, scenario_col
        )
        forecast_value = matrix_value(matrix, start_row, start_col, estimated_total_row, scenario_col)
        reported_sales = matrix_value(matrix, start_row, start_col, reported_sales_row, scenario_col)
        quarterly_sales = matrix_value(matrix, start_row, start_col, quarterly_sales_row, scenario_col)
        growth_rate = matrix_value(matrix, start_row, start_col, growth_rate_row, scenario_col)
        sales_captured = matrix_value(matrix, start_row, start_col, sales_captured_row, scenario_col)
        range_width = calc_range_width(forecast_max, forecast_min)

        rows.append(
            {
                "model": metadata.model,
                "ticker": metadata.ticker,
                "model_period": metadata.model_period,
                "model_date": metadata.model_date,
                "method": "empirical",
                "parameter_name": "avg_penetration_pct",
                "parameter_value": avg_penetration,
                "num_quarters_used": index,
                "last_quarter_used": last_quarter_used,
                "forecast_value": forecast_value,
                "actual_value": reported_sales,
                "forecast_max": forecast_max,
                "forecast_min": forecast_min,
                "range_width": range_width,
                "avg_penetration_pct": avg_penetration,
                "quarterly_sales": quarterly_sales,
                "reported_sales": reported_sales,
                "growth_rate_pct": growth_rate,
                "sales_captured_in_db_pct": sales_captured,
                "source_file": source_file,
            }
        )

    try:
        avg_cell.value = None
    except Exception:
        pass
    return rows


def collect_xy_rows(
    matrix: List[List[Any]], start_row: int, start_col: int, x_col: int, y_col: int
) -> List[int]:
    rows: List[int] = []
    for row in range(start_row, start_row + len(matrix)):
        x_value = matrix_value(matrix, start_row, start_col, row, x_col)
        y_value = matrix_value(matrix, start_row, start_col, row, y_col)
        if is_number(x_value) and is_number(y_value):
            rows.append(row)
    return rows


def signature_value(value: Any) -> Any:
    if is_number(value):
        return round(float(value), 10)
    return value


def extract_regression_rows(wb: Any, metadata: FileMetadata, source_file: str) -> List[Dict[str, Any]]:
    try:
        sheet = wb.sheets["Regression Model"]
    except Exception:
        print(f"Skipped regression extraction for {source_file}: missing 'Regression Model' sheet")
        return []

    matrix, start_row, start_col, max_row, max_col = get_sheet_matrix(sheet)
    if not matrix:
        print(f"Skipped regression extraction for {source_file}: empty used range")
        return []

    anchor = find_anchor_max(matrix, start_row, start_col)
    if anchor is None:
        print(f"Skipped regression extraction for {source_file}: could not find 'max' anchor")
        return []
    anchor_row, anchor_col = anchor

    y_col = anchor_col - 7
    x_col = anchor_col - 11
    numeric_rows = collect_xy_rows(matrix, start_row, start_col, x_col, y_col)
    if len(numeric_rows) < 2:
        print(f"Skipped regression extraction for {source_file}: not enough x/y numeric rows")
        return []

    direction = infer_scenario_direction(matrix, start_row, start_col, anchor_row, anchor_col)
    scenario_cols = build_scenario_columns(anchor_col, direction, start_col, max_col)
    if not scenario_cols:
        print(f"Skipped regression extraction for {source_file}: no scenario columns near anchor")
        return []

    text_cells = collect_text_cells(matrix, start_row, start_col)
    min_row = find_row_by_keywords(
        text_cells,
        include_tokens=("min",),
        anchor_row=anchor_row,
        anchor_col=anchor_col,
    )
    forecast_total_row = (
        find_row_by_keywords(
            text_cells,
            include_tokens=("tot", "fcst", "sa"),
            anchor_row=anchor_row,
            anchor_col=anchor_col,
        )
        or find_row_by_keywords(
            text_cells,
            include_tokens=("total", "forecast", "sa"),
            anchor_row=anchor_row,
            anchor_col=anchor_col,
        )
    )
    actual_row = find_row_by_keywords(
        text_cells,
        include_tokens=("reported", "sales"),
        anchor_row=anchor_row,
        anchor_col=anchor_col,
    )

    temp_row = max_row + 2
    temp_col = max_col + 2
    intercept_cell = sheet.cells(temp_row, temp_col)
    slope_cell = sheet.cells(temp_row + 1, temp_col)

    rows: List[Dict[str, Any]] = []
    previous_signature: Optional[Tuple[Any, ...]] = None
    max_n = min(N_QUARTERS, len(numeric_rows), len(scenario_cols))
    for num_quarters_used in range(2, max_n + 1):
        start_data_row = numeric_rows[-num_quarters_used]
        end_data_row = numeric_rows[-1]

        intercept_formula = (
            f"=INTERCEPT(R{start_data_row}C{y_col}:R{end_data_row}C{y_col},"
            f"R{start_data_row}C{x_col}:R{end_data_row}C{x_col})"
        )
        slope_formula = (
            f"=SLOPE(R{start_data_row}C{y_col}:R{end_data_row}C{y_col},"
            f"R{start_data_row}C{x_col}:R{end_data_row}C{x_col})"
        )
        set_formula2(intercept_cell, intercept_formula)
        set_formula2(slope_cell, slope_formula)
        wb.app.calculate()

        intercept = intercept_cell.value
        slope = slope_cell.value

        scenario_col = scenario_cols[num_quarters_used - 1]
        forecast_total_without_sa = matrix_value(matrix, start_row, start_col, forecast_total_row, scenario_col)
        forecast_max = matrix_value(matrix, start_row, start_col, anchor_row, scenario_col)
        forecast_min = matrix_value(
            matrix, start_row, start_col, min_row if min_row is not None else anchor_row + 1, scenario_col
        )
        actual_value = matrix_value(matrix, start_row, start_col, actual_row, scenario_col)
        range_width = calc_range_width(forecast_max, forecast_min)

        current_signature = (
            signature_value(num_quarters_used),
            signature_value(forecast_total_without_sa),
            signature_value(forecast_max),
            signature_value(forecast_min),
            signature_value(intercept),
            signature_value(slope),
        )
        if previous_signature == current_signature:
            continue
        previous_signature = current_signature

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

    for cell in (intercept_cell, slope_cell):
        try:
            cell.value = None
        except Exception:
            pass
    return rows


def write_sheet(ws: Any, columns: Sequence[str], rows: Sequence[Dict[str, Any]]) -> None:
    ws.append(list(columns))
    for row in rows:
        ws.append([row.get(column) for column in columns])

    for cell in ws[1]:
        cell.font = Font(bold=True)

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions

    for col_idx, column in enumerate(columns, start=1):
        max_length = len(column)
        for row in rows:
            value = row.get(column)
            text = "" if value is None else str(value)
            if len(text) > max_length:
                max_length = len(text)
        ws.column_dimensions[get_column_letter(col_idx)].width = min(max(max_length + 2, 12), 42)


def write_output_workbook(
    output_path: Path, empirical_rows: Sequence[Dict[str, Any]], regression_rows: Sequence[Dict[str, Any]]
) -> None:
    wb_out = Workbook()
    ws_empirical = wb_out.active
    ws_empirical.title = "empirical_candidates"
    ws_regression = wb_out.create_sheet("regression_candidates")

    write_sheet(ws_empirical, EMPIRICAL_COLUMNS, empirical_rows)
    write_sheet(ws_regression, REGRESSION_COLUMNS, regression_rows)
    wb_out.save(output_path)


def iter_candidate_files(input_path: Path) -> Iterable[Path]:
    for file_path in sorted(input_path.iterdir()):
        if not file_path.is_file():
            continue
        if file_path.name.startswith("~"):
            print(f"Skipped {file_path.name}: temp file")
            continue
        if file_path.suffix.lower() != ".xlsx":
            print(f"Skipped {file_path.name}: not an .xlsx file")
            continue
        yield file_path


def main() -> None:
    input_path = Path(input_dir).expanduser().resolve()
    output_path = Path(output_dir).expanduser().resolve()
    if not input_path.exists() or not input_path.is_dir():
        raise FileNotFoundError(f"input_dir does not exist or is not a directory: {input_path}")

    output_file = build_output_path(input_path, output_path)
    empirical_rows: List[Dict[str, Any]] = []
    regression_rows: List[Dict[str, Any]] = []
    processed_files = 0

    app = xw.App(visible=False, add_book=False)
    try:
        app.display_alerts = False
        app.screen_updating = False
    except Exception:
        pass

    try:
        for file_path in iter_candidate_files(input_path):
            wb = None
            try:
                wb = app.books.open(str(file_path), update_links=False)
                metadata = parse_file_metadata(file_path.name)

                before_empirical = len(empirical_rows)
                before_regression = len(regression_rows)
                empirical_rows.extend(extract_empirical_rows(wb, metadata, file_path.name))
                regression_rows.extend(extract_regression_rows(wb, metadata, file_path.name))
                processed_files += 1

                print(
                    f"Processed {file_path.name}: "
                    f"{len(empirical_rows) - before_empirical} empirical rows, "
                    f"{len(regression_rows) - before_regression} regression rows"
                )
            except Exception as exc:
                print(f"Skipped {file_path.name}: {exc}")
            finally:
                if wb is not None:
                    safe_close_workbook(wb)
    finally:
        app.quit()

    write_output_workbook(output_file, empirical_rows, regression_rows)

    print(f"Output path: {output_file}")
    print(f"Number of files processed: {processed_files}")
    print(f"Number of empirical rows: {len(empirical_rows)}")
    print(f"Number of regression rows: {len(regression_rows)}")


if __name__ == "__main__":
    main()
