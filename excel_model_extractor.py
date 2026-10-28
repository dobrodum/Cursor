#!/usr/bin/env python3
"""
Extract empirical and regression candidate rows from all .xlsx files in one folder.

The script:
1) Opens each source workbook exactly once in one hidden Excel app session.
2) Processes both "Empirical Model" and "Regression Model" while open.
3) Never saves source workbooks.
4) Writes one output workbook with:
   - empirical_candidates
   - regression_candidates
"""

from __future__ import annotations

import datetime as dt
import re
from pathlib import Path
from typing import Any, Iterable

import xlwings as xw
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

# -----------------------------
# User-configurable input/output
# -----------------------------
input_dir = "/path/to/input"
output_dir = "/path/to/output"

# -----------------------------
# Constants
# -----------------------------
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

MODEL_FILE_PATTERN = re.compile(
    r"""
    ^.*?-\s*
    (?P<ticker>[A-Za-z0-9]+)\s*-\s*
    (?P<phase>Early|Mid|Late)
    (?P<month>[A-Za-z]{3,9})
    (?P<year>\d{4})
    _Send\.xlsx$
    """,
    re.IGNORECASE | re.VERBOSE,
)

PHASE_DAY = {
    "early": 5,
    "mid": 15,
    "late": 25,
}

MONTH_NUM = {
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


def log(message: str) -> None:
    print(message, flush=True)


def normalize_text(value: Any) -> str:
    if value is None:
        return ""
    return re.sub(r"\s+", " ", str(value).strip().lower())


def is_blank(value: Any) -> bool:
    return value is None or (isinstance(value, str) and value.strip() == "")


def ensure_2d(values: Any) -> list[list[Any]]:
    if values is None:
        return []
    if not isinstance(values, list):
        return [[values]]
    if not values:
        return []
    if isinstance(values[0], list):
        return values
    return [values]


def ensure_1d(values: Any) -> list[Any]:
    if values is None:
        return []
    if not isinstance(values, list):
        return [values]
    if values and isinstance(values[0], list):
        return [row[0] if row else None for row in values]
    return values


def numeric_diff(left: Any, right: Any) -> Any:
    try:
        if left is None or right is None:
            return None
        return float(left) - float(right)
    except Exception:
        return None


def signature_value(value: Any) -> Any:
    if isinstance(value, (int, float)):
        return round(float(value), 10)
    if value is None:
        return None
    return str(value).strip().lower()


def parse_file_metadata(file_name: str) -> dict[str, Any] | None:
    match = MODEL_FILE_PATTERN.match(file_name)
    if not match:
        return None

    ticker = match.group("ticker").upper()
    phase = match.group("phase").title()
    month_token = match.group("month")
    year = match.group("year")

    month_key = month_token[:3].lower()
    month_num = MONTH_NUM.get(month_key)
    if month_num is None:
        return None

    day = PHASE_DAY.get(phase.lower())
    if day is None:
        return None

    model_period = f"{phase}{month_token[:3].title()}_{year}"
    model_date = dt.date(int(year), month_num, day).isoformat()
    model = f"{ticker}_{model_period}"

    return {
        "ticker": ticker,
        "model_period": model_period,
        "model_date": model_date,
        "model": model,
    }


def resolve_output_path(input_path: Path, output_path: Path) -> Path:
    input_folder_name = input_path.name
    base_name = f"{input_folder_name}_PARAM.xlsx"
    candidate = output_path / base_name
    if not candidate.exists():
        return candidate

    idx = 1
    while True:
        candidate = output_path / f"{input_folder_name}_PARAM.{idx}.xlsx"
        if not candidate.exists():
            return candidate
        idx += 1


def close_workbook_safely(wb: xw.Book) -> None:
    try:
        wb.close(save=False)
        return
    except TypeError:
        pass
    except Exception:
        pass

    try:
        wb.api.Close(False)
        return
    except Exception:
        pass

    try:
        wb.close()
    except Exception:
        pass


def set_formula2(cell: xw.Range, formula: str) -> None:
    try:
        cell.formula2 = formula
    except Exception:
        cell.formula = formula


def build_sheet_context(sheet: xw.Sheet) -> dict[str, Any]:
    used = sheet.used_range
    matrix = ensure_2d(used.value)
    if not matrix:
        matrix = [[None]]

    max_width = max(len(row) for row in matrix)
    matrix = [row + [None] * (max_width - len(row)) for row in matrix]

    start_row = used.row
    start_col = used.column
    end_row = start_row + len(matrix) - 1
    end_col = start_col + max_width - 1
    return {
        "matrix": matrix,
        "start_row": start_row,
        "start_col": start_col,
        "end_row": end_row,
        "end_col": end_col,
    }


def find_anchor_cell(context: dict[str, Any], target: str = "max") -> tuple[int, int] | None:
    matrix = context["matrix"]
    start_row = context["start_row"]
    start_col = context["start_col"]
    target_norm = target.lower()

    candidates: list[tuple[int, int, int]] = []
    for r_idx, row in enumerate(matrix):
        for c_idx, value in enumerate(row):
            if normalize_text(value) == target_norm:
                score = 0
                for look_right in range(1, 4):
                    right_idx = c_idx + look_right
                    if right_idx < len(row) and normalize_text(row[right_idx]) == "min":
                        score += 1
                        break
                candidates.append((score, start_row + r_idx, start_col + c_idx))

    if not candidates:
        return None

    candidates.sort(key=lambda item: (-item[0], item[1], item[2]))
    _, row, col = candidates[0]
    return row, col


def get_row_values(context: dict[str, Any], absolute_row: int) -> list[Any]:
    idx = absolute_row - context["start_row"]
    matrix = context["matrix"]
    if idx < 0 or idx >= len(matrix):
        return []
    return matrix[idx]


def find_column_by_patterns(
    row_values: list[Any],
    row_start_col: int,
    patterns: Iterable[str],
    default: int | None = None,
) -> int | None:
    pattern_norm = [p.lower() for p in patterns]
    for offset, value in enumerate(row_values):
        text = normalize_text(value)
        if text and any(pattern in text for pattern in pattern_norm):
            return row_start_col + offset
    return default


def read_block(
    sheet: xw.Sheet,
    start_row: int,
    end_row: int,
    start_col: int,
    end_col: int,
) -> list[list[Any]]:
    if start_col > end_col or start_row > end_row:
        return []
    values = sheet.range((start_row, start_col), (end_row, end_col)).value
    matrix = ensure_2d(values)
    if not matrix:
        matrix = [[None] * (end_col - start_col + 1) for _ in range(end_row - start_row + 1)]
    width = end_col - start_col + 1
    matrix = [row + [None] * (width - len(row)) for row in matrix]
    return matrix


def block_get(
    block: list[list[Any]],
    row_offset: int,
    absolute_col: int | None,
    block_start_col: int,
    block_end_col: int,
) -> Any:
    if absolute_col is None:
        return None
    if absolute_col < block_start_col or absolute_col > block_end_col:
        return None
    if row_offset < 0 or row_offset >= len(block):
        return None
    row = block[row_offset]
    idx = absolute_col - block_start_col
    if idx < 0 or idx >= len(row):
        return None
    return row[idx]


def extract_empirical_candidates(
    wb: xw.Book,
    meta: dict[str, Any],
    source_file: str,
) -> list[dict[str, Any]]:
    rows: list[dict[str, Any]] = []
    try:
        sheet = wb.sheets["Empirical Model"]
    except Exception:
        log(f"skipped empirical extraction for {source_file} (missing sheet: Empirical Model)")
        return rows

    context = build_sheet_context(sheet)
    anchor = find_anchor_cell(context, "max")
    if not anchor:
        log(f"skipped empirical extraction for {source_file} (anchor 'max' not found)")
        return rows

    anchor_row, anchor_col = anchor
    header_values = get_row_values(context, anchor_row)
    header_start_col = context["start_col"]

    max_col = anchor_col
    min_col = find_column_by_patterns(header_values, header_start_col, [" min ", "minimum", "forecast_min"], default=anchor_col + 1)
    forecast_col = find_column_by_patterns(
        header_values,
        header_start_col,
        ["estimated total sold", "forecast value", "forecast_total", "tot fcst", "forecast"],
        default=anchor_col - 1,
    )
    actual_col = find_column_by_patterns(
        header_values,
        header_start_col,
        ["reported sales", "actual value", "actual"],
        default=anchor_col - 2,
    )
    avg_pen_col = find_column_by_patterns(
        header_values,
        header_start_col,
        ["avg penetration", "average penetration", "penetration pct", "penetration %"],
        default=None,
    )
    quarterly_col = find_column_by_patterns(
        header_values,
        header_start_col,
        ["quarterly sales", "qtr sales", "quarter sales"],
        default=anchor_col - 4,
    )
    reported_col = find_column_by_patterns(
        header_values,
        header_start_col,
        ["reported sales", "reported"],
        default=actual_col,
    )
    growth_col = find_column_by_patterns(
        header_values,
        header_start_col,
        ["growth rate", "growth"],
        default=anchor_col - 5,
    )
    captured_col = find_column_by_patterns(
        header_values,
        header_start_col,
        ["sales captured in db", "captured in db", "captured"],
        default=anchor_col - 6,
    )
    last_quarter_col = find_column_by_patterns(
        header_values,
        header_start_col,
        ["last quarter used", "last quarter", "quarter used"],
        default=anchor_col - 7,
    )
    num_quarters_col = find_column_by_patterns(
        header_values,
        header_start_col,
        ["num quarters used", "num_quarters_used", "n quarters", "quarters used"],
        default=None,
    )

    start_row = anchor_row + 1
    end_row = start_row + N_QUARTERS - 1
    candidate_cols = [
        max_col,
        min_col,
        forecast_col,
        actual_col,
        avg_pen_col,
        quarterly_col,
        reported_col,
        growth_col,
        captured_col,
        last_quarter_col,
        num_quarters_col,
    ]
    value_cols = [col for col in candidate_cols if col is not None and col > 0]
    if not value_cols:
        return rows

    block_start_col = min(value_cols)
    block_end_col = max(value_cols)
    data_block = read_block(sheet, start_row, end_row, block_start_col, block_end_col)

    # Use temporary R1C1 formula2 cells to calculate avg penetration from raw quarter support.
    avg_pen_values: list[Any] = [None] * N_QUARTERS
    if quarterly_col and reported_col and quarterly_col > 0 and reported_col > 0:
        scratch_col = max(context["end_col"], block_end_col) + 2
        for n in range(1, N_QUARTERS + 1):
            calc_row = start_row + n - 1
            hist_end_row = start_row + n - 1
            formula = (
                f'=IFERROR(AVERAGE(R{start_row}C{quarterly_col}:R{hist_end_row}C{quarterly_col}/'
                f'(R{start_row}C{reported_col}:R{hist_end_row}C{reported_col})),"")'
            )
            set_formula2(sheet.range((calc_row, scratch_col)), formula)

        wb.app.calculate()
        avg_pen_values = ensure_1d(sheet.range((start_row, scratch_col), (end_row, scratch_col)).value)
        if len(avg_pen_values) < N_QUARTERS:
            avg_pen_values.extend([None] * (N_QUARTERS - len(avg_pen_values)))

    for i in range(N_QUARTERS):
        num_quarters_used = block_get(data_block, i, num_quarters_col, block_start_col, block_end_col)
        if is_blank(num_quarters_used):
            num_quarters_used = i + 1

        last_quarter_used = block_get(data_block, i, last_quarter_col, block_start_col, block_end_col)
        forecast_value = block_get(data_block, i, forecast_col, block_start_col, block_end_col)
        actual_value = block_get(data_block, i, actual_col, block_start_col, block_end_col)
        forecast_max = block_get(data_block, i, max_col, block_start_col, block_end_col)
        forecast_min = block_get(data_block, i, min_col, block_start_col, block_end_col)
        quarterly_sales = block_get(data_block, i, quarterly_col, block_start_col, block_end_col)
        reported_sales = block_get(data_block, i, reported_col, block_start_col, block_end_col)
        growth_rate = block_get(data_block, i, growth_col, block_start_col, block_end_col)
        captured_pct = block_get(data_block, i, captured_col, block_start_col, block_end_col)

        avg_penetration_pct = avg_pen_values[i] if i < len(avg_pen_values) else None
        if is_blank(avg_penetration_pct):
            avg_penetration_pct = block_get(data_block, i, avg_pen_col, block_start_col, block_end_col)

        if all(
            is_blank(v)
            for v in (
                forecast_value,
                forecast_max,
                forecast_min,
                avg_penetration_pct,
                quarterly_sales,
                reported_sales,
            )
        ):
            continue

        rows.append(
            {
                "model": meta["model"],
                "ticker": meta["ticker"],
                "model_period": meta["model_period"],
                "model_date": meta["model_date"],
                "method": "empirical",
                "parameter_name": "avg_penetration_pct",
                "parameter_value": avg_penetration_pct,
                "num_quarters_used": num_quarters_used,
                "last_quarter_used": last_quarter_used,
                "forecast_value": forecast_value,
                "actual_value": actual_value,
                "forecast_max": forecast_max,
                "forecast_min": forecast_min,
                "range_width": numeric_diff(forecast_max, forecast_min),
                "avg_penetration_pct": avg_penetration_pct,
                "quarterly_sales": quarterly_sales,
                "reported_sales": reported_sales,
                "growth_rate_pct": growth_rate,
                "sales_captured_in_db_pct": captured_pct,
                "source_file": source_file,
            }
        )

    return rows


def extract_regression_candidates(
    wb: xw.Book,
    meta: dict[str, Any],
    source_file: str,
) -> list[dict[str, Any]]:
    rows: list[dict[str, Any]] = []
    try:
        sheet = wb.sheets["Regression Model"]
    except Exception:
        log(f"skipped regression extraction for {source_file} (missing sheet: Regression Model)")
        return rows

    context = build_sheet_context(sheet)
    anchor = find_anchor_cell(context, "max")
    if not anchor:
        log(f"skipped regression extraction for {source_file} (anchor 'max' not found)")
        return rows

    anchor_row, anchor_col = anchor
    header_values = get_row_values(context, anchor_row)
    header_start_col = context["start_col"]

    max_col = anchor_col
    min_col = find_column_by_patterns(header_values, header_start_col, [" min ", "minimum", "forecast_min"], default=anchor_col + 1)
    forecast_col = find_column_by_patterns(
        header_values,
        header_start_col,
        ["tot fcst w/o sa", "total fcst w/o sa", "tot fcst without sa", "forecast total without sa", "tot fcst"],
        default=anchor_col - 1,
    )
    actual_col = find_column_by_patterns(
        header_values,
        header_start_col,
        ["actual value", "actual", "reported sales"],
        default=None,
    )
    num_quarters_col = find_column_by_patterns(
        header_values,
        header_start_col,
        ["num quarters used", "num_quarters_used", "n quarters", "quarters used"],
        default=None,
    )

    # Required anchor offsets from existing regression logic.
    y_col = anchor_col - 7
    x_col = anchor_col - 11
    if y_col <= 0 or x_col <= 0:
        log(f"skipped regression extraction for {source_file} (invalid x/y anchor offsets)")
        return rows

    start_row = anchor_row + 1
    end_row = start_row + N_QUARTERS - 1

    candidate_cols = [num_quarters_col, forecast_col, actual_col, max_col, min_col]
    value_cols = [col for col in candidate_cols if col is not None and col > 0]
    if not value_cols:
        return rows

    block_start_col = min(value_cols)
    block_end_col = max(value_cols)
    data_block = read_block(sheet, start_row, end_row, block_start_col, block_end_col)

    scratch_intercept_col = max(context["end_col"], block_end_col) + 2
    scratch_slope_col = scratch_intercept_col + 1

    for n in range(1, N_QUARTERS + 1):
        calc_row = start_row + n - 1
        hist_end_row = start_row + n - 1
        intercept_formula = (
            f'=IFERROR(INTERCEPT(R{start_row}C{y_col}:R{hist_end_row}C{y_col},'
            f'R{start_row}C{x_col}:R{hist_end_row}C{x_col}),"")'
        )
        slope_formula = (
            f'=IFERROR(SLOPE(R{start_row}C{y_col}:R{hist_end_row}C{y_col},'
            f'R{start_row}C{x_col}:R{hist_end_row}C{x_col}),"")'
        )
        set_formula2(sheet.range((calc_row, scratch_intercept_col)), intercept_formula)
        set_formula2(sheet.range((calc_row, scratch_slope_col)), slope_formula)

    wb.app.calculate()
    intercept_values = ensure_1d(
        sheet.range((start_row, scratch_intercept_col), (end_row, scratch_intercept_col)).value
    )
    slope_values = ensure_1d(
        sheet.range((start_row, scratch_slope_col), (end_row, scratch_slope_col)).value
    )
    if len(intercept_values) < N_QUARTERS:
        intercept_values.extend([None] * (N_QUARTERS - len(intercept_values)))
    if len(slope_values) < N_QUARTERS:
        slope_values.extend([None] * (N_QUARTERS - len(slope_values)))

    previous_signature: tuple[Any, ...] | None = None
    for i in range(N_QUARTERS):
        num_quarters_used = block_get(data_block, i, num_quarters_col, block_start_col, block_end_col)
        if is_blank(num_quarters_used):
            num_quarters_used = i + 1

        forecast_value = block_get(data_block, i, forecast_col, block_start_col, block_end_col)
        actual_value = (
            block_get(data_block, i, actual_col, block_start_col, block_end_col)
            if actual_col is not None
            else ""
        )
        forecast_max = block_get(data_block, i, max_col, block_start_col, block_end_col)
        forecast_min = block_get(data_block, i, min_col, block_start_col, block_end_col)
        intercept = intercept_values[i] if i < len(intercept_values) else None
        slope = slope_values[i] if i < len(slope_values) else None

        if all(is_blank(v) for v in (forecast_value, forecast_max, forecast_min, intercept, slope)):
            continue

        current_signature = (
            signature_value(num_quarters_used),
            signature_value(intercept),
            signature_value(slope),
            signature_value(forecast_value),
            signature_value(forecast_max),
            signature_value(forecast_min),
        )

        # Avoid duplicate trailing row from recalculation artifacts.
        if current_signature == previous_signature:
            continue
        previous_signature = current_signature

        rows.append(
            {
                "model": meta["model"],
                "ticker": meta["ticker"],
                "model_period": meta["model_period"],
                "model_date": meta["model_date"],
                "method": "regression",
                "parameter_name": "num_quarters_used",
                "parameter_value": num_quarters_used,
                "num_quarters_used": num_quarters_used,
                "forecast_value": forecast_value,
                "actual_value": actual_value,
                "forecast_max": forecast_max,
                "forecast_min": forecast_min,
                "range_width": numeric_diff(forecast_max, forecast_min),
                "intercept": intercept,
                "slope": slope,
                "source_file": source_file,
            }
        )

    return rows


def write_sheet(
    wb: Workbook,
    sheet_name: str,
    columns: list[str],
    rows: list[dict[str, Any]],
) -> None:
    ws = wb.create_sheet(sheet_name)
    ws.append(columns)

    for row in rows:
        ws.append([row.get(column) for column in columns])

    for cell in ws[1]:
        cell.font = Font(bold=True)

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions

    for col_idx, column_name in enumerate(columns, start=1):
        max_len = len(column_name)
        for row_idx in range(2, ws.max_row + 1):
            value = ws.cell(row=row_idx, column=col_idx).value
            if value is None:
                continue
            max_len = max(max_len, len(str(value)))
        ws.column_dimensions[get_column_letter(col_idx)].width = min(max(max_len + 2, 12), 48)


def write_output_workbook(
    output_file: Path,
    empirical_rows: list[dict[str, Any]],
    regression_rows: list[dict[str, Any]],
) -> None:
    wb = Workbook()
    default_sheet = wb.active
    wb.remove(default_sheet)

    write_sheet(wb, "empirical_candidates", EMPIRICAL_COLUMNS, empirical_rows)
    write_sheet(wb, "regression_candidates", REGRESSION_COLUMNS, regression_rows)

    wb.save(output_file)


def should_skip_file(path: Path) -> tuple[bool, str]:
    name = path.name

    if not path.is_file():
        return True, "not a file"
    if name.startswith("~"):
        return True, "temporary excel file"
    if path.suffix.lower() != ".xlsx":
        return True, "not an .xlsx file"
    if re.search(r"_PARAM(\.\d+)?\.xlsx$", name, re.IGNORECASE):
        return True, "output workbook pattern"
    return False, ""


def main() -> None:
    input_path = Path(input_dir).expanduser()
    output_path = Path(output_dir).expanduser()

    if not input_path.exists() or not input_path.is_dir():
        raise FileNotFoundError(f"input_dir does not exist or is not a directory: {input_path}")
    output_path.mkdir(parents=True, exist_ok=True)

    output_file = resolve_output_path(input_path, output_path)

    empirical_rows: list[dict[str, Any]] = []
    regression_rows: list[dict[str, Any]] = []
    files_processed = 0

    app: xw.App | None = None
    try:
        app = xw.App(visible=False, add_book=False)
        app.display_alerts = False
        app.screen_updating = False
        app.calculation = "manual"

        for file_path in sorted(input_path.iterdir()):
            skip, reason = should_skip_file(file_path)
            if skip:
                log(f"skipped file: {file_path.name} ({reason})")
                continue

            meta = parse_file_metadata(file_path.name)
            if not meta:
                log(f"skipped file: {file_path.name} (filename does not match expected model pattern)")
                continue

            wb: xw.Book | None = None
            try:
                wb = app.books.open(str(file_path), update_links=False)
                empirical_rows.extend(extract_empirical_candidates(wb, meta, file_path.name))
                regression_rows.extend(extract_regression_candidates(wb, meta, file_path.name))
                files_processed += 1
                log(f"processed file: {file_path.name}")
            except Exception as exc:
                log(f"skipped file: {file_path.name} (processing error: {exc})")
            finally:
                if wb is not None:
                    close_workbook_safely(wb)
    finally:
        if app is not None:
            try:
                app.quit()
            except Exception:
                pass

    write_output_workbook(output_file, empirical_rows, regression_rows)

    log(f"output path: {output_file}")
    log(f"number of files processed: {files_processed}")
    log(f"number of empirical rows: {len(empirical_rows)}")
    log(f"number of regression rows: {len(regression_rows)}")


if __name__ == "__main__":
    main()
