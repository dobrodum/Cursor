from __future__ import annotations

import datetime as dt
import re
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Dict, List, Optional, Sequence, Tuple

import xlwings as xw
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

# =====================
# Configure these paths
# =====================
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


@dataclass(frozen=True)
class ModelMeta:
    model: str
    ticker: str
    model_period: str
    model_date: str


def normalize_text(value: Any) -> str:
    if not isinstance(value, str):
        return ""
    normalized = re.sub(r"[^a-z0-9]+", " ", value.strip().lower())
    return re.sub(r"\s+", " ", normalized).strip()


def is_blank(value: Any) -> bool:
    return value is None or (isinstance(value, str) and value.strip() == "")


def to_float(value: Any) -> Optional[float]:
    if is_blank(value):
        return None
    try:
        return float(value)
    except (TypeError, ValueError):
        return None


def to_2d(values: Any, n_rows: int, n_cols: int) -> List[List[Any]]:
    if n_rows <= 0 or n_cols <= 0:
        return []

    if n_rows == 1 and n_cols == 1:
        return [[values]]

    if n_rows == 1:
        row = list(values) if isinstance(values, (list, tuple)) else [values]
        if len(row) < n_cols:
            row = row + [None] * (n_cols - len(row))
        return [row]

    if n_cols == 1:
        col = list(values) if isinstance(values, (list, tuple)) else [values]
        if len(col) < n_rows:
            col = col + [None] * (n_rows - len(col))
        return [[v] for v in col]

    if not isinstance(values, (list, tuple)):
        return [[values] + [None] * (n_cols - 1)] + [[None] * n_cols for _ in range(n_rows - 1)]

    rows: List[List[Any]] = []
    for row in values:
        if isinstance(row, (list, tuple)):
            row_list = list(row)
        else:
            row_list = [row]
        if len(row_list) < n_cols:
            row_list = row_list + [None] * (n_cols - len(row_list))
        rows.append(row_list[:n_cols])

    if len(rows) < n_rows:
        rows.extend([[None] * n_cols for _ in range(n_rows - len(rows))])
    return rows[:n_rows]


def build_output_path(input_folder: Path, out_folder: Path) -> Path:
    out_folder.mkdir(parents=True, exist_ok=True)
    base_name = f"{input_folder.name}_PARAM"
    candidate = out_folder / f"{base_name}.xlsx"
    if not candidate.exists():
        return candidate

    suffix = 1
    while True:
        candidate = out_folder / f"{base_name}.{suffix}.xlsx"
        if not candidate.exists():
            return candidate
        suffix += 1


def parse_model_meta(file_path: Path) -> ModelMeta:
    stem = file_path.stem
    parts = [p.strip() for p in stem.split(" - ")]

    ticker = ""
    if len(parts) >= 2:
        ticker = parts[1].strip().upper()
    if not ticker:
        ticker = parts[0].strip().upper() if parts else "UNKNOWN"

    period_token = ""
    if len(parts) >= 3:
        period_token = parts[2].split("_")[0].strip()

    match = re.search(r"(Early|Mid|Late)([A-Za-z]{3})(\d{4})", period_token, re.IGNORECASE)
    if match:
        phase_raw, month_raw, year_raw = match.groups()
        phase = phase_raw[:1].upper() + phase_raw[1:].lower()
        month_abbr = month_raw[:1].upper() + month_raw[1:].lower()
        year = int(year_raw)
        month_num = dt.datetime.strptime(month_abbr, "%b").month
        day_lookup = {"Early": 5, "Mid": 15, "Late": 25}
        day = day_lookup[phase]
        model_period = f"{phase}{month_abbr}_{year}"
        model_date = dt.date(year, month_num, day).isoformat()
    else:
        fallback_period = period_token if period_token else "UnknownPeriod"
        model_period = fallback_period
        model_date = ""

    model = f"{ticker}_{model_period}" if model_period else ticker
    return ModelMeta(model=model, ticker=ticker, model_period=model_period, model_date=model_date)


def safe_close_source_workbook(wb: xw.Book) -> None:
    try:
        wb.close(save=False)
        return
    except TypeError:
        try:
            wb.close(False)
            return
        except Exception:
            pass
    except Exception:
        pass

    try:
        wb.api.Close(SaveChanges=False)
        return
    except Exception:
        pass

    try:
        wb.close()
    except Exception:
        pass


def get_sheet_case_insensitive(wb: xw.Book, sheet_name: str) -> Optional[xw.Sheet]:
    wanted = sheet_name.strip().lower()
    for sh in wb.sheets:
        if sh.name.strip().lower() == wanted:
            return sh
    return None


def find_anchor_max_cell(sheet: xw.Sheet) -> Optional[Tuple[int, int]]:
    used = sheet.used_range
    n_rows = int(used.rows.count)
    n_cols = int(used.columns.count)
    if n_rows <= 0 or n_cols <= 0:
        return None

    values = to_2d(used.value, n_rows, n_cols)
    start_row = int(used.row)
    start_col = int(used.column)

    for r_offset, row_values in enumerate(values):
        for c_offset, cell_value in enumerate(row_values):
            if normalize_text(cell_value) == "max":
                return start_row + r_offset, start_col + c_offset
    return None


def collect_header_candidates(sheet: xw.Sheet, anchor_row: int, anchor_col: int) -> List[Tuple[str, int, int]]:
    row_start = max(1, anchor_row - 2)
    row_end = anchor_row + 2
    col_start = max(1, anchor_col - 25)
    col_end = anchor_col + 10

    n_rows = row_end - row_start + 1
    n_cols = col_end - col_start + 1
    values = to_2d(sheet.range((row_start, col_start), (row_end, col_end)).value, n_rows, n_cols)

    candidates: List[Tuple[str, int, int]] = []
    for r_offset, row_values in enumerate(values):
        for c_offset, cell_value in enumerate(row_values):
            normalized = normalize_text(cell_value)
            if normalized:
                candidates.append((normalized, row_start + r_offset, col_start + c_offset))
    return candidates


def pick_column(
    candidates: Sequence[Tuple[str, int, int]],
    keywords: Sequence[str],
    anchor_row: int,
    anchor_col: int,
    fallback_col: Optional[int],
) -> Optional[int]:
    best_col: Optional[int] = None
    best_score: Optional[int] = None

    for normalized, row, col in candidates:
        if any(keyword in normalized for keyword in keywords):
            score = abs(col - anchor_col) + (3 * abs(row - anchor_row))
            if best_score is None or score < best_score:
                best_score = score
                best_col = col

    if best_col is not None:
        return best_col
    if fallback_col is not None and fallback_col > 0:
        return fallback_col
    return None


def read_block(
    sheet: xw.Sheet,
    start_row: int,
    end_row: int,
    start_col: int,
    end_col: int,
) -> List[List[Any]]:
    if end_row < start_row or end_col < start_col:
        return []

    n_rows = end_row - start_row + 1
    n_cols = end_col - start_col + 1
    values = sheet.range((start_row, start_col), (end_row, end_col)).value
    return to_2d(values, n_rows, n_cols)


def block_value(block: List[List[Any]], row_offset: int, col: Optional[int], start_col: int) -> Any:
    if col is None:
        return None
    if row_offset < 0 or row_offset >= len(block):
        return None
    col_offset = col - start_col
    if col_offset < 0 or not block or col_offset >= len(block[row_offset]):
        return None
    return block[row_offset][col_offset]


def set_formula_r1c1_formula2(cell: xw.Range, formula: str) -> None:
    try:
        cell.api.Formula2R1C1 = formula
    except Exception:
        try:
            cell.formula2 = formula
        except Exception:
            cell.formula = formula


def compute_range_width(max_value: Any, min_value: Any) -> Optional[float]:
    max_num = to_float(max_value)
    min_num = to_float(min_value)
    if max_num is None or min_num is None:
        return None
    return max_num - min_num


def dedupe_value(value: Any) -> Any:
    num = to_float(value)
    if num is not None:
        return round(num, 10)
    return value


def extract_empirical_rows(wb: xw.Book, meta: ModelMeta, source_file: str) -> List[Dict[str, Any]]:
    sheet = get_sheet_case_insensitive(wb, "Empirical Model")
    if sheet is None:
        return []

    anchor = find_anchor_max_cell(sheet)
    if anchor is None:
        return []

    anchor_row, anchor_col = anchor
    headers = collect_header_candidates(sheet, anchor_row, anchor_col)

    num_quarters_col = pick_column(
        headers, ["num quarters", "quarters used", "n qtr", "n quarter"], anchor_row, anchor_col, anchor_col - 12
    )
    last_quarter_col = pick_column(
        headers, ["last quarter", "last qtr"], anchor_row, anchor_col, anchor_col - 10
    )
    forecast_value_col = pick_column(
        headers,
        ["estimated total sold", "total forecast", "tot fcst", "forecast"],
        anchor_row,
        anchor_col,
        anchor_col - 2,
    )
    actual_value_col = pick_column(
        headers,
        ["reported sales", "actual sales", "actual"],
        anchor_row,
        anchor_col,
        anchor_col - 7,
    )
    forecast_max_col = anchor_col
    forecast_min_col = pick_column(headers, ["min"], anchor_row, anchor_col, anchor_col + 1)
    avg_penetration_col = pick_column(
        headers,
        ["avg penetration", "average penetration", "penetration pct", "pen pct"],
        anchor_row,
        anchor_col,
        anchor_col - 4,
    )
    quarterly_sales_col = pick_column(
        headers,
        ["quarterly sales", "qtr sales", "db sales", "sales in db"],
        anchor_row,
        anchor_col,
        anchor_col - 11,
    )
    reported_sales_col = pick_column(
        headers,
        ["reported sales", "actual sales", "reported"],
        anchor_row,
        anchor_col,
        anchor_col - 7,
    )
    growth_rate_col = pick_column(
        headers, ["growth rate", "growth"], anchor_row, anchor_col, anchor_col - 6
    )
    sales_captured_col = pick_column(
        headers,
        ["sales captured", "captured in db", "db captured"],
        anchor_row,
        anchor_col,
        anchor_col - 5,
    )

    data_start_row = anchor_row + 1
    n_quarters = 10
    data_end_row = data_start_row + n_quarters - 1

    wanted_cols = [
        num_quarters_col,
        last_quarter_col,
        forecast_value_col,
        actual_value_col,
        forecast_max_col,
        forecast_min_col,
        avg_penetration_col,
        quarterly_sales_col,
        reported_sales_col,
        growth_rate_col,
        sales_captured_col,
    ]
    wanted_cols = [col for col in wanted_cols if col is not None and col > 0]
    if not wanted_cols:
        return []

    start_col = min(wanted_cols)
    end_col = max(wanted_cols)
    block = read_block(sheet, data_start_row, data_end_row, start_col, end_col)

    helper_col = min(16384, max(anchor_col + 20, int(sheet.used_range.last_cell.column) + 2))
    avg_pen_calc_cell = sheet.cells(1, helper_col)

    rows: List[Dict[str, Any]] = []
    empty_streak = 0

    for i in range(n_quarters):
        current_row = data_start_row + i

        num_quarters_used = block_value(block, i, num_quarters_col, start_col)
        last_quarter_used = block_value(block, i, last_quarter_col, start_col)
        forecast_value = block_value(block, i, forecast_value_col, start_col)
        actual_value = block_value(block, i, actual_value_col, start_col)
        forecast_max = block_value(block, i, forecast_max_col, start_col)
        forecast_min = block_value(block, i, forecast_min_col, start_col)
        quarterly_sales = block_value(block, i, quarterly_sales_col, start_col)
        reported_sales = block_value(block, i, reported_sales_col, start_col)
        growth_rate_pct = block_value(block, i, growth_rate_col, start_col)
        sales_captured_in_db_pct = block_value(block, i, sales_captured_col, start_col)

        avg_penetration_pct = block_value(block, i, avg_penetration_col, start_col)
        formula_written = False

        if avg_penetration_col is not None:
            formula = (
                f'=IFERROR(AVERAGE(R{data_start_row}C{avg_penetration_col}:R{current_row}C{avg_penetration_col}),"")'
            )
            set_formula_r1c1_formula2(avg_pen_calc_cell, formula)
            formula_written = True
        elif quarterly_sales_col is not None and reported_sales_col is not None:
            formula = (
                f'=LET(q,R{data_start_row}C{quarterly_sales_col}:R{current_row}C{quarterly_sales_col},'
                f'r,R{data_start_row}C{reported_sales_col}:R{current_row}C{reported_sales_col},'
                f'IFERROR(AVERAGE(IFERROR(q/r,"")),""))'
            )
            set_formula_r1c1_formula2(avg_pen_calc_cell, formula)
            formula_written = True

        if formula_written:
            wb.app.calculate()
            avg_penetration_pct = avg_pen_calc_cell.value

        has_data = any(
            not is_blank(v)
            for v in [
                forecast_value,
                actual_value,
                forecast_max,
                forecast_min,
                avg_penetration_pct,
                quarterly_sales,
                reported_sales,
            ]
        )
        if not has_data:
            empty_streak += 1
            if empty_streak >= 2:
                break
            continue
        empty_streak = 0

        if is_blank(num_quarters_used):
            num_quarters_used = i + 1

        row = {
            "model": meta.model,
            "ticker": meta.ticker,
            "model_period": meta.model_period,
            "model_date": meta.model_date,
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
        rows.append(row)

    avg_pen_calc_cell.value = None
    return rows


def extract_regression_rows(wb: xw.Book, meta: ModelMeta, source_file: str) -> List[Dict[str, Any]]:
    sheet = get_sheet_case_insensitive(wb, "Regression Model")
    if sheet is None:
        return []

    anchor = find_anchor_max_cell(sheet)
    if anchor is None:
        return []

    anchor_row, anchor_col = anchor
    headers = collect_header_candidates(sheet, anchor_row, anchor_col)

    # Required anchor-based offsets.
    y_col = anchor_col - 7
    x_col = anchor_col - 11

    num_quarters_col = pick_column(
        headers, ["num quarters", "quarters used", "n qtr", "n quarter"], anchor_row, anchor_col, anchor_col - 12
    )
    forecast_total_col = pick_column(
        headers, ["tot fcst w o sa", "tot fcst", "total forecast", "forecast"], anchor_row, anchor_col, anchor_col - 2
    )
    actual_value_col = pick_column(
        headers, ["actual", "actual sales", "reported sales"], anchor_row, anchor_col, None
    )
    forecast_max_col = anchor_col
    forecast_min_col = pick_column(headers, ["min"], anchor_row, anchor_col, anchor_col + 1)

    data_start_row = anchor_row + 1
    n_quarters = 10
    data_end_row = data_start_row + n_quarters - 1

    wanted_cols = [
        num_quarters_col,
        forecast_total_col,
        actual_value_col,
        forecast_max_col,
        forecast_min_col,
        x_col,
        y_col,
    ]
    wanted_cols = [col for col in wanted_cols if col is not None and col > 0]
    if not wanted_cols:
        return []

    start_col = min(wanted_cols)
    end_col = max(wanted_cols)
    block = read_block(sheet, data_start_row, data_end_row, start_col, end_col)

    helper_col = min(16384, max(anchor_col + 20, int(sheet.used_range.last_cell.column) + 2))
    intercept_cell = sheet.cells(1, helper_col)
    slope_cell = sheet.cells(2, helper_col)

    rows: List[Dict[str, Any]] = []
    prev_key: Optional[Tuple[Any, ...]] = None
    empty_streak = 0

    for i in range(n_quarters):
        current_row = data_start_row + i

        intercept_formula = (
            f'=IFERROR(INTERCEPT(R{data_start_row}C{y_col}:R{current_row}C{y_col},'
            f'R{data_start_row}C{x_col}:R{current_row}C{x_col}),"")'
        )
        slope_formula = (
            f'=IFERROR(SLOPE(R{data_start_row}C{y_col}:R{current_row}C{y_col},'
            f'R{data_start_row}C{x_col}:R{current_row}C{x_col}),"")'
        )

        set_formula_r1c1_formula2(intercept_cell, intercept_formula)
        set_formula_r1c1_formula2(slope_cell, slope_formula)
        wb.app.calculate()

        intercept = intercept_cell.value
        slope = slope_cell.value

        num_quarters_used = block_value(block, i, num_quarters_col, start_col)
        forecast_value = block_value(block, i, forecast_total_col, start_col)
        actual_value = block_value(block, i, actual_value_col, start_col)
        forecast_max = block_value(block, i, forecast_max_col, start_col)
        forecast_min = block_value(block, i, forecast_min_col, start_col)

        if is_blank(forecast_value):
            x_current = block_value(block, i, x_col, start_col)
            x_num = to_float(x_current)
            intercept_num = to_float(intercept)
            slope_num = to_float(slope)
            if x_num is not None and intercept_num is not None and slope_num is not None:
                forecast_value = intercept_num + (slope_num * x_num)

        has_data = any(
            not is_blank(v)
            for v in [num_quarters_used, forecast_value, forecast_max, forecast_min, intercept, slope]
        )
        if not has_data:
            empty_streak += 1
            if empty_streak >= 2:
                break
            continue
        empty_streak = 0

        if is_blank(num_quarters_used):
            num_quarters_used = i + 1

        row_key = (
            dedupe_value(num_quarters_used),
            dedupe_value(intercept),
            dedupe_value(slope),
            dedupe_value(forecast_value),
            dedupe_value(forecast_max),
            dedupe_value(forecast_min),
        )
        if prev_key is not None and row_key == prev_key:
            continue
        prev_key = row_key

        row = {
            "model": meta.model,
            "ticker": meta.ticker,
            "model_period": meta.model_period,
            "model_date": meta.model_date,
            "method": "regression",
            "parameter_name": "num_quarters_used",
            "parameter_value": num_quarters_used,
            "num_quarters_used": num_quarters_used,
            "forecast_value": forecast_value,
            "actual_value": actual_value if not is_blank(actual_value) else None,
            "forecast_max": forecast_max,
            "forecast_min": forecast_min,
            "range_width": compute_range_width(forecast_max, forecast_min),
            "intercept": intercept,
            "slope": slope,
            "source_file": source_file,
        }
        rows.append(row)

    intercept_cell.value = None
    slope_cell.value = None
    return rows


def write_sheet(ws: Any, columns: List[str], rows: List[Dict[str, Any]]) -> None:
    ws.append(columns)
    for row in rows:
        ws.append([row.get(col) for col in columns])

    for cell in ws[1]:
        cell.font = Font(bold=True)

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(ws.max_column)}{max(1, ws.max_row)}"

    for idx, col_name in enumerate(columns, start=1):
        max_len = len(col_name)
        for row_idx in range(2, ws.max_row + 1):
            value = ws.cell(row=row_idx, column=idx).value
            if value is None:
                continue
            max_len = max(max_len, len(str(value)))
        ws.column_dimensions[get_column_letter(idx)].width = min(60, max(12, max_len + 2))


def write_output_workbook(
    output_path: Path,
    empirical_rows: List[Dict[str, Any]],
    regression_rows: List[Dict[str, Any]],
) -> None:
    workbook = Workbook()
    empirical_ws = workbook.active
    empirical_ws.title = "empirical_candidates"
    regression_ws = workbook.create_sheet("regression_candidates")

    write_sheet(empirical_ws, EMPIRICAL_COLUMNS, empirical_rows)
    write_sheet(regression_ws, REGRESSION_COLUMNS, regression_rows)
    workbook.save(output_path)


def process_all_workbooks() -> None:
    if not input_dir.exists():
        raise FileNotFoundError(f"Input directory does not exist: {input_dir}")

    output_path = build_output_path(input_dir, output_dir)
    empirical_rows: List[Dict[str, Any]] = []
    regression_rows: List[Dict[str, Any]] = []
    files_processed = 0

    app = xw.App(visible=False, add_book=False)
    app.display_alerts = False
    app.screen_updating = False

    original_calculation = None
    try:
        try:
            original_calculation = app.calculation
            app.calculation = "manual"
        except Exception:
            original_calculation = None

        for file_path in sorted(input_dir.iterdir(), key=lambda p: p.name.lower()):
            if not file_path.is_file():
                print(f"Skipped {file_path.name}: not a file")
                continue
            if file_path.name.startswith("~"):
                print(f"Skipped {file_path.name}: temp file")
                continue
            if file_path.suffix.lower() != ".xlsx":
                print(f"Skipped {file_path.name}: not an .xlsx file")
                continue

            print(f"Processing {file_path.name}")
            wb: Optional[xw.Book] = None
            try:
                meta = parse_model_meta(file_path)
                wb = app.books.open(str(file_path), update_links=False)
                empirical_rows.extend(extract_empirical_rows(wb, meta, file_path.name))
                regression_rows.extend(extract_regression_rows(wb, meta, file_path.name))
                files_processed += 1
            except Exception as exc:
                print(f"Skipped {file_path.name}: {exc}")
            finally:
                if wb is not None:
                    safe_close_source_workbook(wb)

    finally:
        if original_calculation is not None:
            try:
                app.calculation = original_calculation
            except Exception:
                pass
        app.quit()

    write_output_workbook(output_path, empirical_rows, regression_rows)

    print(f"Output path: {output_path}")
    print(f"Number of files processed: {files_processed}")
    print(f"Number of empirical rows: {len(empirical_rows)}")
    print(f"Number of regression rows: {len(regression_rows)}")


if __name__ == "__main__":
    process_all_workbooks()
