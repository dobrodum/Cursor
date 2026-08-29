from __future__ import annotations

import calendar
import re
from datetime import date
from pathlib import Path
from typing import Any, Dict, List, Optional, Sequence, Tuple

import xlwings as xw
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

# -----------------------------
# Configure these two paths only
# -----------------------------
input_dir = Path("/path/to/input")
output_dir = Path("/path/to/output")

EMPIRICAL_SHEET_NAME = "Empirical Model"
REGRESSION_SHEET_NAME = "Regression Model"
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

EMPIRICAL_HEADER_ALIASES: Dict[str, Sequence[Sequence[str]]] = {
    "num_quarters_used": (("num", "quarter"), ("quarters", "used"), ("n", "quarter")),
    "last_quarter_used": (("last", "quarter"), ("quarter", "used")),
    "quarterly_sales": (("quarterly", "sales"), ("qtr", "sales")),
    "reported_sales": (("reported", "sales"), ("actual", "sales")),
    "growth_rate_pct": (("growth", "rate"),),
    "sales_captured_in_db_pct": (("captured", "db"), ("sales", "captured")),
    "penetration_pct": (("penetration",),),
    "forecast_value": (
        ("estimated", "total", "sold"),
        ("total", "estimated", "sold"),
        ("forecast", "value"),
    ),
    "actual_value": (("reported", "sales"), ("actual", "sales")),
    "forecast_max": (("forecast", "max"), ("max",)),
    "forecast_min": (("forecast", "min"), ("min",)),
}

REGRESSION_HEADER_ALIASES: Dict[str, Sequence[Sequence[str]]] = {
    "num_quarters_used": (("num", "quarter"), ("quarters", "used"), ("n", "quarter")),
    "forecast_value": (
        ("tot", "fcst", "w/o", "sa"),
        ("tot", "fcst", "wo", "sa"),
        ("forecast", "without", "sa"),
        ("tot", "forecast", "sa"),
    ),
    "actual_value": (("actual", "sales"), ("reported", "sales")),
    "forecast_max": (("forecast", "max"), ("max",)),
    "forecast_min": (("forecast", "min"), ("min",)),
}

EMPIRICAL_FALLBACK_OFFSETS = {
    "num_quarters_used": -12,
    "last_quarter_used": -11,
    "quarterly_sales": -10,
    "reported_sales": -9,
    "growth_rate_pct": -8,
    "sales_captured_in_db_pct": -7,
    "penetration_pct": -6,
    "actual_value": -5,
    "forecast_value": -4,
    "forecast_min": -1,
    "forecast_max": 0,
}

REGRESSION_FALLBACK_OFFSETS = {
    "num_quarters_used": -12,
    "forecast_value": -4,
    "actual_value": -5,
    "forecast_min": -1,
    "forecast_max": 0,
}

PERIOD_DAY_MAP = {"early": 5, "mid": 15, "late": 25}


def normalize_text(value: Any) -> str:
    if value is None:
        return ""
    text = str(value).strip().lower()
    text = re.sub(r"[_\-\n\r\t]+", " ", text)
    text = re.sub(r"\s+", " ", text)
    return text


def to_float(value: Any) -> Optional[float]:
    if value in (None, ""):
        return None
    if isinstance(value, bool):
        return None
    try:
        return float(value)
    except (TypeError, ValueError):
        return None


def to_int(value: Any) -> Optional[int]:
    numeric = to_float(value)
    if numeric is None:
        return None
    try:
        return int(round(numeric))
    except (TypeError, ValueError):
        return None


def safe_value(value: Any) -> Any:
    if isinstance(value, date):
        return value.isoformat()
    return value


def compute_range_width(max_value: Any, min_value: Any) -> Optional[float]:
    max_num = to_float(max_value)
    min_num = to_float(min_value)
    if max_num is None or min_num is None:
        return None
    return max_num - min_num


def parse_filename_labels(file_name: str) -> Dict[str, str]:
    stem = Path(file_name).stem
    parts = [p.strip() for p in re.split(r"\s*-\s*", stem) if p.strip()]

    ticker = parts[1] if len(parts) >= 2 else stem
    period_block = parts[2] if len(parts) >= 3 else stem

    match = re.search(
        r"(Early|Mid|Late)\s*([A-Za-z]{3,9})\s*(\d{4})",
        period_block.replace("_", " "),
        re.IGNORECASE,
    )

    model_period = ""
    model_date = ""
    if match:
        period_label = match.group(1).title()
        month_token = match.group(2).title()[:3]
        year = int(match.group(3))
        month_num = list(calendar.month_abbr).index(month_token)
        day = PERIOD_DAY_MAP[period_label.lower()]
        model_period = f"{period_label}{month_token}_{year}"
        model_date = date(year, month_num, day).isoformat()

    model = f"{ticker}_{model_period}" if model_period else ticker
    return {
        "model": model,
        "ticker": ticker,
        "model_period": model_period,
        "model_date": model_date,
    }


def build_output_path(input_path: Path, output_path: Path) -> Path:
    folder_name = input_path.name
    base_name = f"{folder_name}_PARAM"

    candidate = output_path / f"{base_name}.xlsx"
    if not candidate.exists():
        return candidate

    suffix = 1
    while True:
        candidate = output_path / f"{base_name}.{suffix}.xlsx"
        if not candidate.exists():
            return candidate
        suffix += 1


def close_workbook_safe(wb: xw.Book) -> None:
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
    except Exception:
        pass


def get_sheet_by_name(wb: xw.Book, sheet_name: str) -> Optional[xw.Sheet]:
    for sheet in wb.sheets:
        if sheet.name.strip().lower() == sheet_name.lower():
            return sheet
    return None


def normalize_grid(values: Any) -> List[List[Any]]:
    if values is None:
        return []
    if isinstance(values, list):
        if not values:
            return []
        if isinstance(values[0], list):
            return values
        return [values]
    if isinstance(values, tuple):
        if not values:
            return []
        if isinstance(values[0], (list, tuple)):
            return [list(row) if isinstance(row, tuple) else row for row in values]
        return [list(values)]
    return [[values]]


def read_used_grid(sheet: xw.Sheet) -> Optional[Tuple[int, int, List[List[Any]]]]:
    used = sheet.used_range
    values = normalize_grid(used.value)
    if not values:
        return None
    return used.row, used.column, values


def find_anchor(grid: List[List[Any]], base_row: int, base_col: int, needle: str = "max") -> Optional[Tuple[int, int]]:
    target = needle.strip().lower()
    for r_idx, row in enumerate(grid):
        for c_idx, raw_value in enumerate(row):
            if normalize_text(raw_value) == target:
                return base_row + r_idx, base_col + c_idx
    return None


def collect_nearby_headers(
    grid: List[List[Any]],
    base_row: int,
    base_col: int,
    anchor_row: int,
    max_row_distance: int = 3,
) -> List[Tuple[int, int, str]]:
    collected: List[Tuple[int, int, str]] = []
    for r_idx, row in enumerate(grid):
        abs_row = base_row + r_idx
        if abs(abs_row - anchor_row) > max_row_distance:
            continue
        for c_idx, raw_value in enumerate(row):
            text = normalize_text(raw_value)
            if text:
                collected.append((abs_row, base_col + c_idx, text))
    return collected


def resolve_columns(
    headers: List[Tuple[int, int, str]],
    anchor_row: int,
    anchor_col: int,
    aliases: Dict[str, Sequence[Sequence[str]]],
    fallback_offsets: Dict[str, int],
) -> Dict[str, Optional[int]]:
    resolved: Dict[str, Optional[int]] = {}

    for field_name, alternatives in aliases.items():
        best_col: Optional[int] = None
        best_score: Optional[int] = None
        for cell_row, cell_col, text in headers:
            for token_group in alternatives:
                if all(token in text for token in token_group):
                    score = abs(cell_row - anchor_row) * 100 + abs(cell_col - anchor_col)
                    if best_score is None or score < best_score:
                        best_score = score
                        best_col = cell_col

        if best_col is None and field_name in fallback_offsets:
            best_col = anchor_col + fallback_offsets[field_name]

        resolved[field_name] = best_col

    return resolved


def get_cell_value(sheet: xw.Sheet, row: int, col: Optional[int]) -> Any:
    if col is None or row < 1 or col < 1:
        return None
    return sheet.range((row, col)).value


def determine_data_rows(sheet: xw.Sheet, anchor_row: int, reference_col: int, n_rows: int) -> Tuple[List[int], int]:
    down_rows = [anchor_row + i for i in range(1, n_rows + 1)]
    up_rows = [anchor_row - i for i in range(1, n_rows + 1)]

    down_score = sum(1 for row in down_rows if to_float(get_cell_value(sheet, row, reference_col)) is not None)
    up_score = sum(1 for row in up_rows if to_float(get_cell_value(sheet, row, reference_col)) is not None)

    if up_score > down_score:
        return up_rows, -1
    return down_rows, 1


def rolling_bounds(row: int, window_size: int, direction: int) -> Tuple[int, int]:
    if direction >= 0:
        return row - window_size + 1, row
    return row, row + window_size - 1


def col_values_by_row(sheet: xw.Sheet, rows: List[int], col: int) -> Dict[int, Any]:
    if not rows:
        return {}

    min_row = min(rows)
    max_row = max(rows)
    values = normalize_grid(sheet.range((min_row, col), (max_row, col)).value)

    result: Dict[int, Any] = {}
    for index, row_num in enumerate(range(min_row, max_row + 1)):
        value = values[index][0] if index < len(values) and values[index] else None
        result[row_num] = value
    return result


def clear_temp_column(sheet: xw.Sheet, rows: List[int], col: int) -> None:
    if not rows:
        return
    min_row = min(rows)
    max_row = max(rows)
    sheet.range((min_row, col), (max_row, col)).clear_contents()


def extract_empirical_rows(wb: xw.Book, metadata: Dict[str, str], source_file: str) -> List[Dict[str, Any]]:
    sheet = get_sheet_by_name(wb, EMPIRICAL_SHEET_NAME)
    if sheet is None:
        print(f"  skipped empirical for {source_file}: missing '{EMPIRICAL_SHEET_NAME}' sheet")
        return []

    grid_info = read_used_grid(sheet)
    if grid_info is None:
        print(f"  skipped empirical for {source_file}: empty '{EMPIRICAL_SHEET_NAME}' sheet")
        return []

    base_row, base_col, grid = grid_info
    anchor = find_anchor(grid, base_row, base_col, "max")
    if anchor is None:
        print(f"  skipped empirical for {source_file}: unable to find 'max' anchor")
        return []

    anchor_row, anchor_col = anchor
    headers = collect_nearby_headers(grid, base_row, base_col, anchor_row)
    columns = resolve_columns(
        headers=headers,
        anchor_row=anchor_row,
        anchor_col=anchor_col,
        aliases=EMPIRICAL_HEADER_ALIASES,
        fallback_offsets=EMPIRICAL_FALLBACK_OFFSETS,
    )

    reference_col = columns.get("forecast_max") or anchor_col
    rows, direction = determine_data_rows(sheet, anchor_row, reference_col, N_QUARTERS)

    used_last_col = sheet.used_range.last_cell.column
    temp_avg_col = max(used_last_col + 2, anchor_col + 20)

    penetration_col = columns.get("penetration_pct")
    avg_by_row: Dict[int, Any] = {row: None for row in rows}
    if penetration_col is not None:
        for index, row in enumerate(rows, start=1):
            start_row, end_row = rolling_bounds(row, index, direction)
            sheet.range((row, temp_avg_col)).formula2 = (
                f"=AVERAGE(R{start_row}C{penetration_col}:R{end_row}C{penetration_col})"
            )
        wb.app.calculate()
        avg_by_row = col_values_by_row(sheet, rows, temp_avg_col)
        clear_temp_column(sheet, rows, temp_avg_col)

    extracted_rows: List[Dict[str, Any]] = []
    for index, row in enumerate(rows, start=1):
        num_quarters_used = to_int(get_cell_value(sheet, row, columns.get("num_quarters_used"))) or index
        last_quarter_used = safe_value(get_cell_value(sheet, row, columns.get("last_quarter_used")))
        forecast_value = get_cell_value(sheet, row, columns.get("forecast_value"))
        actual_value = get_cell_value(sheet, row, columns.get("actual_value"))
        forecast_max = get_cell_value(sheet, row, columns.get("forecast_max"))
        forecast_min = get_cell_value(sheet, row, columns.get("forecast_min"))
        avg_penetration_pct = avg_by_row.get(row)
        quarterly_sales = get_cell_value(sheet, row, columns.get("quarterly_sales"))
        reported_sales = get_cell_value(sheet, row, columns.get("reported_sales"))
        growth_rate_pct = get_cell_value(sheet, row, columns.get("growth_rate_pct"))
        sales_captured_in_db_pct = get_cell_value(sheet, row, columns.get("sales_captured_in_db_pct"))

        extracted_rows.append(
            {
                "model": metadata["model"],
                "ticker": metadata["ticker"],
                "model_period": metadata["model_period"],
                "model_date": metadata["model_date"],
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

    return extracted_rows


def extract_regression_rows(wb: xw.Book, metadata: Dict[str, str], source_file: str) -> List[Dict[str, Any]]:
    sheet = get_sheet_by_name(wb, REGRESSION_SHEET_NAME)
    if sheet is None:
        print(f"  skipped regression for {source_file}: missing '{REGRESSION_SHEET_NAME}' sheet")
        return []

    grid_info = read_used_grid(sheet)
    if grid_info is None:
        print(f"  skipped regression for {source_file}: empty '{REGRESSION_SHEET_NAME}' sheet")
        return []

    base_row, base_col, grid = grid_info
    anchor = find_anchor(grid, base_row, base_col, "max")
    if anchor is None:
        print(f"  skipped regression for {source_file}: unable to find 'max' anchor")
        return []

    anchor_row, anchor_col = anchor
    headers = collect_nearby_headers(grid, base_row, base_col, anchor_row)
    columns = resolve_columns(
        headers=headers,
        anchor_row=anchor_row,
        anchor_col=anchor_col,
        aliases=REGRESSION_HEADER_ALIASES,
        fallback_offsets=REGRESSION_FALLBACK_OFFSETS,
    )

    # Required by the workflow specification.
    y_col = anchor_col - 7
    x_col = anchor_col - 11

    reference_col = columns.get("forecast_max") or anchor_col
    rows, direction = determine_data_rows(sheet, anchor_row, reference_col, N_QUARTERS)

    used_last_col = sheet.used_range.last_cell.column
    temp_intercept_col = max(used_last_col + 2, anchor_col + 20)
    temp_slope_col = temp_intercept_col + 1

    for index, row in enumerate(rows, start=1):
        start_row, end_row = rolling_bounds(row, index, direction)
        intercept_formula = (
            f"=INTERCEPT(R{start_row}C{y_col}:R{end_row}C{y_col},R{start_row}C{x_col}:R{end_row}C{x_col})"
        )
        slope_formula = (
            f"=SLOPE(R{start_row}C{y_col}:R{end_row}C{y_col},R{start_row}C{x_col}:R{end_row}C{x_col})"
        )
        sheet.range((row, temp_intercept_col)).formula2 = intercept_formula
        sheet.range((row, temp_slope_col)).formula2 = slope_formula

    wb.app.calculate()

    intercept_by_row = col_values_by_row(sheet, rows, temp_intercept_col)
    slope_by_row = col_values_by_row(sheet, rows, temp_slope_col)
    clear_temp_column(sheet, rows, temp_intercept_col)
    clear_temp_column(sheet, rows, temp_slope_col)

    extracted_rows: List[Dict[str, Any]] = []
    previous_signature: Optional[Tuple[Any, ...]] = None

    for index, row in enumerate(rows, start=1):
        num_quarters_used = to_int(get_cell_value(sheet, row, columns.get("num_quarters_used"))) or index
        forecast_value = get_cell_value(sheet, row, columns.get("forecast_value"))
        actual_value = get_cell_value(sheet, row, columns.get("actual_value"))
        forecast_max = get_cell_value(sheet, row, columns.get("forecast_max"))
        forecast_min = get_cell_value(sheet, row, columns.get("forecast_min"))
        intercept = intercept_by_row.get(row)
        slope = slope_by_row.get(row)

        row_signature = (
            num_quarters_used,
            round(to_float(forecast_value), 12) if to_float(forecast_value) is not None else None,
            round(to_float(forecast_max), 12) if to_float(forecast_max) is not None else None,
            round(to_float(forecast_min), 12) if to_float(forecast_min) is not None else None,
            round(to_float(intercept), 12) if to_float(intercept) is not None else None,
            round(to_float(slope), 12) if to_float(slope) is not None else None,
        )

        if index == len(rows) and previous_signature == row_signature:
            continue
        previous_signature = row_signature

        extracted_rows.append(
            {
                "model": metadata["model"],
                "ticker": metadata["ticker"],
                "model_period": metadata["model_period"],
                "model_date": metadata["model_date"],
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
        )

    return extracted_rows


def write_sheet(ws, columns: List[str], rows: List[Dict[str, Any]]) -> None:
    ws.append(columns)
    for row in rows:
        ws.append([safe_value(row.get(col_name)) for col_name in columns])

    header_font = Font(bold=True)
    for cell in ws[1]:
        cell.font = header_font

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions

    for col_idx, col_name in enumerate(columns, start=1):
        max_len = len(col_name)
        for row_idx in range(2, ws.max_row + 1):
            value = ws.cell(row=row_idx, column=col_idx).value
            if value is None:
                continue
            max_len = max(max_len, len(str(value)))
        ws.column_dimensions[get_column_letter(col_idx)].width = max(12, min(max_len + 2, 42))


def write_output_workbook(output_file: Path, empirical_rows: List[Dict[str, Any]], regression_rows: List[Dict[str, Any]]) -> None:
    workbook = Workbook()
    default_sheet = workbook.active
    workbook.remove(default_sheet)

    empirical_ws = workbook.create_sheet("empirical_candidates")
    regression_ws = workbook.create_sheet("regression_candidates")

    write_sheet(empirical_ws, EMPIRICAL_COLUMNS, empirical_rows)
    write_sheet(regression_ws, REGRESSION_COLUMNS, regression_rows)

    workbook.save(output_file)


def run() -> None:
    input_path = Path(input_dir)
    output_path = Path(output_dir)
    output_path.mkdir(parents=True, exist_ok=True)

    if not input_path.exists() or not input_path.is_dir():
        raise FileNotFoundError(f"input_dir does not exist or is not a directory: {input_path}")

    files = sorted(input_path.iterdir())
    processed_files = 0
    empirical_rows: List[Dict[str, Any]] = []
    regression_rows: List[Dict[str, Any]] = []

    app = xw.App(visible=False, add_book=False)
    app.display_alerts = False
    app.screen_updating = False
    app.calculation = "manual"

    try:
        for file_path in files:
            if not file_path.is_file():
                continue
            if file_path.name.startswith("~"):
                print(f"skipped file: {file_path.name} (temporary file)")
                continue
            if file_path.suffix.lower() != ".xlsx":
                print(f"skipped file: {file_path.name} (not .xlsx)")
                continue

            print(f"processed file: {file_path.name}")
            metadata = parse_filename_labels(file_path.name)

            try:
                wb = app.books.open(str(file_path), update_links=False)
            except Exception as exc:
                print(f"skipped file: {file_path.name} (open failed: {exc})")
                continue

            try:
                empirical_rows.extend(extract_empirical_rows(wb, metadata, file_path.name))
                regression_rows.extend(extract_regression_rows(wb, metadata, file_path.name))
                processed_files += 1
            except Exception as exc:
                print(f"skipped file: {file_path.name} (processing failed: {exc})")
            finally:
                close_workbook_safe(wb)
    finally:
        try:
            app.calculation = "automatic"
        except Exception:
            pass
        app.quit()

    output_file = build_output_path(input_path, output_path)
    write_output_workbook(output_file, empirical_rows, regression_rows)

    print(f"output path: {output_file}")
    print(f"number of files processed: {processed_files}")
    print(f"number of empirical rows: {len(empirical_rows)}")
    print(f"number of regression rows: {len(regression_rows)}")


if __name__ == "__main__":
    run()
