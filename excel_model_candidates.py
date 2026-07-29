from __future__ import annotations

import math
import re
from datetime import date, datetime
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Sequence, Tuple

import xlwings as xw
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter


# Configure these two paths before running.
input_dir = Path("./input")
output_dir = Path("./output")


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

DAY_MAP = {
    "early": 5,
    "mid": 15,
    "late": 25,
}

MONTH_MAP = {
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


def normalize_text(value: Any) -> str:
    if value is None:
        return ""
    return str(value).strip().lower()


def ensure_2d(values: Any) -> List[List[Any]]:
    if isinstance(values, (list, tuple)):
        if values and isinstance(values[0], (list, tuple)):
            return [list(row) for row in values]
        return [list(values)]
    return [[values]]


def clean_value(value: Any) -> Any:
    if value is None:
        return None
    if isinstance(value, float) and math.isnan(value):
        return None
    if isinstance(value, (date, datetime)):
        return value.isoformat()
    if isinstance(value, str):
        stripped = value.strip()
        if stripped in {"", "#N/A", "#VALUE!", "#DIV/0!", "#REF!", "#NUM!"}:
            return None
        return stripped
    return value


def to_number(value: Any) -> Optional[float]:
    value = clean_value(value)
    if value is None:
        return None
    if isinstance(value, bool):
        return float(value)
    if isinstance(value, (int, float)):
        return float(value)
    if isinstance(value, str):
        text = value.replace(",", "")
        pct = text.endswith("%")
        if pct:
            text = text[:-1]
        try:
            num = float(text)
            return num / 100.0 if pct else num
        except ValueError:
            return None
    return None


def subtract_numbers(left: Any, right: Any) -> Optional[float]:
    left_num = to_number(left)
    right_num = to_number(right)
    if left_num is None or right_num is None:
        return None
    return left_num - right_num


def maybe_rounded(value: Any, places: int = 8) -> Optional[float]:
    num = to_number(value)
    if num is None:
        return None
    return round(num, places)


def get_sheet_by_name(wb: xw.Book, name: str) -> Optional[xw.Sheet]:
    target = normalize_text(name)
    for sheet in wb.sheets:
        if normalize_text(sheet.name) == target:
            return sheet
    return None


def find_anchor_cell(sheet: xw.Sheet, anchor_text: str = "max") -> Optional[Tuple[int, int]]:
    used = sheet.used_range
    values = ensure_2d(used.value)
    start_row = used.row
    start_col = used.column
    target = normalize_text(anchor_text)

    for row_offset, row in enumerate(values):
        for col_offset, value in enumerate(row):
            if normalize_text(value) == target:
                return start_row + row_offset, start_col + col_offset
    return None


def build_header_map(sheet: xw.Sheet, row_number: int) -> Dict[int, str]:
    used = sheet.used_range
    start_col = used.column
    end_col = used.last_cell.column
    values = sheet.range((row_number, start_col), (row_number, end_col)).value
    row_values = values if isinstance(values, list) else [values]

    result: Dict[int, str] = {}
    for index, value in enumerate(row_values):
        result[start_col + index] = normalize_text(value)
    return result


def find_column(
    header_map: Dict[int, str],
    keywords: Iterable[str],
    default: Optional[int] = None,
) -> Optional[int]:
    normalized_keywords = [kw.strip().lower() for kw in keywords]
    for col, header in header_map.items():
        for keyword in normalized_keywords:
            if keyword and keyword in header:
                return col
    return default


def read_block(
    sheet: xw.Sheet,
    start_row: int,
    end_row: int,
    start_col: int,
    end_col: int,
) -> List[List[Any]]:
    if end_row < start_row or end_col < start_col:
        return []
    values = sheet.range((start_row, start_col), (end_row, end_col)).value
    return ensure_2d(values)


def block_value(
    block: List[List[Any]],
    block_start_row: int,
    block_start_col: int,
    target_row: int,
    target_col: Optional[int],
) -> Any:
    if target_col is None or not block:
        return None

    row_index = target_row - block_start_row
    col_index = target_col - block_start_col
    if row_index < 0 or col_index < 0:
        return None
    if row_index >= len(block):
        return None
    row = block[row_index]
    if col_index >= len(row):
        return None
    return clean_value(row[col_index])


def parse_model_metadata(file_path: Path) -> Dict[str, str]:
    stem = file_path.stem
    pattern = re.compile(
        r".*-\s*(?P<ticker>[A-Za-z0-9]+)\s*-\s*(?P<phase>Early|Mid|Late)"
        r"(?P<month>[A-Za-z]{3,9})(?P<year>\d{4})_Send",
        re.IGNORECASE,
    )
    match = pattern.search(stem)

    if match:
        ticker = match.group("ticker").upper()
        phase_raw = match.group("phase")
        phase = phase_raw[0].upper() + phase_raw[1:].lower()
        month_token = match.group("month")[:3].lower()
        year = match.group("year")
        month_num = MONTH_MAP.get(month_token)
        day = DAY_MAP.get(phase.lower())

        month_title = month_token.title()
        model_period = f"{phase}{month_title}_{year}"
        model_date = ""
        if month_num and day:
            model_date = date(int(year), month_num, day).isoformat()
        model = f"{ticker}_{model_period}"

        return {
            "model": model,
            "ticker": ticker,
            "model_period": model_period,
            "model_date": model_date,
        }

    parts = [segment.strip() for segment in stem.split("-")]
    ticker = parts[1].upper() if len(parts) >= 2 else stem.upper()
    period_raw = parts[2] if len(parts) >= 3 else ""
    period_clean = period_raw.replace("_Send", "").replace(" ", "")
    model_period = period_clean or "UNKNOWN_PERIOD"
    model = f"{ticker}_{model_period}"
    return {
        "model": model,
        "ticker": ticker,
        "model_period": model_period,
        "model_date": "",
    }


def close_without_saving(workbook: xw.Book) -> None:
    try:
        workbook.close(save=False)
        return
    except TypeError:
        pass
    except Exception:
        pass

    try:
        workbook.api.Close(False)
    except Exception:
        try:
            workbook.close()
        except Exception:
            pass


def next_output_path(input_path: Path, output_path: Path) -> Path:
    output_path.mkdir(parents=True, exist_ok=True)
    input_folder_name = input_path.name or "input"
    base_name = f"{input_folder_name}_PARAM"
    candidate = output_path / f"{base_name}.xlsx"

    index = 1
    while candidate.exists():
        candidate = output_path / f"{base_name}.{index}.xlsx"
        index += 1
    return candidate


def extract_empirical_rows(
    workbook: xw.Book,
    metadata: Dict[str, str],
    source_file: str,
) -> List[Dict[str, Any]]:
    sheet = get_sheet_by_name(workbook, EMPIRICAL_MODEL_SHEET)
    if sheet is None:
        print(f"Skipped empirical for {source_file}: sheet not found")
        return []

    anchor = find_anchor_cell(sheet, "max")
    if anchor is None:
        print(f"Skipped empirical for {source_file}: 'max' anchor not found")
        return []

    anchor_row, anchor_col = anchor
    used = sheet.used_range
    header_map = build_header_map(sheet, anchor_row)

    num_quarters_col = find_column(header_map, ["num quarter", "quarters used", "# quarters", "n quarters"])
    last_quarter_col = find_column(header_map, ["last quarter"])
    forecast_value_col = find_column(
        header_map,
        ["estimated total sold", "forecast value", "forecast total", "tot fcst", "total sold"],
    )
    actual_value_col = find_column(header_map, ["reported sales", "actual value", "actual"])
    forecast_max_col = anchor_col
    forecast_min_col = find_column(header_map, ["min"], default=anchor_col + 1)
    quarterly_sales_col = find_column(header_map, ["quarterly sales", "quarter sales"])
    reported_sales_col = find_column(header_map, ["reported sales"])
    growth_rate_col = find_column(header_map, ["growth rate"])
    sales_captured_col = find_column(
        header_map,
        ["sales captured in db", "captured in db", "captured_in_db", "penetration"],
    )

    penetration_col = sales_captured_col
    if penetration_col is None:
        penetration_col = max(1, anchor_col - 5)

    history_end_row = anchor_row - 1
    history_start_limit = used.row

    staging_row = used.last_cell.row + 5
    staging_col = used.last_cell.column + 2

    windows: List[Tuple[int, int, int]] = []
    for n_quarters in range(1, N_QUARTERS + 1):
        start_row = history_end_row - n_quarters + 1
        if start_row < history_start_limit:
            break
        target_row = staging_row + len(windows)
        formula = f"=AVERAGE(R{start_row}C{penetration_col}:R{history_end_row}C{penetration_col})"
        sheet.cells(target_row, staging_col).formula2 = formula
        windows.append((n_quarters, start_row, target_row))

    if windows:
        workbook.app.calculate()

    table_start_row = anchor_row + 1
    table_end_row = anchor_row + N_QUARTERS
    table_cols = [
        col
        for col in [
            num_quarters_col,
            last_quarter_col,
            forecast_value_col,
            actual_value_col,
            forecast_max_col,
            forecast_min_col,
            quarterly_sales_col,
            reported_sales_col,
            growth_rate_col,
            sales_captured_col,
        ]
        if col is not None and col >= 1
    ]

    block: List[List[Any]] = []
    block_start_col = 1
    if table_cols:
        block_start_col = min(table_cols)
        block_end_col = max(table_cols)
        block = read_block(sheet, table_start_row, table_end_row, block_start_col, block_end_col)

    fallback_last_quarter = clean_value(sheet.cells(history_end_row, used.column).value)
    rows: List[Dict[str, Any]] = []

    for n_quarters, history_start_row, avg_row in windows:
        output_row = anchor_row + n_quarters

        raw_num_quarters = block_value(block, table_start_row, block_start_col, output_row, num_quarters_col)
        if num_quarters_col is not None and raw_num_quarters is None:
            break
        num_quarters_used = raw_num_quarters if raw_num_quarters is not None else n_quarters

        last_quarter_used = block_value(block, table_start_row, block_start_col, output_row, last_quarter_col)
        if last_quarter_used is None:
            last_quarter_used = fallback_last_quarter

        forecast_value = block_value(block, table_start_row, block_start_col, output_row, forecast_value_col)
        actual_value = block_value(block, table_start_row, block_start_col, output_row, actual_value_col)
        forecast_max = block_value(block, table_start_row, block_start_col, output_row, forecast_max_col)
        forecast_min = block_value(block, table_start_row, block_start_col, output_row, forecast_min_col)
        range_width = subtract_numbers(forecast_max, forecast_min)

        quarterly_sales = block_value(block, table_start_row, block_start_col, output_row, quarterly_sales_col)
        reported_sales = block_value(block, table_start_row, block_start_col, output_row, reported_sales_col)
        growth_rate = block_value(block, table_start_row, block_start_col, output_row, growth_rate_col)
        sales_captured = block_value(block, table_start_row, block_start_col, output_row, sales_captured_col)

        avg_penetration = clean_value(sheet.cells(avg_row, staging_col).value)
        if sales_captured is None:
            sales_captured = avg_penetration

        row = {
            "model": metadata["model"],
            "ticker": metadata["ticker"],
            "model_period": metadata["model_period"],
            "model_date": metadata["model_date"],
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
            "growth_rate_pct": growth_rate,
            "sales_captured_in_db_pct": sales_captured,
            "source_file": source_file,
        }
        rows.append(row)

    return rows


def extract_regression_rows(
    workbook: xw.Book,
    metadata: Dict[str, str],
    source_file: str,
) -> List[Dict[str, Any]]:
    sheet = get_sheet_by_name(workbook, REGRESSION_MODEL_SHEET)
    if sheet is None:
        print(f"Skipped regression for {source_file}: sheet not found")
        return []

    anchor = find_anchor_cell(sheet, "max")
    if anchor is None:
        print(f"Skipped regression for {source_file}: 'max' anchor not found")
        return []

    anchor_row, anchor_col = anchor
    used = sheet.used_range
    header_map = build_header_map(sheet, anchor_row)

    x_col = max(1, anchor_col - 11)
    y_col = max(1, anchor_col - 7)

    num_quarters_col = find_column(header_map, ["num quarter", "quarters used", "# quarters", "n quarters"])
    forecast_total_col = find_column(
        header_map,
        ["tot fcst w/o sa", "total fcst w/o sa", "forecast total without sa", "forecast total", "tot fcst"],
        default=anchor_col - 1,
    )
    actual_value_col = find_column(header_map, ["actual value", "actual", "reported sales"])
    forecast_max_col = anchor_col
    forecast_min_col = find_column(header_map, ["min"], default=anchor_col + 1)

    history_end_row = anchor_row - 1
    history_start_limit = used.row

    next_x = to_number(clean_value(sheet.cells(history_end_row + 1, x_col).value))
    if next_x is None:
        next_x = to_number(clean_value(sheet.cells(history_end_row, x_col).value))

    staging_row = used.last_cell.row + 5
    staging_col = used.last_cell.column + 4

    windows: List[Tuple[int, int, int]] = []
    for n_quarters in range(1, N_QUARTERS + 1):
        start_row = history_end_row - n_quarters + 1
        if start_row < history_start_limit:
            break

        target_row = staging_row + len(windows)
        intercept_formula = f"=INTERCEPT(R{start_row}C{y_col}:R{history_end_row}C{y_col},R{start_row}C{x_col}:R{history_end_row}C{x_col})"
        slope_formula = f"=SLOPE(R{start_row}C{y_col}:R{history_end_row}C{y_col},R{start_row}C{x_col}:R{history_end_row}C{x_col})"
        sheet.cells(target_row, staging_col).formula2 = intercept_formula
        sheet.cells(target_row, staging_col + 1).formula2 = slope_formula

        if next_x is None:
            sheet.cells(target_row, staging_col + 2).formula2 = "=NA()"
        else:
            sheet.cells(target_row, staging_col + 2).formula2 = f"=RC[-2]+RC[-1]*{next_x}"

        windows.append((n_quarters, start_row, target_row))

    if windows:
        workbook.app.calculate()

    table_start_row = anchor_row + 1
    table_end_row = anchor_row + N_QUARTERS
    table_cols = [
        col
        for col in [
            num_quarters_col,
            forecast_total_col,
            actual_value_col,
            forecast_max_col,
            forecast_min_col,
        ]
        if col is not None and col >= 1
    ]

    block: List[List[Any]] = []
    block_start_col = 1
    if table_cols:
        block_start_col = min(table_cols)
        block_end_col = max(table_cols)
        block = read_block(sheet, table_start_row, table_end_row, block_start_col, block_end_col)

    rows: List[Dict[str, Any]] = []
    previous_signature: Optional[Tuple[Any, ...]] = None

    for n_quarters, _, calc_row in windows:
        output_row = anchor_row + n_quarters

        raw_num_quarters = block_value(block, table_start_row, block_start_col, output_row, num_quarters_col)
        if num_quarters_col is not None and raw_num_quarters is None:
            break
        num_quarters_used = raw_num_quarters if raw_num_quarters is not None else n_quarters

        intercept = clean_value(sheet.cells(calc_row, staging_col).value)
        slope = clean_value(sheet.cells(calc_row, staging_col + 1).value)
        forecast_calc = clean_value(sheet.cells(calc_row, staging_col + 2).value)

        forecast_value = block_value(block, table_start_row, block_start_col, output_row, forecast_total_col)
        if forecast_value is None:
            forecast_value = forecast_calc

        actual_value = block_value(block, table_start_row, block_start_col, output_row, actual_value_col)
        if actual_value is None:
            actual_value = ""

        forecast_max = block_value(block, table_start_row, block_start_col, output_row, forecast_max_col)
        forecast_min = block_value(block, table_start_row, block_start_col, output_row, forecast_min_col)
        range_width = subtract_numbers(forecast_max, forecast_min)

        signature = (
            maybe_rounded(num_quarters_used),
            maybe_rounded(forecast_value),
            maybe_rounded(intercept),
            maybe_rounded(slope),
            maybe_rounded(forecast_max),
            maybe_rounded(forecast_min),
        )
        if previous_signature is not None and signature == previous_signature:
            continue
        previous_signature = signature

        row = {
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
            "range_width": range_width,
            "intercept": intercept,
            "slope": slope,
            "source_file": source_file,
        }
        rows.append(row)

    return rows


def write_rows_to_sheet(workbook_sheet, headers: Sequence[str], rows: List[Dict[str, Any]]) -> None:
    workbook_sheet.append(list(headers))
    for item in rows:
        workbook_sheet.append([item.get(col) for col in headers])

    for cell in workbook_sheet[1]:
        cell.font = Font(bold=True)

    workbook_sheet.freeze_panes = "A2"
    workbook_sheet.auto_filter.ref = workbook_sheet.dimensions

    for col_index in range(1, len(headers) + 1):
        values = [
            workbook_sheet.cell(row=row_index, column=col_index).value
            for row_index in range(1, workbook_sheet.max_row + 1)
        ]
        max_width = max((len(str(value)) for value in values if value is not None), default=0)
        workbook_sheet.column_dimensions[get_column_letter(col_index)].width = min(max(max_width + 2, 12), 48)


def write_output_workbook(
    output_path: Path,
    empirical_rows: List[Dict[str, Any]],
    regression_rows: List[Dict[str, Any]],
) -> None:
    wb = Workbook()
    default_sheet = wb.active
    wb.remove(default_sheet)

    empirical_sheet = wb.create_sheet(EMPIRICAL_OUTPUT_SHEET)
    write_rows_to_sheet(empirical_sheet, EMPIRICAL_COLUMNS, empirical_rows)

    regression_sheet = wb.create_sheet(REGRESSION_OUTPUT_SHEET)
    write_rows_to_sheet(regression_sheet, REGRESSION_COLUMNS, regression_rows)

    wb.save(output_path)


def main() -> None:
    source_dir = Path(input_dir).expanduser().resolve()
    target_dir = Path(output_dir).expanduser().resolve()

    if not source_dir.exists():
        raise FileNotFoundError(f"input_dir does not exist: {source_dir}")
    if not source_dir.is_dir():
        raise NotADirectoryError(f"input_dir is not a directory: {source_dir}")

    output_path = next_output_path(source_dir, target_dir)

    empirical_rows: List[Dict[str, Any]] = []
    regression_rows: List[Dict[str, Any]] = []
    files_processed = 0

    app = xw.App(visible=False, add_book=False)
    app.display_alerts = False
    app.screen_updating = False

    try:
        entries = sorted(source_dir.iterdir(), key=lambda path: path.name.lower())
        for file_path in entries:
            if not file_path.is_file():
                print(f"Skipped {file_path.name}: not a file")
                continue
            if file_path.name.startswith("~"):
                print(f"Skipped {file_path.name}: temporary file")
                continue
            if file_path.suffix.lower() != ".xlsx":
                print(f"Skipped {file_path.name}: not .xlsx")
                continue

            print(f"Processing {file_path.name}")
            metadata = parse_model_metadata(file_path)
            workbook: Optional[xw.Book] = None

            try:
                workbook = app.books.open(str(file_path), update_links=False)
                empirical_rows.extend(extract_empirical_rows(workbook, metadata, file_path.name))
                regression_rows.extend(extract_regression_rows(workbook, metadata, file_path.name))
                files_processed += 1
            except Exception as exc:
                print(f"Skipped {file_path.name}: processing error ({exc})")
            finally:
                if workbook is not None:
                    close_without_saving(workbook)
    finally:
        app.quit()

    write_output_workbook(output_path, empirical_rows, regression_rows)

    print(f"Output path: {output_path}")
    print(f"Files processed: {files_processed}")
    print(f"Empirical rows: {len(empirical_rows)}")
    print(f"Regression rows: {len(regression_rows)}")


if __name__ == "__main__":
    main()
