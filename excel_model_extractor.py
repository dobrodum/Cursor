from __future__ import annotations

import re
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List, Optional, Sequence, Tuple

import xlwings as xw
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

# Configure these two paths before running.
input_dir = Path("/path/to/input")
output_dir = Path("/path/to/output")

N_QUARTERS = 10
EXCEL_MAX_ROW = 1_048_576
EXCEL_MAX_COL = 16_384

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

FILE_PATTERN = re.compile(
    r"^\s*.+?\s*-\s*(?P<ticker>[A-Za-z0-9]+)\s*-\s*(?P<period>[A-Za-z]+\d{4})(?:_.+)?\s*$",
    re.IGNORECASE,
)
PERIOD_PATTERN = re.compile(r"^(Early|Mid|Late)([A-Za-z]+)(\d{4})$", re.IGNORECASE)
PERIOD_DAY_MAP = {"early": 5, "mid": 15, "late": 25}


def normalize_text(value: Any) -> str:
    if value is None:
        return ""
    return re.sub(r"[^a-z0-9]+", "", str(value).strip().lower())


def as_float(value: Any) -> Optional[float]:
    if value is None or value == "":
        return None
    try:
        return float(value)
    except (TypeError, ValueError):
        return None


def as_int(value: Any) -> Optional[int]:
    number = as_float(value)
    if number is None:
        return None
    return int(round(number))


def clean_value(value: Any) -> Any:
    if isinstance(value, float) and value.is_integer():
        return int(value)
    return value


def to_2d(values: Any) -> List[List[Any]]:
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


def parse_file_label(file_path: Path) -> Dict[str, str]:
    stem = file_path.stem
    ticker = "UNKNOWN"
    period_token = ""

    match = FILE_PATTERN.match(stem)
    if match:
        ticker = match.group("ticker").upper()
        period_token = match.group("period")
    else:
        parts = [part.strip() for part in stem.split("-")]
        if len(parts) >= 2 and parts[1]:
            ticker = parts[1].upper()
        if len(parts) >= 3 and parts[2]:
            period_token = parts[2].split("_")[0].strip()

    model_period = "UNKNOWN"
    model_date = ""
    period_match = PERIOD_PATTERN.match(period_token)
    if period_match:
        stage = period_match.group(1).title()
        month_token = period_match.group(2)
        year = period_match.group(3)
        month_abbr = month_token[:3].title()
        month_number = 1
        try:
            month_number = datetime.strptime(month_abbr, "%b").month
        except ValueError:
            month_number = 1
        day = PERIOD_DAY_MAP.get(stage.lower(), 15)
        model_period = f"{stage}{month_abbr}_{year}"
        model_date = f"{year}-{month_number:02d}-{day:02d}"

    model = f"{ticker}_{model_period}" if model_period != "UNKNOWN" else ticker
    return {
        "model": model,
        "ticker": ticker,
        "model_period": model_period,
        "model_date": model_date,
    }


def next_output_path(source_input_dir: Path, target_output_dir: Path) -> Path:
    base_name = f"{source_input_dir.name}_PARAM"
    candidate = target_output_dir / f"{base_name}.xlsx"
    if not candidate.exists():
        return candidate

    suffix = 1
    while True:
        candidate = target_output_dir / f"{base_name}.{suffix}.xlsx"
        if not candidate.exists():
            return candidate
        suffix += 1


def workbook_sheet(book: xw.Book, sheet_name: str) -> Optional[xw.Sheet]:
    try:
        return book.sheets[sheet_name]
    except Exception:
        return None


def cell_value(sheet: xw.Sheet, row: int, col: int) -> Any:
    if row < 1 or col < 1:
        return None
    return sheet.range((row, col)).value


def find_anchor(sheet: xw.Sheet, anchor_text: str = "max") -> Optional[Tuple[int, int]]:
    used_range = sheet.used_range
    start_row = used_range.row
    start_col = used_range.column
    matrix = to_2d(used_range.value)
    target = anchor_text.strip().lower()

    for row_offset, row_values in enumerate(matrix):
        for col_offset, value in enumerate(row_values):
            if isinstance(value, str) and value.strip().lower() == target:
                return start_row + row_offset, start_col + col_offset
    return None


def read_header_window(
    sheet: xw.Sheet,
    header_row: int,
    anchor_col: int,
    span: int = 40,
) -> Dict[int, str]:
    min_col = max(1, anchor_col - span)
    max_col = min(EXCEL_MAX_COL, anchor_col + span)
    values = sheet.range((header_row, min_col), (header_row, max_col)).value
    if not isinstance(values, list):
        values = [values]

    headers: Dict[int, str] = {}
    for index, value in enumerate(values):
        if value is None:
            continue
        label = str(value).strip()
        if not label:
            continue
        headers[min_col + index] = label
    return headers


def pick_column(
    headers: Dict[int, str],
    anchor_col: int,
    default_offset: int,
    patterns: Sequence[str],
    excluded: Optional[Sequence[int]] = None,
) -> int:
    excluded_set = set(excluded or [])
    sorted_columns = sorted(headers.items(), key=lambda item: abs(item[0] - anchor_col))

    for pattern in patterns:
        normalized_pattern = normalize_text(pattern)
        for col, label in sorted_columns:
            if col in excluded_set:
                continue
            if normalized_pattern and normalized_pattern in normalize_text(label):
                return col
    return anchor_col + default_offset


def find_data_bounds(
    sheet: xw.Sheet,
    columns: Sequence[int],
    start_row: int,
    max_rows: int = 5_000,
    blank_streak_to_stop: int = 30,
) -> Tuple[Optional[int], Optional[int]]:
    checked_cols = [col for col in columns if col >= 1]
    if not checked_cols:
        return None, None

    first_row: Optional[int] = None
    last_row: Optional[int] = None
    blanks = 0
    end_row = min(EXCEL_MAX_ROW, start_row + max_rows - 1)

    for row in range(start_row, end_row + 1):
        row_has_data = False
        for col in checked_cols:
            if as_float(cell_value(sheet, row, col)) is not None:
                row_has_data = True
                break
        if row_has_data:
            if first_row is None:
                first_row = row
            last_row = row
            blanks = 0
            continue

        if first_row is not None:
            blanks += 1
            if blanks >= blank_streak_to_stop:
                break

    return first_row, last_row


def scratch_cells(sheet: xw.Sheet, count: int) -> List[xw.Range]:
    used_last_row = sheet.used_range.last_cell.row
    used_last_col = sheet.used_range.last_cell.column
    row = min(EXCEL_MAX_ROW, used_last_row + 4)
    start_col = min(EXCEL_MAX_COL, used_last_col + 4)

    cells: List[xw.Range] = []
    for offset in range(count):
        col = min(EXCEL_MAX_COL, start_col + offset)
        cells.append(sheet.range((row, col)))
    return cells


def safe_close_workbook(book: xw.Book) -> None:
    try:
        book.close(save=False)
        return
    except TypeError:
        pass
    except Exception:
        pass

    try:
        book.close(False)
        return
    except Exception:
        pass

    try:
        book.api.Close(SaveChanges=False)
        return
    except Exception:
        pass

    try:
        book.api.Close(False)
    except Exception:
        pass


def calc_range_width(max_value: Any, min_value: Any) -> Any:
    max_num = as_float(max_value)
    min_num = as_float(min_value)
    if max_num is None or min_num is None:
        return None
    width = max_num - min_num
    if width.is_integer():
        return int(width)
    return width


def signature_value(value: Any) -> Any:
    number = as_float(value)
    if number is None:
        return clean_value(value)
    return round(number, 10)


def extract_empirical_rows(
    sheet: xw.Sheet,
    model_meta: Dict[str, str],
    source_file: str,
) -> List[Dict[str, Any]]:
    anchor = find_anchor(sheet, "max")
    if anchor is None:
        print(f"  skipped empirical extraction for {source_file}: 'max' anchor not found")
        return []

    anchor_row, anchor_col = anchor
    headers = read_header_window(sheet, anchor_row, anchor_col)

    forecast_max_col = anchor_col
    forecast_min_col = pick_column(headers, anchor_col, 1, ["min"], excluded=[forecast_max_col])
    forecast_value_col = pick_column(
        headers,
        anchor_col,
        -1,
        ["estimated total sold", "tot fcst", "forecast total", "fcst w/o sa"],
        excluded=[forecast_max_col, forecast_min_col],
    )
    actual_col = pick_column(
        headers,
        anchor_col,
        -2,
        ["actual", "reported sales"],
        excluded=[forecast_max_col, forecast_min_col, forecast_value_col],
    )
    num_quarters_col = pick_column(
        headers,
        anchor_col,
        -3,
        ["num quarters", "quarters used", "n quarters"],
    )
    last_quarter_col = pick_column(headers, anchor_col, -4, ["last quarter"])
    quarterly_sales_col = pick_column(headers, anchor_col, -5, ["quarterly sales", "quarter sales"])
    reported_sales_col = pick_column(headers, anchor_col, -6, ["reported sales"])
    growth_rate_col = pick_column(headers, anchor_col, -7, ["growth rate"])
    captured_col = pick_column(headers, anchor_col, -8, ["captured in db", "sales captured"])
    penetration_col = pick_column(headers, anchor_col, -9, ["penetration"])

    first_data_row, last_data_row = find_data_bounds(
        sheet,
        [forecast_max_col, forecast_min_col, forecast_value_col, num_quarters_col],
        anchor_row + 1,
    )
    if first_data_row is None:
        print(f"  skipped empirical extraction for {source_file}: no numeric rows found")
        return []

    history_start, history_end = find_data_bounds(
        sheet,
        [penetration_col],
        first_data_row,
    )
    if history_start is None or history_end is None:
        history_start, history_end = first_data_row, last_data_row or first_data_row

    avg_cell, forecast_cell = scratch_cells(sheet, 2)
    avg_row, avg_col = avg_cell.row, avg_cell.column

    output_rows: List[Dict[str, Any]] = []
    for index in range(N_QUARTERS):
        row = first_data_row + index
        num_quarters_used = as_int(cell_value(sheet, row, num_quarters_col)) or (index + 1)
        lookback = max(1, num_quarters_used)
        start_row = max(history_start, history_end - lookback + 1)

        if penetration_col >= 1:
            avg_cell.formula2 = (
                f'=IFERROR(AVERAGE(R{start_row}C{penetration_col}:R{history_end}C{penetration_col}),"")'
            )
        else:
            avg_cell.formula2 = '=""'

        if quarterly_sales_col >= 1:
            forecast_cell.formula2 = (
                f'=IFERROR(R{row}C{quarterly_sales_col}*R{avg_row}C{avg_col},"")'
            )
        else:
            forecast_cell.formula2 = '=""'

        sheet.book.app.calculate()

        forecast_max = clean_value(cell_value(sheet, row, forecast_max_col))
        forecast_min = clean_value(cell_value(sheet, row, forecast_min_col))
        forecast_value = clean_value(cell_value(sheet, row, forecast_value_col))
        quarterly_sales = clean_value(cell_value(sheet, row, quarterly_sales_col))
        reported_sales = clean_value(cell_value(sheet, row, reported_sales_col))
        actual_value = clean_value(cell_value(sheet, row, actual_col))

        if forecast_value in (None, ""):
            forecast_value = clean_value(forecast_cell.value)
        if actual_value in (None, ""):
            actual_value = reported_sales

        avg_penetration = clean_value(avg_cell.value)
        last_quarter_used = clean_value(cell_value(sheet, row, last_quarter_col))
        growth_rate = clean_value(cell_value(sheet, row, growth_rate_col))
        sales_captured = clean_value(cell_value(sheet, row, captured_col))

        if all(
            value in (None, "")
            for value in (
                forecast_max,
                forecast_min,
                forecast_value,
                actual_value,
                quarterly_sales,
                reported_sales,
            )
        ):
            continue

        output_rows.append(
            {
                "model": model_meta["model"],
                "ticker": model_meta["ticker"],
                "model_period": model_meta["model_period"],
                "model_date": model_meta["model_date"],
                "method": "empirical",
                "parameter_name": "avg_penetration_pct",
                "parameter_value": avg_penetration,
                "num_quarters_used": num_quarters_used,
                "last_quarter_used": last_quarter_used,
                "forecast_value": forecast_value,
                "actual_value": actual_value,
                "forecast_max": forecast_max,
                "forecast_min": forecast_min,
                "range_width": calc_range_width(forecast_max, forecast_min),
                "avg_penetration_pct": avg_penetration,
                "quarterly_sales": quarterly_sales,
                "reported_sales": reported_sales,
                "growth_rate_pct": growth_rate,
                "sales_captured_in_db_pct": sales_captured,
                "source_file": source_file,
            }
        )

    return output_rows


def extract_regression_rows(
    sheet: xw.Sheet,
    model_meta: Dict[str, str],
    source_file: str,
) -> List[Dict[str, Any]]:
    anchor = find_anchor(sheet, "max")
    if anchor is None:
        print(f"  skipped regression extraction for {source_file}: 'max' anchor not found")
        return []

    anchor_row, anchor_col = anchor
    headers = read_header_window(sheet, anchor_row, anchor_col)

    forecast_max_col = anchor_col
    forecast_min_col = pick_column(headers, anchor_col, 1, ["min"], excluded=[forecast_max_col])
    forecast_value_col = pick_column(
        headers,
        anchor_col,
        -1,
        ["tot fcst w/o sa", "tot fcst wo sa", "forecast w/o sa", "forecast total"],
        excluded=[forecast_max_col, forecast_min_col],
    )
    actual_col = pick_column(
        headers,
        anchor_col,
        -2,
        ["actual", "reported sales"],
        excluded=[forecast_max_col, forecast_min_col, forecast_value_col],
    )
    num_quarters_col = pick_column(
        headers,
        anchor_col,
        -3,
        ["num quarters", "quarters used", "n quarters"],
    )

    y_col = anchor_col - 7
    x_col = anchor_col - 11

    first_data_row, _ = find_data_bounds(
        sheet,
        [forecast_max_col, forecast_min_col, forecast_value_col, num_quarters_col],
        anchor_row + 1,
    )
    if first_data_row is None:
        print(f"  skipped regression extraction for {source_file}: no numeric rows found")
        return []

    history_start, history_end = find_data_bounds(
        sheet,
        [x_col, y_col],
        anchor_row + 1,
    )
    if history_start is None or history_end is None:
        print(f"  skipped regression extraction for {source_file}: x/y history not found")
        return []

    intercept_cell, slope_cell = scratch_cells(sheet, 2)
    output_rows: List[Dict[str, Any]] = []
    previous_signature: Optional[Tuple[Any, ...]] = None

    for index in range(N_QUARTERS):
        row = first_data_row + index
        num_quarters_used = as_int(cell_value(sheet, row, num_quarters_col)) or (index + 1)
        lookback = max(2, num_quarters_used)
        start_row = max(history_start, history_end - lookback + 1)

        intercept_cell.formula2 = (
            f'=IFERROR(INTERCEPT(R{start_row}C{y_col}:R{history_end}C{y_col},'
            f"R{start_row}C{x_col}:R{history_end}C{x_col}),\"\")"
        )
        slope_cell.formula2 = (
            f'=IFERROR(SLOPE(R{start_row}C{y_col}:R{history_end}C{y_col},'
            f"R{start_row}C{x_col}:R{history_end}C{x_col}),\"\")"
        )
        sheet.book.app.calculate()

        intercept_value = clean_value(intercept_cell.value)
        slope_value = clean_value(slope_cell.value)
        forecast_value = clean_value(cell_value(sheet, row, forecast_value_col))
        actual_value = clean_value(cell_value(sheet, row, actual_col))
        forecast_max = clean_value(cell_value(sheet, row, forecast_max_col))
        forecast_min = clean_value(cell_value(sheet, row, forecast_min_col))

        if all(
            value in (None, "")
            for value in (
                forecast_value,
                forecast_max,
                forecast_min,
                intercept_value,
                slope_value,
            )
        ):
            continue

        signature = (
            num_quarters_used,
            signature_value(intercept_value),
            signature_value(slope_value),
            signature_value(forecast_value),
            signature_value(forecast_max),
            signature_value(forecast_min),
        )
        if signature == previous_signature:
            continue
        previous_signature = signature

        output_rows.append(
            {
                "model": model_meta["model"],
                "ticker": model_meta["ticker"],
                "model_period": model_meta["model_period"],
                "model_date": model_meta["model_date"],
                "method": "regression",
                "parameter_name": "num_quarters_used",
                "parameter_value": num_quarters_used,
                "num_quarters_used": num_quarters_used,
                "forecast_value": forecast_value,
                "actual_value": actual_value,
                "forecast_max": forecast_max,
                "forecast_min": forecast_min,
                "range_width": calc_range_width(forecast_max, forecast_min),
                "intercept": intercept_value,
                "slope": slope_value,
                "source_file": source_file,
            }
        )

    return output_rows


def write_sheet(
    workbook_sheet_obj,
    headers: Sequence[str],
    rows: Sequence[Dict[str, Any]],
) -> None:
    workbook_sheet_obj.append(list(headers))
    for row in rows:
        workbook_sheet_obj.append([row.get(header) for header in headers])

    for cell in workbook_sheet_obj[1]:
        cell.font = Font(bold=True)

    workbook_sheet_obj.freeze_panes = "A2"
    max_row = max(workbook_sheet_obj.max_row, 1)
    max_col = len(headers)
    workbook_sheet_obj.auto_filter.ref = f"A1:{get_column_letter(max_col)}{max_row}"

    for col_index, header in enumerate(headers, start=1):
        width = len(header)
        for row_index in range(2, workbook_sheet_obj.max_row + 1):
            value = workbook_sheet_obj.cell(row=row_index, column=col_index).value
            if value is None:
                continue
            width = max(width, len(str(value)))
        workbook_sheet_obj.column_dimensions[get_column_letter(col_index)].width = min(width + 2, 40)


def write_output_workbook(
    destination_path: Path,
    empirical_rows: Sequence[Dict[str, Any]],
    regression_rows: Sequence[Dict[str, Any]],
) -> None:
    workbook = Workbook()
    default_sheet = workbook.active
    workbook.remove(default_sheet)

    empirical_sheet = workbook.create_sheet("empirical_candidates")
    regression_sheet = workbook.create_sheet("regression_candidates")

    write_sheet(empirical_sheet, EMPIRICAL_COLUMNS, empirical_rows)
    write_sheet(regression_sheet, REGRESSION_COLUMNS, regression_rows)

    workbook.save(destination_path)


def iter_source_files(source_input_dir: Path) -> List[Path]:
    return sorted(path for path in source_input_dir.iterdir() if path.is_file())


def run() -> None:
    if not input_dir.exists() or not input_dir.is_dir():
        raise FileNotFoundError(f"input_dir does not exist or is not a directory: {input_dir}")

    output_dir.mkdir(parents=True, exist_ok=True)
    output_path = next_output_path(input_dir, output_dir)

    empirical_rows: List[Dict[str, Any]] = []
    regression_rows: List[Dict[str, Any]] = []
    processed_files = 0

    app = xw.App(visible=False, add_book=False)
    app.display_alerts = False
    app.screen_updating = False
    app.calculation = "manual"

    try:
        for file_path in iter_source_files(input_dir):
            if file_path.name.startswith("~"):
                print(f"skipped {file_path.name}: temp file")
                continue
            if file_path.suffix.lower() != ".xlsx":
                print(f"skipped {file_path.name}: not an .xlsx file")
                continue

            print(f"processing {file_path.name}")
            workbook: Optional[xw.Book] = None
            try:
                workbook = app.books.open(str(file_path), update_links=False)
                model_meta = parse_file_label(file_path)

                empirical_sheet = workbook_sheet(workbook, "Empirical Model")
                if empirical_sheet is None:
                    print(f"  skipped empirical extraction for {file_path.name}: sheet missing")
                else:
                    empirical_rows.extend(
                        extract_empirical_rows(empirical_sheet, model_meta, file_path.name)
                    )

                regression_sheet = workbook_sheet(workbook, "Regression Model")
                if regression_sheet is None:
                    print(f"  skipped regression extraction for {file_path.name}: sheet missing")
                else:
                    regression_rows.extend(
                        extract_regression_rows(regression_sheet, model_meta, file_path.name)
                    )

                processed_files += 1
                print(f"processed {file_path.name}")
            except Exception as exc:
                print(f"skipped {file_path.name}: {exc}")
            finally:
                if workbook is not None:
                    safe_close_workbook(workbook)
    finally:
        app.quit()

    write_output_workbook(output_path, empirical_rows, regression_rows)
    print(f"output path: {output_path}")
    print(f"files processed: {processed_files}")
    print(f"empirical rows: {len(empirical_rows)}")
    print(f"regression rows: {len(regression_rows)}")


if __name__ == "__main__":
    run()
