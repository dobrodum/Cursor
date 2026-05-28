from __future__ import annotations

import re
from datetime import date
from pathlib import Path
from statistics import mean
from typing import Any, Dict, List, Optional, Sequence, Tuple

import xlwings as xw
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

# Configure these two paths before running.
input_dir = Path("./input")
output_dir = Path("./output")

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

DAY_BY_LABEL = {"early": 5, "mid": 15, "late": 25}
MONTH_BY_TEXT = {
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


def normalize_text(value: Any) -> str:
    if value is None:
        return ""
    return str(value).strip().lower()


def is_blank(value: Any) -> bool:
    return value is None or (isinstance(value, str) and not value.strip())


def to_float(value: Any) -> Optional[float]:
    if value is None or value == "":
        return None
    if isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        return float(value)
    text = str(value).strip().replace(",", "")
    if not text:
        return None
    try:
        return float(text)
    except ValueError:
        return None


def to_int(value: Any) -> Optional[int]:
    numeric = to_float(value)
    if numeric is None:
        return None
    return int(round(numeric))


def to_2d(values: Any) -> List[List[Any]]:
    if values is None:
        return []
    if not isinstance(values, list):
        return [[values]]
    if values and not isinstance(values[0], list):
        return [values]
    return values


def get_matrix_value(
    matrix: Sequence[Sequence[Any]],
    top_row: int,
    left_col: int,
    row: int,
    col: int,
) -> Any:
    row_idx = row - top_row
    col_idx = col - left_col
    if row_idx < 0 or col_idx < 0:
        return None
    if row_idx >= len(matrix):
        return None
    row_values = matrix[row_idx]
    if col_idx >= len(row_values):
        return None
    return row_values[col_idx]


def find_anchor_max(sheet: xw.Sheet) -> Optional[Tuple[int, int, List[List[Any]], int, int]]:
    used = sheet.used_range
    matrix = to_2d(used.value)
    if not matrix:
        return None

    top_row = used.row
    left_col = used.column

    exact_match: Optional[Tuple[int, int]] = None
    contains_match: Optional[Tuple[int, int]] = None

    for r_idx, row_values in enumerate(matrix):
        for c_idx, cell_value in enumerate(row_values):
            cell_text = normalize_text(cell_value)
            if not cell_text:
                continue
            abs_row = top_row + r_idx
            abs_col = left_col + c_idx
            if cell_text == "max" and exact_match is None:
                exact_match = (abs_row, abs_col)
            if "max" in cell_text and contains_match is None:
                contains_match = (abs_row, abs_col)

    anchor = exact_match or contains_match
    if anchor is None:
        return None
    return anchor[0], anchor[1], matrix, top_row, left_col


def build_header_lookup(
    matrix: Sequence[Sequence[Any]],
    top_row: int,
    left_col: int,
    header_row: int,
    min_col: int,
    max_col: int,
) -> Dict[int, str]:
    lookup: Dict[int, str] = {}
    for col in range(min_col, max_col + 1):
        text = normalize_text(get_matrix_value(matrix, top_row, left_col, header_row, col))
        if text:
            lookup[col] = text
    return lookup


def resolve_col(
    anchor_col: int,
    header_lookup: Dict[int, str],
    keyword_sets: Sequence[Sequence[str]],
    default_offset: int,
) -> int:
    for col, text in header_lookup.items():
        for keywords in keyword_sets:
            if all(keyword in text for keyword in keywords):
                return col
    return anchor_col + default_offset


def parse_filename_metadata(file_name: str) -> Dict[str, str]:
    stem = Path(file_name).stem
    parts = [part.strip() for part in stem.split("-")]

    ticker = ""
    period_token = ""
    if len(parts) >= 2:
        ticker = parts[1].strip().upper()
    if len(parts) >= 3:
        period_token = parts[2].strip().split("_")[0].strip()

    model_period = ""
    model_date = ""

    match = re.search(r"(Early|Mid|Late)([A-Za-z]{3,9})(\d{4})", period_token, re.IGNORECASE)
    if match:
        period_label = match.group(1).title()
        month_text = match.group(2)
        year = int(match.group(3))

        month_key = month_text.lower()
        if month_key not in MONTH_BY_TEXT:
            month_key = month_key[:3]
        month_num = MONTH_BY_TEXT.get(month_key)
        if month_num:
            month_token = month_text[:3].title()
            model_period = f"{period_label}{month_token}_{year}"
            model_day = DAY_BY_LABEL[period_label.lower()]
            model_date = date(year, month_num, model_day).isoformat()

    model = f"{ticker}_{model_period}" if ticker and model_period else (ticker or stem)

    return {
        "ticker": ticker,
        "model_period": model_period,
        "model_date": model_date,
        "model": model,
    }


def safe_close_source_workbook(workbook: xw.Book) -> None:
    try:
        workbook.close(save=False)
        return
    except TypeError:
        pass
    except Exception:
        pass

    # xlwings variants differ on close() signatures; force no-save at COM layer.
    try:
        workbook.api.Close(SaveChanges=False)
        return
    except Exception:
        pass

    try:
        workbook.close()
    except Exception:
        pass


def get_sheet_case_insensitive(workbook: xw.Book, target_name: str) -> Optional[xw.Sheet]:
    target = target_name.strip().lower()
    for sheet in workbook.sheets:
        if sheet.name.strip().lower() == target:
            return sheet
    return None


def compute_output_path(input_folder: Path, output_folder: Path) -> Path:
    output_folder.mkdir(parents=True, exist_ok=True)
    base_name = f"{input_folder.name}_PARAM"
    candidate = output_folder / f"{base_name}.xlsx"
    suffix = 1
    while candidate.exists():
        candidate = output_folder / f"{base_name}.{suffix}.xlsx"
        suffix += 1
    return candidate


def round_for_signature(value: Optional[float]) -> Optional[float]:
    if value is None:
        return None
    return round(value, 10)


def process_empirical_sheet(
    workbook: xw.Book,
    sheet: xw.Sheet,
    metadata: Dict[str, str],
    source_file: str,
) -> List[Dict[str, Any]]:
    anchor_info = find_anchor_max(sheet)
    if anchor_info is None:
        print(f"Skipped empirical extraction for {source_file} (anchor 'max' not found)")
        return []

    anchor_row, anchor_col, matrix, top_row, left_col = anchor_info
    header_lookup = build_header_lookup(
        matrix=matrix,
        top_row=top_row,
        left_col=left_col,
        header_row=anchor_row,
        min_col=max(1, anchor_col - 24),
        max_col=anchor_col + 24,
    )

    num_quarters_col = resolve_col(anchor_col, header_lookup, [["num", "quarter"], ["quarter", "used"]], -4)
    last_quarter_col = resolve_col(anchor_col, header_lookup, [["last", "quarter"]], -5)
    avg_penetration_col = resolve_col(anchor_col, header_lookup, [["avg", "penetration"], ["average", "penetration"]], -3)
    forecast_col = resolve_col(anchor_col, header_lookup, [["estimated", "total", "sold"], ["forecast"]], -1)
    actual_col = resolve_col(anchor_col, header_lookup, [["reported", "sales"], ["actual"]], -2)
    min_col = resolve_col(anchor_col, header_lookup, [["min"]], 1)
    quarterly_sales_col = resolve_col(anchor_col, header_lookup, [["quarterly", "sales"], ["db", "sales"]], -11)
    reported_sales_col = resolve_col(anchor_col, header_lookup, [["reported", "sales"]], -7)
    growth_rate_col = resolve_col(anchor_col, header_lookup, [["growth"]], -8)
    sales_captured_col = resolve_col(anchor_col, header_lookup, [["captured", "db"], ["sales", "captured"], ["penetration"]], -6)

    quarter_label_col = resolve_col(anchor_col, header_lookup, [["quarter"]], -12)

    history_rows: List[Tuple[int, Any, Optional[float], Optional[float], Optional[float]]] = []
    for row in range(top_row, anchor_row):
        quarterly_sales = to_float(get_matrix_value(matrix, top_row, left_col, row, quarterly_sales_col))
        reported_sales = to_float(get_matrix_value(matrix, top_row, left_col, row, reported_sales_col))
        sales_captured = to_float(get_matrix_value(matrix, top_row, left_col, row, sales_captured_col))
        if (
            sales_captured is None
            and quarterly_sales is not None
            and reported_sales is not None
            and reported_sales != 0
        ):
            sales_captured = quarterly_sales / reported_sales

        if quarterly_sales is None and reported_sales is None and sales_captured is None:
            continue

        quarter_label = get_matrix_value(matrix, top_row, left_col, row, quarter_label_col)
        history_rows.append((row, quarter_label, quarterly_sales, reported_sales, sales_captured))

    helper_row = max(anchor_row + N_QUARTERS + 2, sheet.used_range.last_cell.row + 2)
    helper_col = max(anchor_col + 4, sheet.used_range.last_cell.column + 2)
    avg_formula_cell = sheet.cells(helper_row, helper_col)

    rows: List[Dict[str, Any]] = []
    for idx in range(1, N_QUARTERS + 1):
        data_row = anchor_row + idx

        num_quarters_used = to_int(sheet.cells(data_row, num_quarters_col).value) or idx
        quarter_count = min(max(num_quarters_used, 1), len(history_rows)) if history_rows else 0

        avg_penetration = to_float(sheet.cells(data_row, avg_penetration_col).value)
        if quarter_count > 0:
            start_row = history_rows[-quarter_count][0]
            end_row = history_rows[-1][0]
            avg_formula_cell.formula2 = (
                f"=AVERAGE(R{start_row}C{sales_captured_col}:R{end_row}C{sales_captured_col})"
            )
            workbook.app.calculate()
            avg_penetration = to_float(avg_formula_cell.value) or avg_penetration

        if avg_penetration is None and quarter_count > 0:
            captures = [entry[4] for entry in history_rows[-quarter_count:] if entry[4] is not None]
            avg_penetration = mean(captures) if captures else None

        forecast_value = to_float(sheet.cells(data_row, forecast_col).value)
        actual_value = to_float(sheet.cells(data_row, actual_col).value)
        forecast_max = to_float(sheet.cells(data_row, anchor_col).value)
        forecast_min = to_float(sheet.cells(data_row, min_col).value)
        quarterly_sales = to_float(sheet.cells(data_row, quarterly_sales_col).value)
        reported_sales = to_float(sheet.cells(data_row, reported_sales_col).value)
        growth_rate = to_float(sheet.cells(data_row, growth_rate_col).value)
        sales_captured = to_float(sheet.cells(data_row, sales_captured_col).value)

        if (
            sales_captured is None
            and quarterly_sales is not None
            and reported_sales is not None
            and reported_sales != 0
        ):
            sales_captured = quarterly_sales / reported_sales

        if actual_value is None:
            actual_value = reported_sales

        if (
            forecast_value is None
            and avg_penetration is not None
            and avg_penetration != 0
            and quarterly_sales is not None
        ):
            forecast_value = quarterly_sales / avg_penetration

        if forecast_max is None and forecast_value is not None:
            forecast_max = forecast_value
        if forecast_min is None and forecast_value is not None:
            forecast_min = forecast_value

        range_width = (
            forecast_max - forecast_min
            if forecast_max is not None and forecast_min is not None
            else None
        )

        last_quarter_used = sheet.cells(data_row, last_quarter_col).value
        if is_blank(last_quarter_used) and quarter_count > 0:
            last_quarter_used = history_rows[-1][1]

        has_output_signal = any(
            value is not None
            for value in (
                forecast_value,
                forecast_max,
                forecast_min,
                avg_penetration,
                quarterly_sales,
                reported_sales,
            )
        )
        if not has_output_signal and idx > len(history_rows):
            break

        rows.append(
            {
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
        )

    try:
        avg_formula_cell.value = None
    except Exception:
        pass

    return rows


def process_regression_sheet(
    workbook: xw.Book,
    sheet: xw.Sheet,
    metadata: Dict[str, str],
    source_file: str,
) -> List[Dict[str, Any]]:
    anchor_info = find_anchor_max(sheet)
    if anchor_info is None:
        print(f"Skipped regression extraction for {source_file} (anchor 'max' not found)")
        return []

    anchor_row, anchor_col, matrix, top_row, left_col = anchor_info
    header_lookup = build_header_lookup(
        matrix=matrix,
        top_row=top_row,
        left_col=left_col,
        header_row=anchor_row,
        min_col=max(1, anchor_col - 24),
        max_col=anchor_col + 24,
    )

    x_col = anchor_col - 11
    y_col = anchor_col - 7
    num_quarters_col = resolve_col(anchor_col, header_lookup, [["num", "quarter"], ["quarter", "used"]], -4)
    forecast_col = resolve_col(anchor_col, header_lookup, [["tot", "fcst", "w/o", "sa"], ["forecast"]], -1)
    min_col = resolve_col(anchor_col, header_lookup, [["min"]], 1)
    actual_col = resolve_col(anchor_col, header_lookup, [["actual"]], -2)

    history_xy: List[Tuple[int, float, float]] = []
    for row in range(top_row, anchor_row):
        x_value = to_float(get_matrix_value(matrix, top_row, left_col, row, x_col))
        y_value = to_float(get_matrix_value(matrix, top_row, left_col, row, y_col))
        if x_value is None or y_value is None:
            continue
        history_xy.append((row, x_value, y_value))

    if not history_xy:
        print(f"Skipped regression extraction for {source_file} (no numeric x/y history)")
        return []

    helper_row = max(anchor_row + N_QUARTERS + 2, sheet.used_range.last_cell.row + 2)
    helper_col = max(anchor_col + 4, sheet.used_range.last_cell.column + 2)
    intercept_cell = sheet.cells(helper_row, helper_col)
    slope_cell = sheet.cells(helper_row, helper_col + 1)

    rows: List[Dict[str, Any]] = []
    previous_signature: Optional[Tuple[Any, ...]] = None

    for idx in range(1, N_QUARTERS + 1):
        data_row = anchor_row + idx
        num_quarters_used = to_int(sheet.cells(data_row, num_quarters_col).value) or idx
        quarter_count = min(max(num_quarters_used, 1), len(history_xy))

        sample = history_xy[-quarter_count:]
        start_row = sample[0][0]
        end_row = sample[-1][0]

        intercept_cell.formula2 = (
            f"=INTERCEPT(R{start_row}C{y_col}:R{end_row}C{y_col},R{start_row}C{x_col}:R{end_row}C{x_col})"
        )
        slope_cell.formula2 = (
            f"=SLOPE(R{start_row}C{y_col}:R{end_row}C{y_col},R{start_row}C{x_col}:R{end_row}C{x_col})"
        )
        workbook.app.calculate()

        intercept = to_float(intercept_cell.value)
        slope = to_float(slope_cell.value)

        forecast_value = to_float(sheet.cells(data_row, forecast_col).value)
        if forecast_value is None and intercept is not None and slope is not None:
            forecast_x = to_float(sheet.cells(end_row + 1, x_col).value)
            if forecast_x is None:
                forecast_x = to_float(sheet.cells(data_row, x_col).value)
            if forecast_x is None:
                forecast_x = sample[-1][1]
            forecast_value = intercept + (slope * forecast_x) if forecast_x is not None else None

        actual_value = to_float(sheet.cells(data_row, actual_col).value)
        forecast_max = to_float(sheet.cells(data_row, anchor_col).value)
        forecast_min = to_float(sheet.cells(data_row, min_col).value)

        if forecast_max is None and forecast_value is not None:
            forecast_max = forecast_value
        if forecast_min is None and forecast_value is not None:
            forecast_min = forecast_value

        range_width = (
            forecast_max - forecast_min
            if forecast_max is not None and forecast_min is not None
            else None
        )

        signature = (
            num_quarters_used,
            round_for_signature(intercept),
            round_for_signature(slope),
            round_for_signature(forecast_value),
            round_for_signature(forecast_max),
            round_for_signature(forecast_min),
        )

        if signature == previous_signature:
            continue
        previous_signature = signature

        has_output_signal = any(
            value is not None
            for value in (
                forecast_value,
                intercept,
                slope,
                forecast_max,
                forecast_min,
            )
        )
        if not has_output_signal and idx > len(history_xy):
            break

        rows.append(
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
                "range_width": range_width,
                "intercept": intercept,
                "slope": slope,
                "source_file": source_file,
            }
        )

    try:
        intercept_cell.value = None
        slope_cell.value = None
    except Exception:
        pass

    return rows


def autosize_columns(sheet) -> None:
    for col_idx in range(1, sheet.max_column + 1):
        max_len = 0
        for row_idx in range(1, sheet.max_row + 1):
            value = sheet.cell(row=row_idx, column=col_idx).value
            if value is None:
                continue
            value_len = len(str(value))
            if value_len > max_len:
                max_len = value_len
        adjusted = min(50, max(12, max_len + 2))
        sheet.column_dimensions[get_column_letter(col_idx)].width = adjusted


def write_output_workbook(
    path: Path,
    empirical_rows: Sequence[Dict[str, Any]],
    regression_rows: Sequence[Dict[str, Any]],
) -> None:
    wb = Workbook()
    default_sheet = wb.active
    wb.remove(default_sheet)

    empirical_ws = wb.create_sheet("empirical_candidates")
    regression_ws = wb.create_sheet("regression_candidates")

    empirical_ws.append(EMPIRICAL_COLUMNS)
    for row in empirical_rows:
        empirical_ws.append([row.get(col, "") for col in EMPIRICAL_COLUMNS])

    regression_ws.append(REGRESSION_COLUMNS)
    for row in regression_rows:
        regression_ws.append([row.get(col, "") for col in REGRESSION_COLUMNS])

    for ws in (empirical_ws, regression_ws):
        for header_cell in ws[1]:
            header_cell.font = Font(bold=True)
        ws.freeze_panes = "A2"
        last_col = get_column_letter(ws.max_column)
        ws.auto_filter.ref = f"A1:{last_col}{max(ws.max_row, 1)}"
        autosize_columns(ws)

    wb.save(path)


def list_input_workbooks(folder: Path) -> List[Path]:
    files: List[Path] = []

    if not folder.exists():
        print(f"Skipped input directory: {folder} (does not exist)")
        return files

    for file_path in sorted(folder.iterdir()):
        if not file_path.is_file():
            continue
        if file_path.name.startswith("~"):
            print(f"Skipped file: {file_path.name} (temporary file)")
            continue
        if file_path.suffix.lower() != ".xlsx":
            print(f"Skipped file: {file_path.name} (not an .xlsx file)")
            continue
        files.append(file_path)

    return files


def main() -> None:
    files_to_process = list_input_workbooks(input_dir)
    output_path = compute_output_path(input_dir, output_dir)

    empirical_rows: List[Dict[str, Any]] = []
    regression_rows: List[Dict[str, Any]] = []
    processed_files = 0

    app: Optional[xw.App] = None
    try:
        app = xw.App(visible=False, add_book=False)
        app.display_alerts = False
        app.screen_updating = False
        try:
            app.calculation = "manual"
        except Exception:
            pass

        for file_path in files_to_process:
            print(f"Processing file: {file_path.name}")
            workbook: Optional[xw.Book] = None
            try:
                workbook = app.books.open(str(file_path), update_links=False)
                metadata = parse_filename_metadata(file_path.name)

                empirical_sheet = get_sheet_case_insensitive(workbook, "Empirical Model")
                if empirical_sheet is None:
                    print(f"Skipped empirical extraction for {file_path.name} (sheet not found)")
                else:
                    empirical_rows.extend(
                        process_empirical_sheet(
                            workbook=workbook,
                            sheet=empirical_sheet,
                            metadata=metadata,
                            source_file=file_path.name,
                        )
                    )

                regression_sheet = get_sheet_case_insensitive(workbook, "Regression Model")
                if regression_sheet is None:
                    print(f"Skipped regression extraction for {file_path.name} (sheet not found)")
                else:
                    regression_rows.extend(
                        process_regression_sheet(
                            workbook=workbook,
                            sheet=regression_sheet,
                            metadata=metadata,
                            source_file=file_path.name,
                        )
                    )

                processed_files += 1
            except Exception as exc:
                print(f"Skipped file: {file_path.name} (processing error: {exc})")
            finally:
                if workbook is not None:
                    safe_close_source_workbook(workbook)
    finally:
        if app is not None:
            try:
                app.quit()
            except Exception:
                pass

    write_output_workbook(
        path=output_path,
        empirical_rows=empirical_rows,
        regression_rows=regression_rows,
    )

    print(f"Output path: {output_path}")
    print(f"Number of files processed: {processed_files}")
    print(f"Number of empirical rows: {len(empirical_rows)}")
    print(f"Number of regression rows: {len(regression_rows)}")


if __name__ == "__main__":
    main()
