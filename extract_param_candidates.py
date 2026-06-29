from __future__ import annotations

import re
from datetime import date
from pathlib import Path
from typing import Any, Dict, List, Optional, Sequence, Tuple

import xlwings as xw


# --- User-configurable paths ---
input_dir = r"/workspace/input"
output_dir = r"/workspace/output"


# --- Static configuration ---
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


# Offsets are relative to the column containing the "max" anchor cell.
EMPIRICAL_OFFSETS = {
    "data_start_row": 1,
    "last_quarter_used_col": -12,
    "estimated_total_sold_col": -9,
    "reported_sales_col": -8,
    "quarterly_sales_col": -7,
    "growth_rate_pct_col": -6,
    "sales_captured_in_db_pct_col": -5,
    "penetration_pct_col": -4,
    "forecast_max_col": 0,
    "forecast_min_col": 1,
    "avg_penetration_scratch_col": 18,
}

REGRESSION_OFFSETS = {
    "data_start_row": 1,
    "num_quarters_used_col": -12,
    "forecast_total_without_sa_col": -1,
    "actual_value_col": -2,
    "forecast_max_col": 0,
    "forecast_min_col": 1,
    "intercept_scratch_col": 18,
    "slope_scratch_col": 19,
}


PERIOD_PATTERN = re.compile(r"^(Early|Mid|Late)([A-Za-z]+)(\d{4})$")
MONTH_MAP = {
    "Jan": 1,
    "Feb": 2,
    "Mar": 3,
    "Apr": 4,
    "May": 5,
    "Jun": 6,
    "Jul": 7,
    "Aug": 8,
    "Sep": 9,
    "Oct": 10,
    "Nov": 11,
    "Dec": 12,
}
DAY_MAP = {"Early": 5, "Mid": 15, "Late": 25}


def to_float(value: Any) -> Optional[float]:
    if value is None or value == "":
        return None
    if isinstance(value, (int, float)):
        return float(value)
    if isinstance(value, str):
        cleaned = value.strip().replace(",", "")
        if cleaned.endswith("%"):
            cleaned = cleaned[:-1]
        try:
            return float(cleaned)
        except ValueError:
            return None
    return None


def to_text(value: Any) -> str:
    if value is None:
        return ""
    text = str(value).strip()
    return text


def is_blank(value: Any) -> bool:
    if value is None:
        return True
    if isinstance(value, str) and value.strip() == "":
        return True
    return False


def normalize_2d(values: Any) -> List[List[Any]]:
    if values is None:
        return []
    if isinstance(values, list):
        if not values:
            return []
        if isinstance(values[0], list):
            return values
        return [values]
    return [[values]]


def get_sheet(wb: xw.Book, sheet_name: str) -> Optional[xw.Sheet]:
    for sh in wb.sheets:
        if sh.name == sheet_name:
            return sh
    return None


def find_anchor_cell(sheet: xw.Sheet, anchor_label: str = "max") -> Optional[Tuple[int, int]]:
    used = sheet.used_range
    matrix = normalize_2d(used.value)
    if not matrix:
        return None

    start_row = used.row
    start_col = used.column
    target = anchor_label.lower()

    for row_idx, row_values in enumerate(matrix):
        for col_idx, value in enumerate(row_values):
            if isinstance(value, str) and value.strip().lower() == target:
                return (start_row + row_idx, start_col + col_idx)
    return None


def parse_file_labels(file_path: Path) -> Dict[str, str]:
    parts = [part.strip() for part in file_path.stem.split(" - ")]
    if len(parts) < 3:
        raise ValueError("filename does not match expected pattern '<prefix> - <ticker> - <period>_Send'")

    ticker = parts[1]
    period_token = parts[2].split("_")[0]
    match = PERIOD_PATTERN.match(period_token)
    if not match:
        raise ValueError("filename period does not match Early/Mid/Late + Month + Year pattern")

    period_prefix = match.group(1)
    month_token = match.group(2)[:3].title()
    year = int(match.group(3))

    month_num = MONTH_MAP.get(month_token)
    if month_num is None:
        raise ValueError(f"unknown month token '{month_token}' in filename")

    model_period = f"{period_prefix}{month_token}_{year}"
    model_date = date(year, month_num, DAY_MAP[period_prefix]).isoformat()
    model = f"{ticker}_{model_period}"

    return {
        "ticker": ticker,
        "model_period": model_period,
        "model_date": model_date,
        "model": model,
    }


def close_source_workbook_safely(wb: xw.Book) -> None:
    try:
        wb.close(save=False)
        return
    except TypeError:
        pass
    except Exception:
        pass

    close_attempts = [
        lambda: wb.api.Close(False),
        lambda: wb.api.Close(SaveChanges=False),
        lambda: wb.close(),
    ]
    for close_call in close_attempts:
        try:
            close_call()
            return
        except Exception:
            continue

    print(f"WARNING: could not safely close workbook: {wb.fullname}")


def get_output_path(input_path: Path, output_path: Path) -> Path:
    base_name = f"{input_path.name}_PARAM"
    candidate = output_path / f"{base_name}.xlsx"
    suffix = 1
    while candidate.exists():
        candidate = output_path / f"{base_name}.{suffix}.xlsx"
        suffix += 1
    return candidate


def restore_cell_state(cell: xw.Range, original_formula2: Any, original_value: Any) -> None:
    try:
        if original_formula2 not in (None, ""):
            cell.formula2 = original_formula2
            return
        if is_blank(original_value):
            cell.clear_contents()
            return
        cell.value = original_value
    except Exception:
        # Source workbook is closed without saving, so this is harmless.
        pass


def calc_range_formula(cell: xw.Range, formula2: str, wb: xw.Book) -> Any:
    cell.formula2 = formula2
    wb.app.calculate()
    return cell.value


def range_width(max_value: Optional[float], min_value: Optional[float]) -> Optional[float]:
    if max_value is None or min_value is None:
        return None
    return max_value - min_value


def extract_empirical_rows(
    wb: xw.Book,
    labels: Dict[str, str],
    source_file: str,
) -> List[Dict[str, Any]]:
    sheet = get_sheet(wb, EMPIRICAL_MODEL_SHEET)
    if sheet is None:
        print(f"Skipped empirical extraction for {source_file}: missing sheet '{EMPIRICAL_MODEL_SHEET}'")
        return []

    anchor = find_anchor_cell(sheet, "max")
    if anchor is None:
        print(f"Skipped empirical extraction for {source_file}: could not find 'max' anchor")
        return []

    anchor_row, anchor_col = anchor
    data_start_row = anchor_row + EMPIRICAL_OFFSETS["data_start_row"]
    penetration_col = anchor_col + EMPIRICAL_OFFSETS["penetration_pct_col"]

    scratch_cell = sheet.range((anchor_row, anchor_col + EMPIRICAL_OFFSETS["avg_penetration_scratch_col"]))
    original_formula2 = scratch_cell.formula2
    original_value = scratch_cell.value

    rows: List[Dict[str, Any]] = []
    try:
        for i in range(N_QUARTERS):
            n_quarters = i + 1
            row = data_start_row + i
            start_row = max(data_start_row, row - n_quarters + 1)

            avg_formula = f"=AVERAGE(R{start_row}C{penetration_col}:R{row}C{penetration_col})"
            avg_penetration_pct = to_float(calc_range_formula(scratch_cell, avg_formula, wb))

            last_quarter_used = sheet.range(
                (row, anchor_col + EMPIRICAL_OFFSETS["last_quarter_used_col"])
            ).value
            estimated_total_sold = to_float(
                sheet.range((row, anchor_col + EMPIRICAL_OFFSETS["estimated_total_sold_col"])).value
            )
            reported_sales = to_float(
                sheet.range((row, anchor_col + EMPIRICAL_OFFSETS["reported_sales_col"])).value
            )
            quarterly_sales = to_float(
                sheet.range((row, anchor_col + EMPIRICAL_OFFSETS["quarterly_sales_col"])).value
            )
            growth_rate_pct = to_float(
                sheet.range((row, anchor_col + EMPIRICAL_OFFSETS["growth_rate_pct_col"])).value
            )
            sales_captured_in_db_pct = to_float(
                sheet.range((row, anchor_col + EMPIRICAL_OFFSETS["sales_captured_in_db_pct_col"])).value
            )
            forecast_max = to_float(sheet.range((row, anchor_col + EMPIRICAL_OFFSETS["forecast_max_col"])).value)
            forecast_min = to_float(sheet.range((row, anchor_col + EMPIRICAL_OFFSETS["forecast_min_col"])).value)

            if estimated_total_sold is None and avg_penetration_pct is not None and quarterly_sales is not None:
                estimated_total_sold = avg_penetration_pct * quarterly_sales

            if all(
                is_blank(v)
                for v in (
                    last_quarter_used,
                    estimated_total_sold,
                    reported_sales,
                    quarterly_sales,
                    forecast_max,
                    forecast_min,
                    avg_penetration_pct,
                )
            ):
                continue

            rows.append(
                {
                    "model": labels["model"],
                    "ticker": labels["ticker"],
                    "model_period": labels["model_period"],
                    "model_date": labels["model_date"],
                    "method": "empirical",
                    "parameter_name": "avg_penetration_pct",
                    "parameter_value": avg_penetration_pct,
                    "num_quarters_used": n_quarters,
                    "last_quarter_used": to_text(last_quarter_used),
                    "forecast_value": estimated_total_sold,
                    "actual_value": reported_sales,
                    "forecast_max": forecast_max,
                    "forecast_min": forecast_min,
                    "range_width": range_width(forecast_max, forecast_min),
                    "avg_penetration_pct": avg_penetration_pct,
                    "quarterly_sales": quarterly_sales,
                    "reported_sales": reported_sales,
                    "growth_rate_pct": growth_rate_pct,
                    "sales_captured_in_db_pct": sales_captured_in_db_pct,
                    "source_file": source_file,
                }
            )
    finally:
        restore_cell_state(scratch_cell, original_formula2, original_value)

    return rows


def rounded_signature(values: Sequence[Any]) -> Tuple[Any, ...]:
    out: List[Any] = []
    for value in values:
        if isinstance(value, float):
            out.append(round(value, 8))
        else:
            out.append(value)
    return tuple(out)


def extract_regression_rows(
    wb: xw.Book,
    labels: Dict[str, str],
    source_file: str,
) -> List[Dict[str, Any]]:
    sheet = get_sheet(wb, REGRESSION_MODEL_SHEET)
    if sheet is None:
        print(f"Skipped regression extraction for {source_file}: missing sheet '{REGRESSION_MODEL_SHEET}'")
        return []

    anchor = find_anchor_cell(sheet, "max")
    if anchor is None:
        print(f"Skipped regression extraction for {source_file}: could not find 'max' anchor")
        return []

    anchor_row, anchor_col = anchor
    data_start_row = anchor_row + REGRESSION_OFFSETS["data_start_row"]

    y_col = anchor_col - 7
    x_col = anchor_col - 11

    intercept_cell = sheet.range((anchor_row, anchor_col + REGRESSION_OFFSETS["intercept_scratch_col"]))
    slope_cell = sheet.range((anchor_row, anchor_col + REGRESSION_OFFSETS["slope_scratch_col"]))

    intercept_original_formula2 = intercept_cell.formula2
    intercept_original_value = intercept_cell.value
    slope_original_formula2 = slope_cell.formula2
    slope_original_value = slope_cell.value

    rows: List[Dict[str, Any]] = []
    previous_signature: Optional[Tuple[Any, ...]] = None

    try:
        for i in range(N_QUARTERS):
            n_quarters = i + 1
            row = data_start_row + i
            start_row = max(data_start_row, row - n_quarters + 1)

            intercept_formula = (
                f"=INTERCEPT(R{start_row}C{y_col}:R{row}C{y_col},R{start_row}C{x_col}:R{row}C{x_col})"
            )
            slope_formula = f"=SLOPE(R{start_row}C{y_col}:R{row}C{y_col},R{start_row}C{x_col}:R{row}C{x_col})"

            intercept_cell.formula2 = intercept_formula
            slope_cell.formula2 = slope_formula
            wb.app.calculate()

            intercept = to_float(intercept_cell.value)
            slope = to_float(slope_cell.value)

            num_quarters_used = to_float(
                sheet.range((row, anchor_col + REGRESSION_OFFSETS["num_quarters_used_col"])).value
            )
            if num_quarters_used is None:
                num_quarters_used = float(n_quarters)

            forecast_total_without_sa = to_float(
                sheet.range((row, anchor_col + REGRESSION_OFFSETS["forecast_total_without_sa_col"])).value
            )
            actual_value = to_float(sheet.range((row, anchor_col + REGRESSION_OFFSETS["actual_value_col"])).value)
            forecast_max = to_float(sheet.range((row, anchor_col + REGRESSION_OFFSETS["forecast_max_col"])).value)
            forecast_min = to_float(sheet.range((row, anchor_col + REGRESSION_OFFSETS["forecast_min_col"])).value)

            if all(
                is_blank(v) for v in (forecast_total_without_sa, forecast_max, forecast_min, intercept, slope)
            ):
                continue

            current_signature = rounded_signature(
                (
                    num_quarters_used,
                    intercept,
                    slope,
                    forecast_total_without_sa,
                    forecast_max,
                    forecast_min,
                )
            )
            if previous_signature is not None and current_signature == previous_signature:
                continue
            previous_signature = current_signature

            rows.append(
                {
                    "model": labels["model"],
                    "ticker": labels["ticker"],
                    "model_period": labels["model_period"],
                    "model_date": labels["model_date"],
                    "method": "regression",
                    "parameter_name": "num_quarters_used",
                    "parameter_value": num_quarters_used,
                    "num_quarters_used": num_quarters_used,
                    "forecast_value": forecast_total_without_sa,
                    "actual_value": actual_value if actual_value is not None else "",
                    "forecast_max": forecast_max,
                    "forecast_min": forecast_min,
                    "range_width": range_width(forecast_max, forecast_min),
                    "intercept": intercept,
                    "slope": slope,
                    "source_file": source_file,
                }
            )
    finally:
        restore_cell_state(intercept_cell, intercept_original_formula2, intercept_original_value)
        restore_cell_state(slope_cell, slope_original_formula2, slope_original_value)

    return rows


def rows_to_matrix(rows: List[Dict[str, Any]], columns: List[str]) -> List[List[Any]]:
    return [[row.get(col, "") for col in columns] for row in rows]


def auto_column_widths(matrix: List[List[Any]], headers: List[str]) -> List[int]:
    widths: List[int] = []
    for col_idx, header in enumerate(headers):
        max_len = len(header)
        for row in matrix:
            value = row[col_idx]
            if value is None:
                continue
            max_len = max(max_len, len(str(value)))
        widths.append(min(max(max_len + 2, 12), 48))
    return widths


def write_output_sheet(sheet: xw.Sheet, headers: List[str], rows: List[Dict[str, Any]]) -> None:
    matrix = rows_to_matrix(rows, headers)
    last_row = 1 + len(matrix)
    last_col = len(headers)

    sheet.range((1, 1)).value = headers
    if matrix:
        sheet.range((2, 1)).value = matrix

    sheet.range((1, 1), (1, last_col)).api.Font.Bold = True
    sheet.range((1, 1), (max(last_row, 1), last_col)).api.AutoFilter()

    widths = auto_column_widths(matrix, headers)
    for col_idx, width in enumerate(widths, start=1):
        sheet.range((1, col_idx)).column_width = width

    try:
        sheet.activate()
        sheet.range("A2").select()
        sheet.book.app.api.ActiveWindow.FreezePanes = False
        sheet.book.app.api.ActiveWindow.FreezePanes = True
    except Exception:
        # Freeze panes may fail in certain headless Excel configurations.
        pass


def write_output_workbook(
    app: xw.App,
    output_path: Path,
    empirical_rows: List[Dict[str, Any]],
    regression_rows: List[Dict[str, Any]],
) -> None:
    out_wb = app.books.add()
    try:
        empirical_sheet = out_wb.sheets[0]
        empirical_sheet.name = EMPIRICAL_OUTPUT_SHEET
        write_output_sheet(empirical_sheet, EMPIRICAL_COLUMNS, empirical_rows)

        regression_sheet = out_wb.sheets.add(REGRESSION_OUTPUT_SHEET, after=empirical_sheet)
        write_output_sheet(regression_sheet, REGRESSION_COLUMNS, regression_rows)

        out_wb.save(str(output_path))
    finally:
        out_wb.close()


def process_all_workbooks() -> None:
    input_path = Path(input_dir).expanduser().resolve()
    output_path = Path(output_dir).expanduser().resolve()

    if not input_path.exists() or not input_path.is_dir():
        raise FileNotFoundError(f"input_dir does not exist or is not a folder: {input_path}")

    output_path.mkdir(parents=True, exist_ok=True)

    source_paths = sorted(input_path.iterdir(), key=lambda p: p.name.lower())

    empirical_rows: List[Dict[str, Any]] = []
    regression_rows: List[Dict[str, Any]] = []
    processed_count = 0

    app = xw.App(visible=False, add_book=False)
    app.display_alerts = False
    app.screen_updating = False

    try:
        for file_path in source_paths:
            if file_path.name.startswith("~"):
                print(f"Skipped {file_path.name}: temporary Excel file")
                continue
            if not file_path.is_file():
                print(f"Skipped {file_path.name}: not a file")
                continue
            if file_path.suffix.lower() != ".xlsx":
                print(f"Skipped {file_path.name}: not an .xlsx file")
                continue

            try:
                labels = parse_file_labels(file_path)
            except ValueError as exc:
                print(f"Skipped {file_path.name}: {exc}")
                continue

            wb: Optional[xw.Book] = None
            try:
                wb = app.books.open(str(file_path), update_links=False)
                file_empirical_rows = extract_empirical_rows(wb, labels, file_path.name)
                file_regression_rows = extract_regression_rows(wb, labels, file_path.name)
                empirical_rows.extend(file_empirical_rows)
                regression_rows.extend(file_regression_rows)
                processed_count += 1
                print(f"Processed {file_path.name}")
            except Exception as exc:
                print(f"Skipped {file_path.name}: processing failure ({exc})")
            finally:
                if wb is not None:
                    close_source_workbook_safely(wb)

        final_output_path = get_output_path(input_path, output_path)
        write_output_workbook(app, final_output_path, empirical_rows, regression_rows)

        print(f"Output path: {final_output_path}")
        print(f"Number of files processed: {processed_count}")
        print(f"Number of empirical rows: {len(empirical_rows)}")
        print(f"Number of regression rows: {len(regression_rows)}")
    finally:
        app.quit()


if __name__ == "__main__":
    process_all_workbooks()
