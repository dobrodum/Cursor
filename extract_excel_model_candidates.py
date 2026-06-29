from __future__ import annotations

import datetime as dt
import re
from pathlib import Path
from typing import Any, Dict, List, Optional, Sequence, Tuple

import xlwings as xw
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

# -----------------------------
# User-configurable directories
# -----------------------------
input_dir = "/path/to/input"
output_dir = "/path/to/output"


EMPIRICAL_COLUMNS: List[str] = [
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

REGRESSION_COLUMNS: List[str] = [
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


MONTH_MAP: Dict[str, int] = {
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

PHASE_TO_DAY: Dict[str, int] = {"Early": 5, "Mid": 15, "Late": 25}


def normalize_header(value: Any) -> str:
    if value is None:
        return ""
    text = str(value).strip().lower()
    if not text:
        return ""
    text = text.replace("w/o", "without")
    text = text.replace("%", " pct ")
    text = re.sub(r"[^a-z0-9]+", "_", text)
    text = re.sub(r"_+", "_", text).strip("_")
    return text


def is_blank(value: Any) -> bool:
    if value is None:
        return True
    if isinstance(value, str) and not value.strip():
        return True
    return False


def to_float(value: Any) -> Optional[float]:
    if is_blank(value):
        return None
    try:
        return float(value)
    except (TypeError, ValueError):
        return None


def numeric_diff(a: Any, b: Any) -> Optional[float]:
    left = to_float(a)
    right = to_float(b)
    if left is None or right is None:
        return None
    return left - right


def read_matrix(rng: Any) -> List[List[Any]]:
    values = rng.options(ndim=2).value
    if values is None:
        return []
    return [list(row) for row in values]


def get_sheet_case_insensitive(workbook: xw.Book, name: str) -> Optional[xw.Sheet]:
    target = name.strip().lower()
    for sheet in workbook.sheets:
        if sheet.name.strip().lower() == target:
            return sheet
    return None


def find_max_anchor(sheet: xw.Sheet) -> Optional[Tuple[int, int]]:
    used = sheet.used_range
    matrix = read_matrix(used)
    if not matrix:
        return None

    base_row = used.row
    base_col = used.column
    for row_idx, row_values in enumerate(matrix):
        for col_idx, value in enumerate(row_values):
            if isinstance(value, str) and value.strip().lower() == "max":
                return base_row + row_idx, base_col + col_idx
    return None


def build_header_lookup(
    sheet: xw.Sheet,
    header_row: int,
    anchor_col: int,
    scan_left: int = 25,
    scan_right: int = 25,
) -> Dict[str, int]:
    start_col = max(1, anchor_col - scan_left)
    end_col = anchor_col + scan_right
    row_values = read_matrix(sheet.range((header_row, start_col), (header_row, end_col)))[0]
    lookup: Dict[str, int] = {}
    for idx, value in enumerate(row_values):
        key = normalize_header(value)
        if key and key not in lookup:
            lookup[key] = start_col + idx
    return lookup


def find_column(
    headers: Dict[str, int],
    aliases: Sequence[str],
    fallback: Optional[int],
) -> Optional[int]:
    alias_keys = [normalize_header(alias) for alias in aliases]

    for key in alias_keys:
        if key and key in headers:
            return headers[key]

    for header_key, column in headers.items():
        for alias_key in alias_keys:
            if alias_key and (alias_key in header_key or header_key in alias_key):
                return column

    return fallback


def parse_file_labels(file_name: str) -> Dict[str, str]:
    stem = Path(file_name).stem
    parts = [part.strip() for part in stem.split(" - ")]

    ticker = "UNKNOWN"
    if len(parts) >= 2 and parts[1]:
        cleaned = re.sub(r"[^A-Za-z0-9]", "", parts[1]).upper()
        if cleaned:
            ticker = cleaned

    if ticker == "UNKNOWN":
        ticker_match = re.search(r"\b([A-Z]{2,8})\b", stem)
        if ticker_match:
            ticker = ticker_match.group(1)

    period_match = re.search(
        r"(Early|Mid|Late)\s*([A-Za-z]{3,9})\s*(\d{4})",
        stem,
        flags=re.IGNORECASE,
    )
    if not period_match:
        period_match = re.search(
            r"(Early|Mid|Late)_?([A-Za-z]{3,9})_?(\d{4})",
            stem,
            flags=re.IGNORECASE,
        )

    model_period = "UnknownPeriod"
    model_date = ""
    if period_match:
        phase = period_match.group(1).title()
        month_token = period_match.group(2).title()
        year = int(period_match.group(3))
        month_number = MONTH_MAP.get(month_token[:3].lower())
        if month_number:
            month_abbr = dt.date(year, month_number, 1).strftime("%b")
            model_period = f"{phase}{month_abbr}_{year}"
            day = PHASE_TO_DAY[phase]
            model_date = dt.date(year, month_number, day).isoformat()

    model = f"{ticker}_{model_period}"

    return {
        "model": model,
        "ticker": ticker,
        "model_period": model_period,
        "model_date": model_date,
    }


def build_output_path(input_path: Path, output_path: Path) -> Path:
    base_name = f"{input_path.name}_PARAM"
    candidate = output_path / f"{base_name}.xlsx"
    suffix = 1
    while candidate.exists():
        candidate = output_path / f"{base_name}.{suffix}.xlsx"
        suffix += 1
    return candidate


def close_workbook_no_save(workbook: xw.Book) -> None:
    close_attempts = [
        lambda: workbook.close(save=False),
        lambda: workbook.close(False),
        lambda: workbook.api.Close(SaveChanges=False),  # type: ignore[attr-defined]
    ]
    for close_action in close_attempts:
        try:
            close_action()
            return
        except Exception:
            continue

    try:
        workbook.close()
    except Exception:
        pass


def set_formula2_or_formula(target: Any, formula: str) -> None:
    try:
        target.formula2 = formula
    except Exception:
        target.formula = formula


def matrix_cell(matrix: List[List[Any]], row_idx: int, col: Optional[int], min_col: int) -> Any:
    if col is None:
        return None
    if row_idx < 0 or row_idx >= len(matrix):
        return None
    col_idx = col - min_col
    if col_idx < 0 or col_idx >= len(matrix[row_idx]):
        return None
    return matrix[row_idx][col_idx]


def extract_empirical_rows(
    workbook: xw.Book,
    labels: Dict[str, str],
    source_file: str,
) -> List[Dict[str, Any]]:
    sheet = get_sheet_case_insensitive(workbook, "Empirical Model")
    if sheet is None:
        print(f"  Empirical Model missing in {source_file}; skipped empirical extraction.")
        return []

    anchor = find_max_anchor(sheet)
    if anchor is None:
        print(f"  Empirical max anchor missing in {source_file}; skipped empirical extraction.")
        return []

    anchor_row, anchor_col = anchor
    headers = build_header_lookup(sheet, anchor_row, anchor_col)
    n_quarters = 10
    start_row = anchor_row + 1
    end_row = start_row + n_quarters - 1

    col_map = {
        "num_quarters_used": find_column(
            headers,
            ["num_quarters_used", "num_quarters", "quarters_used", "n_quarters"],
            anchor_col - 10,
        ),
        "last_quarter_used": find_column(
            headers,
            ["last_quarter_used", "last_quarter", "last_qtr_used"],
            anchor_col - 9,
        ),
        "forecast_value": find_column(
            headers,
            ["estimated_total_sold", "forecast_value", "forecast", "tot_fcst"],
            anchor_col - 4,
        ),
        "actual_value": find_column(
            headers,
            ["actual_value", "reported_sales", "actual", "reported"],
            anchor_col - 3,
        ),
        "forecast_max": anchor_col,
        "forecast_min": find_column(headers, ["forecast_min", "min"], anchor_col + 1),
        "avg_penetration_pct": find_column(
            headers,
            ["avg_penetration_pct", "avg_penetration", "average_penetration"],
            anchor_col - 6,
        ),
        "quarterly_sales": find_column(
            headers,
            ["quarterly_sales", "sales_quarterly", "quarter_sales"],
            anchor_col - 8,
        ),
        "reported_sales": find_column(
            headers,
            ["reported_sales", "reported"],
            anchor_col - 7,
        ),
        "growth_rate_pct": find_column(
            headers,
            ["growth_rate_pct", "growth_rate", "growth_pct"],
            anchor_col - 5,
        ),
        "sales_captured_in_db_pct": find_column(
            headers,
            ["sales_captured_in_db_pct", "sales_captured_pct", "captured_in_db_pct"],
            anchor_col - 2,
        ),
    }

    required_columns = [col for col in col_map.values() if col is not None]
    if not required_columns:
        return []

    min_col = min(required_columns)
    max_col = max(required_columns)
    data_matrix = read_matrix(sheet.range((start_row, min_col), (end_row, max_col)))

    # Use R1C1 + formula2 to compute avg penetration values in one pass.
    avg_values: List[Any] = [None] * n_quarters
    source_avg_col = col_map.get("avg_penetration_pct") or col_map.get("sales_captured_in_db_pct")
    if source_avg_col is not None:
        temp_col = sheet.used_range.last_cell.column + 1
        target_range = sheet.range((start_row, temp_col), (end_row, temp_col))
        offset = source_avg_col - temp_col
        avg_formula = f'=IFERROR(AVERAGE(RC[{offset}]:RC[{offset}]),"")'
        set_formula2_or_formula(target_range, avg_formula)
        workbook.app.calculate()
        avg_values = [row[0] for row in read_matrix(target_range)]

    rows: List[Dict[str, Any]] = []
    for idx in range(n_quarters):
        num_quarters_value = matrix_cell(data_matrix, idx, col_map["num_quarters_used"], min_col)
        num_quarters_used = num_quarters_value if not is_blank(num_quarters_value) else idx + 1

        last_quarter_used = matrix_cell(data_matrix, idx, col_map["last_quarter_used"], min_col)
        forecast_value = matrix_cell(data_matrix, idx, col_map["forecast_value"], min_col)
        actual_value = matrix_cell(data_matrix, idx, col_map["actual_value"], min_col)
        forecast_max = matrix_cell(data_matrix, idx, col_map["forecast_max"], min_col)
        forecast_min = matrix_cell(data_matrix, idx, col_map["forecast_min"], min_col)
        avg_penetration_pct = avg_values[idx]
        if is_blank(avg_penetration_pct):
            avg_penetration_pct = matrix_cell(data_matrix, idx, col_map["avg_penetration_pct"], min_col)

        quarterly_sales = matrix_cell(data_matrix, idx, col_map["quarterly_sales"], min_col)
        reported_sales = matrix_cell(data_matrix, idx, col_map["reported_sales"], min_col)
        growth_rate_pct = matrix_cell(data_matrix, idx, col_map["growth_rate_pct"], min_col)
        sales_captured_in_db_pct = matrix_cell(
            data_matrix,
            idx,
            col_map["sales_captured_in_db_pct"],
            min_col,
        )
        range_width = numeric_diff(forecast_max, forecast_min)

        if all(
            is_blank(value)
            for value in (
                forecast_value,
                actual_value,
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
                "model": labels["model"],
                "ticker": labels["ticker"],
                "model_period": labels["model_period"],
                "model_date": labels["model_date"],
                "method": "empirical",
                "parameter_name": "avg_penetration_pct",
                "parameter_value": avg_penetration_pct,
                "num_quarters_used": num_quarters_used,
                "last_quarter_used": last_quarter_used,
                "forecast_value": forecast_value,
                "actual_value": actual_value,
                "forecast_max": forecast_max,
                "forecast_min": forecast_min,
                "range_width": range_width,
                "avg_penetration_pct": avg_penetration_pct,
                "quarterly_sales": quarterly_sales,
                "reported_sales": reported_sales,
                "growth_rate_pct": growth_rate_pct,
                "sales_captured_in_db_pct": sales_captured_in_db_pct,
                "source_file": source_file,
            }
        )

    return rows


def signature_value(value: Any) -> Any:
    numeric_value = to_float(value)
    if numeric_value is not None:
        return round(numeric_value, 10)
    if isinstance(value, str):
        return value.strip()
    return value


def extract_regression_rows(
    workbook: xw.Book,
    labels: Dict[str, str],
    source_file: str,
) -> List[Dict[str, Any]]:
    sheet = get_sheet_case_insensitive(workbook, "Regression Model")
    if sheet is None:
        print(f"  Regression Model missing in {source_file}; skipped regression extraction.")
        return []

    anchor = find_max_anchor(sheet)
    if anchor is None:
        print(f"  Regression max anchor missing in {source_file}; skipped regression extraction.")
        return []

    anchor_row, anchor_col = anchor
    headers = build_header_lookup(sheet, anchor_row, anchor_col)
    y_col = anchor_col - 7
    x_col = anchor_col - 11
    n_quarters = 10
    start_row = anchor_row + 1
    end_row = start_row + n_quarters - 1

    col_map = {
        "num_quarters_used": find_column(
            headers,
            ["num_quarters_used", "num_quarters", "quarters_used", "n_quarters"],
            anchor_col - 10,
        ),
        "forecast_value": find_column(
            headers,
            [
                "tot_fcst_without_sa",
                "tot_fcst_wo_sa",
                "tot_fcst_w_o_sa",
                "forecast_value",
                "tot_fcst",
            ],
            anchor_col - 4,
        ),
        "actual_value": find_column(headers, ["actual_value", "actual", "reported_sales"], None),
        "forecast_max": anchor_col,
        "forecast_min": find_column(headers, ["forecast_min", "min"], anchor_col + 1),
    }

    required_columns = [col for col in col_map.values() if col is not None]
    if not required_columns:
        return []

    min_col = min(required_columns)
    max_col = max(required_columns)
    data_matrix = read_matrix(sheet.range((start_row, min_col), (end_row, max_col)))

    # Use R1C1 + formula2 for INTERCEPT and SLOPE once per candidate row.
    temp_intercept_col = sheet.used_range.last_cell.column + 1
    temp_slope_col = temp_intercept_col + 1
    formula_written = False
    historical_end_row = anchor_row - 1

    if historical_end_row >= 1:
        for idx in range(n_quarters):
            row = start_row + idx
            used_quarters = idx + 1
            historical_start_row = max(1, historical_end_row - used_quarters + 1)
            intercept_formula = (
                f'=IFERROR(INTERCEPT(R{historical_start_row}C{y_col}:R{historical_end_row}C{y_col},'
                f'R{historical_start_row}C{x_col}:R{historical_end_row}C{x_col}),"")'
            )
            slope_formula = (
                f'=IFERROR(SLOPE(R{historical_start_row}C{y_col}:R{historical_end_row}C{y_col},'
                f'R{historical_start_row}C{x_col}:R{historical_end_row}C{x_col}),"")'
            )
            set_formula2_or_formula(sheet.cells(row, temp_intercept_col), intercept_formula)
            set_formula2_or_formula(sheet.cells(row, temp_slope_col), slope_formula)
            formula_written = True

    if formula_written:
        workbook.app.calculate()

    intercept_values = [row[0] for row in read_matrix(sheet.range((start_row, temp_intercept_col), (end_row, temp_intercept_col)))]
    slope_values = [row[0] for row in read_matrix(sheet.range((start_row, temp_slope_col), (end_row, temp_slope_col)))]

    rows: List[Dict[str, Any]] = []
    previous_signature: Optional[Tuple[Any, ...]] = None

    for idx in range(n_quarters):
        num_quarters_value = matrix_cell(data_matrix, idx, col_map["num_quarters_used"], min_col)
        num_quarters_used = num_quarters_value if not is_blank(num_quarters_value) else idx + 1

        forecast_value = matrix_cell(data_matrix, idx, col_map["forecast_value"], min_col)
        actual_value = matrix_cell(data_matrix, idx, col_map["actual_value"], min_col)
        forecast_max = matrix_cell(data_matrix, idx, col_map["forecast_max"], min_col)
        forecast_min = matrix_cell(data_matrix, idx, col_map["forecast_min"], min_col)
        intercept = intercept_values[idx] if idx < len(intercept_values) else None
        slope = slope_values[idx] if idx < len(slope_values) else None
        range_width = numeric_diff(forecast_max, forecast_min)

        if all(
            is_blank(value)
            for value in (forecast_value, forecast_max, forecast_min, intercept, slope)
        ):
            continue

        signature = (
            signature_value(num_quarters_used),
            signature_value(forecast_value),
            signature_value(actual_value),
            signature_value(forecast_max),
            signature_value(forecast_min),
            signature_value(intercept),
            signature_value(slope),
        )
        if previous_signature is not None and signature == previous_signature:
            continue

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
        previous_signature = signature

    return rows


def write_sheet(
    workbook: Workbook,
    title: str,
    headers: Sequence[str],
    rows: List[Dict[str, Any]],
) -> None:
    worksheet = workbook.active if workbook.active.title == "Sheet" and workbook.active.max_row == 1 else workbook.create_sheet()
    worksheet.title = title

    worksheet.append(list(headers))
    widths = [len(header) for header in headers]

    for row_dict in rows:
        row_values = [row_dict.get(header) for header in headers]
        worksheet.append(row_values)
        for idx, value in enumerate(row_values):
            text = "" if value is None else str(value)
            widths[idx] = min(max(widths[idx], len(text)), 60)

    for cell in worksheet[1]:
        cell.font = Font(bold=True)

    worksheet.freeze_panes = "A2"
    last_col_letter = get_column_letter(len(headers))
    last_row = max(worksheet.max_row, 1)
    worksheet.auto_filter.ref = f"A1:{last_col_letter}{last_row}"

    for idx, width in enumerate(widths, start=1):
        worksheet.column_dimensions[get_column_letter(idx)].width = width + 2


def write_output_workbook(
    output_path: Path,
    empirical_rows: List[Dict[str, Any]],
    regression_rows: List[Dict[str, Any]],
) -> None:
    workbook = Workbook()
    # Remove default sheet handling in write_sheet by renaming it there.
    write_sheet(workbook, "empirical_candidates", EMPIRICAL_COLUMNS, empirical_rows)
    write_sheet(workbook, "regression_candidates", REGRESSION_COLUMNS, regression_rows)

    if "Sheet" in workbook.sheetnames:
        del workbook["Sheet"]
    workbook.save(output_path)


def main() -> None:
    input_path = Path(input_dir).expanduser().resolve()
    output_path = Path(output_dir).expanduser().resolve()

    if not input_path.exists() or not input_path.is_dir():
        raise FileNotFoundError(f"input_dir does not exist or is not a directory: {input_path}")
    output_path.mkdir(parents=True, exist_ok=True)

    output_file = build_output_path(input_path, output_path)
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

        for file_path in sorted(input_path.iterdir()):
            if not file_path.is_file():
                print(f"Skipped {file_path.name}: not a file.")
                continue
            if file_path.name.startswith("~"):
                print(f"Skipped {file_path.name}: temporary Excel file.")
                continue
            if file_path.suffix.lower() != ".xlsx":
                print(f"Skipped {file_path.name}: not an .xlsx file.")
                continue

            print(f"Processing {file_path.name}")
            workbook: Optional[xw.Book] = None
            try:
                workbook = app.books.open(str(file_path), update_links=False)
                labels = parse_file_labels(file_path.name)
                empirical_rows.extend(extract_empirical_rows(workbook, labels, file_path.name))
                regression_rows.extend(extract_regression_rows(workbook, labels, file_path.name))
                processed_files += 1
            except Exception as exc:
                print(f"Skipped {file_path.name}: processing error ({exc}).")
            finally:
                if workbook is not None:
                    close_workbook_no_save(workbook)
    finally:
        if app is not None:
            try:
                app.quit()
            except Exception:
                pass

    write_output_workbook(output_file, empirical_rows, regression_rows)
    print(f"Output saved: {output_file}")
    print(f"Files processed: {processed_files}")
    print(f"Empirical rows: {len(empirical_rows)}")
    print(f"Regression rows: {len(regression_rows)}")


if __name__ == "__main__":
    main()
