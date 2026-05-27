#!/usr/bin/env python3
from __future__ import annotations

import calendar
import re
from dataclasses import dataclass
from datetime import date
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Sequence, Tuple

import xlwings as xw
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

# User-configurable paths
input_dir = Path("input")
output_dir = Path("output")

N_QUARTERS = 10

EMPIRICAL_SHEET_NAME = "Empirical Model"
REGRESSION_SHEET_NAME = "Regression Model"

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
class FileLabels:
    model: str
    ticker: str
    model_period: str
    model_date: str


def normalize_label(value: Any) -> str:
    if value is None:
        return ""
    text = str(value).strip().lower()
    text = text.replace("%", " pct ").replace("/", " ").replace("-", " ").replace("_", " ")
    text = re.sub(r"\s+", " ", text)
    return text.strip()


def to_number(value: Any) -> Optional[float]:
    if value is None or value == "":
        return None
    if isinstance(value, (int, float)):
        return float(value)
    try:
        return float(str(value).replace(",", ""))
    except ValueError:
        return None


def calc_range_width(max_value: Any, min_value: Any) -> Optional[float]:
    max_num = to_number(max_value)
    min_num = to_number(min_value)
    if max_num is None or min_num is None:
        return None
    return max_num - min_num


def as_matrix(values: Any, n_rows: int, n_cols: int) -> List[List[Any]]:
    if n_rows <= 0 or n_cols <= 0:
        return []
    if n_rows == 1 and n_cols == 1:
        return [[values]]
    if n_rows == 1:
        if isinstance(values, list):
            return [values]
        return [[values]]
    if n_cols == 1:
        if isinstance(values, list):
            return [[item] for item in values]
        return [[values] for _ in range(n_rows)]
    return values


def read_range_matrix(
    sheet: xw.Sheet, start_row: int, start_col: int, end_row: int, end_col: int
) -> List[List[Any]]:
    values = sheet.range((start_row, start_col), (end_row, end_col)).value
    return as_matrix(values, end_row - start_row + 1, end_col - start_col + 1)


def safe_close_workbook(wb: xw.Book) -> None:
    try:
        wb.close(save=False)
        return
    except TypeError:
        pass
    except Exception:
        pass

    try:
        wb.api.Close(SaveChanges=False)
    except Exception:
        try:
            wb.close()
        except Exception:
            pass


def find_anchor_cell(sheet: xw.Sheet, label: str = "max") -> Optional[Tuple[int, int]]:
    used = sheet.used_range
    rows = used.last_cell.row - used.row + 1
    cols = used.last_cell.column - used.column + 1
    matrix = as_matrix(used.value, rows, cols)
    for r_idx, row in enumerate(matrix):
        for c_idx, value in enumerate(row):
            if isinstance(value, str) and value.strip().lower() == label:
                return used.row + r_idx, used.column + c_idx
    return None


def build_header_map(
    sheet: xw.Sheet, header_row: int, anchor_col: int, radius: int = 30
) -> Dict[str, int]:
    start_col = max(1, anchor_col - radius)
    end_col = anchor_col + radius
    values = sheet.range((header_row, start_col), (header_row, end_col)).value
    if not isinstance(values, list):
        values = [values]

    header_map: Dict[str, int] = {}
    for index, value in enumerate(values):
        normalized = normalize_label(value)
        if normalized and normalized not in header_map:
            header_map[normalized] = start_col + index
    return header_map


def find_col_by_keywords(
    header_map: Dict[str, int], keyword_sets: Sequence[Sequence[str]]
) -> Optional[int]:
    for key, col in header_map.items():
        for keywords in keyword_sets:
            if all(keyword in key for keyword in keywords):
                return col
    return None


def set_formula2_r1c1(cell: xw.Range, formula_r1c1: str) -> None:
    try:
        cell.formula2 = formula_r1c1
    except Exception:
        try:
            cell.api.Formula2R1C1 = formula_r1c1
        except Exception:
            cell.api.FormulaR1C1 = formula_r1c1


def month_to_number(month_text: str) -> Optional[int]:
    cleaned = re.sub(r"[^a-zA-Z]", "", month_text).lower()
    if not cleaned:
        return None
    month_names = [name.lower() for name in calendar.month_name if name]
    month_abbr = [abbr.lower() for abbr in calendar.month_abbr if abbr]
    for idx, (name, abbr) in enumerate(zip(month_names, month_abbr), start=1):
        if cleaned in {name, abbr, name[:3], "sept" if idx == 9 else ""}:
            return idx
    return None


def parse_file_labels(file_path: Path) -> FileLabels:
    stem = file_path.stem
    parts = [part.strip() for part in stem.split(" - ")]
    ticker = parts[1] if len(parts) >= 2 else ""
    period_raw = parts[2] if len(parts) >= 3 else ""
    period_token = period_raw.split("_")[0].strip()

    match = re.search(r"(Early|Mid|Late)([A-Za-z]+)(\d{4})", period_token, flags=re.IGNORECASE)
    model_period = ""
    model_date = ""
    if match:
        period_prefix = match.group(1).title()
        month_part = match.group(2)
        year = int(match.group(3))
        month_num = month_to_number(month_part)
        if month_num is not None:
            day_lookup = {"Early": 5, "Mid": 15, "Late": 25}
            day = day_lookup[period_prefix]
            model_period = f"{period_prefix}{calendar.month_abbr[month_num]}_{year}"
            model_date = date(year, month_num, day).isoformat()

    if not model_period:
        model_period = period_token or "unknown_period"
    if not model_date:
        model_date = ""

    if ticker:
        model = f"{ticker}_{model_period}"
    else:
        model = model_period

    return FileLabels(model=model, ticker=ticker, model_period=model_period, model_date=model_date)


def get_output_path(input_root: Path, output_root: Path) -> Path:
    input_folder_name = input_root.resolve().name
    base = output_root / f"{input_folder_name}_PARAM.xlsx"
    if not base.exists():
        return base

    counter = 1
    while True:
        candidate = output_root / f"{input_folder_name}_PARAM.{counter}.xlsx"
        if not candidate.exists():
            return candidate
        counter += 1


def get_matrix_value(matrix: List[List[Any]], row_idx: int, col_idx: int) -> Any:
    if row_idx < 0 or col_idx < 0:
        return None
    if row_idx >= len(matrix):
        return None
    row = matrix[row_idx]
    if col_idx >= len(row):
        return None
    return row[col_idx]


def extract_empirical_rows(wb: xw.Book, labels: FileLabels, source_file: str) -> List[Dict[str, Any]]:
    try:
        sheet = wb.sheets[EMPIRICAL_SHEET_NAME]
    except Exception:
        print(f"  - Skipped empirical extraction: sheet '{EMPIRICAL_SHEET_NAME}' not found")
        return []

    anchor = find_anchor_cell(sheet, "max")
    if anchor is None:
        print("  - Skipped empirical extraction: 'max' anchor not found")
        return []

    anchor_row, anchor_col = anchor
    header_map = build_header_map(sheet, anchor_row, anchor_col)
    row_start = anchor_row + 1
    row_end = row_start + N_QUARTERS - 1

    col_forecast_max = anchor_col
    col_forecast_min = find_col_by_keywords(header_map, [("min",)]) or (anchor_col + 1)
    col_forecast_value = find_col_by_keywords(
        header_map,
        [
            ("estimated", "total", "sold"),
            ("forecast", "value"),
            ("tot", "fcst"),
        ],
    ) or (anchor_col - 1)
    col_reported_sales = find_col_by_keywords(
        header_map,
        [
            ("reported", "sales"),
            ("actual", "sales"),
        ],
    ) or (anchor_col - 2)
    col_quarterly_sales = find_col_by_keywords(
        header_map, [("quarterly", "sales"), ("quarterly",)]
    ) or (anchor_col - 3)
    col_growth_rate_pct = find_col_by_keywords(
        header_map, [("growth", "rate"), ("growth", "pct")]
    ) or (anchor_col - 4)
    col_sales_captured = find_col_by_keywords(
        header_map,
        [
            ("sales", "captured", "db"),
            ("captured", "db"),
            ("penetration", "pct"),
        ],
    ) or (anchor_col - 5)
    col_last_quarter_used = find_col_by_keywords(
        header_map,
        [
            ("last", "quarter", "used"),
            ("last", "quarter"),
        ],
    ) or (anchor_col - 6)
    col_num_quarters = find_col_by_keywords(
        header_map,
        [
            ("num", "quarters", "used"),
            ("quarters", "used"),
        ],
    )

    used_last_col = sheet.used_range.last_cell.column
    helper_col = used_last_col + 2
    for offset in range(N_QUARTERS):
        row = row_start + offset
        n_quarters = offset + 1
        start_row = row - n_quarters + 1
        avg_formula = f"=AVERAGE(R{start_row}C{col_sales_captured}:R{row}C{col_sales_captured})"
        set_formula2_r1c1(sheet.range((row, helper_col)), avg_formula)

    wb.app.calculate()
    avg_values = sheet.range((row_start, helper_col), (row_end, helper_col)).value
    if not isinstance(avg_values, list):
        avg_values = [avg_values]

    required_cols = [
        col_forecast_max,
        col_forecast_min,
        col_forecast_value,
        col_reported_sales,
        col_quarterly_sales,
        col_growth_rate_pct,
        col_sales_captured,
        col_last_quarter_used,
    ]
    if col_num_quarters is not None:
        required_cols.append(col_num_quarters)
    min_col = min(required_cols)
    max_col = max(required_cols)
    matrix = read_range_matrix(sheet, row_start, min_col, row_end, max_col)

    rows: List[Dict[str, Any]] = []
    for offset in range(N_QUARTERS):
        row_idx = offset
        n_quarters_default = offset + 1
        num_quarters_used = (
            get_matrix_value(matrix, row_idx, col_num_quarters - min_col)
            if col_num_quarters is not None
            else n_quarters_default
        )
        if num_quarters_used in (None, ""):
            num_quarters_used = n_quarters_default

        forecast_max = get_matrix_value(matrix, row_idx, col_forecast_max - min_col)
        forecast_min = get_matrix_value(matrix, row_idx, col_forecast_min - min_col)
        forecast_value = get_matrix_value(matrix, row_idx, col_forecast_value - min_col)
        reported_sales = get_matrix_value(matrix, row_idx, col_reported_sales - min_col)
        quarterly_sales = get_matrix_value(matrix, row_idx, col_quarterly_sales - min_col)
        growth_rate_pct = get_matrix_value(matrix, row_idx, col_growth_rate_pct - min_col)
        sales_captured_in_db_pct = get_matrix_value(matrix, row_idx, col_sales_captured - min_col)
        last_quarter_used = get_matrix_value(matrix, row_idx, col_last_quarter_used - min_col)
        avg_penetration_pct = avg_values[row_idx] if row_idx < len(avg_values) else None

        rows.append(
            {
                "model": labels.model,
                "ticker": labels.ticker,
                "model_period": labels.model_period,
                "model_date": labels.model_date,
                "method": "empirical",
                "parameter_name": "avg_penetration_pct",
                "parameter_value": avg_penetration_pct,
                "num_quarters_used": num_quarters_used,
                "last_quarter_used": last_quarter_used,
                "forecast_value": forecast_value,
                "actual_value": reported_sales,
                "forecast_max": forecast_max,
                "forecast_min": forecast_min,
                "range_width": calc_range_width(forecast_max, forecast_min),
                "avg_penetration_pct": avg_penetration_pct,
                "quarterly_sales": quarterly_sales,
                "reported_sales": reported_sales,
                "growth_rate_pct": growth_rate_pct,
                "sales_captured_in_db_pct": sales_captured_in_db_pct,
                "source_file": source_file,
            }
        )

    return rows


def value_signature(values: Iterable[Any]) -> Tuple[Any, ...]:
    signature: List[Any] = []
    for value in values:
        num = to_number(value)
        if num is not None:
            signature.append(round(num, 10))
        elif value is None:
            signature.append(None)
        else:
            signature.append(str(value).strip())
    return tuple(signature)


def extract_regression_rows(wb: xw.Book, labels: FileLabels, source_file: str) -> List[Dict[str, Any]]:
    try:
        sheet = wb.sheets[REGRESSION_SHEET_NAME]
    except Exception:
        print(f"  - Skipped regression extraction: sheet '{REGRESSION_SHEET_NAME}' not found")
        return []

    anchor = find_anchor_cell(sheet, "max")
    if anchor is None:
        print("  - Skipped regression extraction: 'max' anchor not found")
        return []

    anchor_row, anchor_col = anchor
    header_map = build_header_map(sheet, anchor_row, anchor_col)
    row_start = anchor_row + 1
    row_end = row_start + N_QUARTERS - 1

    y_col = anchor_col - 7
    x_col = anchor_col - 11

    col_num_quarters = find_col_by_keywords(
        header_map, [("num", "quarters", "used"), ("quarters", "used")]
    )
    col_forecast_value = find_col_by_keywords(
        header_map,
        [
            ("tot", "fcst", "w", "o", "sa"),
            ("tot", "fcst", "wo", "sa"),
            ("forecast", "without", "sa"),
        ],
    ) or (anchor_col - 1)
    col_forecast_max = anchor_col
    col_forecast_min = find_col_by_keywords(header_map, [("min",)]) or (anchor_col + 1)

    used_last_col = sheet.used_range.last_cell.column
    helper_intercept_col = used_last_col + 2
    helper_slope_col = used_last_col + 3
    for offset in range(N_QUARTERS):
        row = row_start + offset
        n_quarters = offset + 1
        start_row = row - n_quarters + 1
        intercept_formula = (
            f"=INTERCEPT(R{start_row}C{y_col}:R{row}C{y_col},R{start_row}C{x_col}:R{row}C{x_col})"
        )
        slope_formula = (
            f"=SLOPE(R{start_row}C{y_col}:R{row}C{y_col},R{start_row}C{x_col}:R{row}C{x_col})"
        )
        set_formula2_r1c1(sheet.range((row, helper_intercept_col)), intercept_formula)
        set_formula2_r1c1(sheet.range((row, helper_slope_col)), slope_formula)

    wb.app.calculate()
    intercept_values = sheet.range(
        (row_start, helper_intercept_col), (row_end, helper_intercept_col)
    ).value
    slope_values = sheet.range((row_start, helper_slope_col), (row_end, helper_slope_col)).value
    if not isinstance(intercept_values, list):
        intercept_values = [intercept_values]
    if not isinstance(slope_values, list):
        slope_values = [slope_values]

    required_cols = [col_forecast_value, col_forecast_max, col_forecast_min]
    if col_num_quarters is not None:
        required_cols.append(col_num_quarters)
    min_col = min(required_cols)
    max_col = max(required_cols)
    matrix = read_range_matrix(sheet, row_start, min_col, row_end, max_col)

    rows: List[Dict[str, Any]] = []
    previous_signature: Optional[Tuple[Any, ...]] = None
    for offset in range(N_QUARTERS):
        row_idx = offset
        n_quarters_default = offset + 1
        num_quarters_used = (
            get_matrix_value(matrix, row_idx, col_num_quarters - min_col)
            if col_num_quarters is not None
            else n_quarters_default
        )
        if num_quarters_used in (None, ""):
            num_quarters_used = n_quarters_default

        forecast_value = get_matrix_value(matrix, row_idx, col_forecast_value - min_col)
        forecast_max = get_matrix_value(matrix, row_idx, col_forecast_max - min_col)
        forecast_min = get_matrix_value(matrix, row_idx, col_forecast_min - min_col)
        intercept = intercept_values[row_idx] if row_idx < len(intercept_values) else None
        slope = slope_values[row_idx] if row_idx < len(slope_values) else None

        current_signature = value_signature(
            [forecast_value, forecast_max, forecast_min, intercept, slope]
        )
        if row_idx == N_QUARTERS - 1 and previous_signature == current_signature:
            continue
        previous_signature = current_signature

        rows.append(
            {
                "model": labels.model,
                "ticker": labels.ticker,
                "model_period": labels.model_period,
                "model_date": labels.model_date,
                "method": "regression",
                "parameter_name": "num_quarters_used",
                "parameter_value": num_quarters_used,
                "num_quarters_used": num_quarters_used,
                "forecast_value": forecast_value,
                "actual_value": "",
                "forecast_max": forecast_max,
                "forecast_min": forecast_min,
                "range_width": calc_range_width(forecast_max, forecast_min),
                "intercept": intercept,
                "slope": slope,
                "source_file": source_file,
            }
        )

    return rows


def autosize_columns(ws) -> None:
    for col_idx in range(1, ws.max_column + 1):
        max_len = 0
        for row_idx in range(1, ws.max_row + 1):
            value = ws.cell(row=row_idx, column=col_idx).value
            if value is None:
                continue
            max_len = max(max_len, len(str(value)))
        ws.column_dimensions[get_column_letter(col_idx)].width = min(max(max_len + 2, 12), 48)


def write_output_workbook(
    output_path: Path, empirical_rows: List[Dict[str, Any]], regression_rows: List[Dict[str, Any]]
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
        for cell in ws[1]:
            cell.font = Font(bold=True)
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions
        autosize_columns(ws)

    wb.save(output_path)


def iter_input_files(root: Path) -> Iterable[Path]:
    for file_path in sorted(root.iterdir()):
        if file_path.is_file():
            yield file_path


def main() -> None:
    input_root = Path(input_dir)
    output_root = Path(output_dir)

    if not input_root.exists() or not input_root.is_dir():
        raise FileNotFoundError(f"Input folder not found: {input_root}")

    output_root.mkdir(parents=True, exist_ok=True)
    output_path = get_output_path(input_root, output_root)

    empirical_rows: List[Dict[str, Any]] = []
    regression_rows: List[Dict[str, Any]] = []
    files_processed = 0

    with xw.App(visible=False, add_book=False) as app:
        app.display_alerts = False
        app.screen_updating = False

        for file_path in iter_input_files(input_root):
            if file_path.name.startswith("~"):
                print(f"Skipped {file_path.name}: temporary workbook")
                continue
            if file_path.suffix.lower() != ".xlsx":
                print(f"Skipped {file_path.name}: not an .xlsx file")
                continue

            print(f"Processing {file_path.name}")
            wb = None
            try:
                wb = app.books.open(str(file_path), update_links=False)
                labels = parse_file_labels(file_path)
                empirical_rows.extend(extract_empirical_rows(wb, labels, file_path.name))
                regression_rows.extend(extract_regression_rows(wb, labels, file_path.name))
                files_processed += 1
            except Exception as exc:
                print(f"Skipped {file_path.name}: processing error: {exc}")
            finally:
                if wb is not None:
                    safe_close_workbook(wb)

    write_output_workbook(output_path, empirical_rows, regression_rows)

    print(f"Output path: {output_path}")
    print(f"Files processed: {files_processed}")
    print(f"Empirical rows: {len(empirical_rows)}")
    print(f"Regression rows: {len(regression_rows)}")


if __name__ == "__main__":
    main()
