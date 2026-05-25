#!/usr/bin/env python3
"""
Batch-extract empirical/regression model candidates from Excel workbooks.

The script opens each source workbook once, processes both the
"Empirical Model" and "Regression Model" sheets while open, and writes one
combined output workbook with two tabs:
    - empirical_candidates
    - regression_candidates
"""

from __future__ import annotations

import math
import re
from dataclasses import dataclass
from datetime import date
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Sequence, Tuple

from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.worksheet import Worksheet

try:
    import xlwings as xw
except ModuleNotFoundError as exc:
    xw = None
    _XLWINGS_IMPORT_ERROR = exc
else:
    _XLWINGS_IMPORT_ERROR = None

# ===========================
# User-configurable locations
# ===========================
input_dir = Path("/path/to/input")
output_dir = Path("/path/to/output")

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

EMPIRICAL_ALIASES = {
    "num_quarters_used": ["num quarters used", "quarters used", "num quarters", "n quarters"],
    "last_quarter_used": ["last quarter used", "last quarter", "last qtr", "last quarter in sample"],
    "forecast_value": [
        "estimated total sold",
        "forecast value",
        "forecast",
        "total forecast",
        "tot fcst",
    ],
    "actual_value": ["reported sales", "actual value", "actual", "reported"],
    "avg_penetration_pct": [
        "avg penetration pct",
        "average penetration",
        "avg penetration",
        "penetration",
    ],
    "quarterly_sales": ["quarterly sales", "quarter sales", "sales"],
    "reported_sales": ["reported sales", "reported"],
    "growth_rate_pct": ["growth rate pct", "growth rate", "growth"],
    "sales_captured_in_db_pct": [
        "sales captured in db pct",
        "sales captured in db",
        "captured in db",
    ],
    "forecast_min": ["min", "forecast min", "minimum"],
}

REGRESSION_ALIASES = {
    "num_quarters_used": ["num quarters used", "quarters used", "num quarters", "n quarters"],
    "forecast_value": [
        "tot fcst w/o sa",
        "tot fcst wo sa",
        "forecast total without sa",
        "forecast without sa",
    ],
    "actual_value": ["actual", "reported sales", "actual value"],
    "forecast_min": ["min", "forecast min", "minimum"],
}

MONTH_TO_NUMBER = {
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

PHASE_TO_DAY = {"early": 5, "mid": 15, "late": 25}


@dataclass
class FileLabels:
    model: str
    ticker: str
    model_period: str
    model_date: str


def normalize_key(value: Any) -> str:
    if value is None:
        return ""
    return re.sub(r"[^a-z0-9]+", "", str(value).strip().lower())


def is_number(value: Any) -> bool:
    if value is None or isinstance(value, bool):
        return False
    if isinstance(value, (int, float)):
        return not (isinstance(value, float) and math.isnan(value))
    text = str(value).strip().replace(",", "")
    if not text:
        return False
    try:
        float(text)
        return True
    except ValueError:
        return False


def to_float(value: Any) -> Optional[float]:
    if not is_number(value):
        return None
    if isinstance(value, (int, float)):
        return float(value)
    return float(str(value).strip().replace(",", ""))


def to_int(value: Any) -> Optional[int]:
    number = to_float(value)
    if number is None:
        return None
    return int(round(number))


def clean_value(value: Any) -> Any:
    if isinstance(value, float) and math.isnan(value):
        return None
    if isinstance(value, str):
        trimmed = value.strip()
        return trimmed if trimmed else None
    return value


def subtract_if_numeric(max_value: Any, min_value: Any) -> Optional[float]:
    max_num = to_float(max_value)
    min_num = to_float(min_value)
    if max_num is None or min_num is None:
        return None
    return max_num - min_num


def has_content(values: Iterable[Any]) -> bool:
    for value in values:
        cleaned = clean_value(value)
        if cleaned is None:
            continue
        return True
    return False


def normalize_signature_value(value: Any) -> Any:
    number = to_float(value)
    if number is not None:
        return round(number, 10)
    cleaned = clean_value(value)
    return cleaned


def parse_file_labels(file_path: Path) -> FileLabels:
    stem = file_path.stem
    parts = [part.strip() for part in stem.split(" - ")]

    ticker = ""
    if len(parts) >= 2:
        ticker = re.sub(r"[^A-Za-z0-9]+", "", parts[1]).upper()
    if not ticker:
        ticker_match = re.search(r"-\s*([A-Za-z0-9]{1,10})\s*-", stem)
        if ticker_match:
            ticker = ticker_match.group(1).upper()

    period_source = parts[2] if len(parts) >= 3 else stem
    period_source = period_source.split("_")[0]

    period_match = re.search(
        r"(?i)\b(early|mid|late)(jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)(\d{4})\b",
        period_source,
    )
    if period_match is None:
        period_match = re.search(
            r"(?i)\b(early|mid|late)(jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)(\d{4})\b",
            stem,
        )

    model_period = ""
    model_date = ""
    if period_match:
        phase = period_match.group(1).title()
        month_abbrev = period_match.group(2).title()
        year = period_match.group(3)
        month_num = MONTH_TO_NUMBER[month_abbrev.lower()]
        day = PHASE_TO_DAY[phase.lower()]
        model_period = f"{phase}{month_abbrev}_{year}"
        model_date = date(int(year), month_num, day).isoformat()

    if not model_period:
        model_period = "unknown_period"
    if not ticker:
        ticker = "UNKNOWN"

    model = f"{ticker}_{model_period}"
    return FileLabels(model=model, ticker=ticker, model_period=model_period, model_date=model_date)


def next_output_path(input_folder: Path, output_folder: Path) -> Path:
    output_folder.mkdir(parents=True, exist_ok=True)
    base_name = f"{input_folder.name}_PARAM"
    candidate = output_folder / f"{base_name}.xlsx"
    if not candidate.exists():
        return candidate

    suffix = 1
    while True:
        numbered = output_folder / f"{base_name}.{suffix}.xlsx"
        if not numbered.exists():
            return numbered
        suffix += 1


def to_2d(values: Any) -> List[List[Any]]:
    if values is None:
        return []
    if not isinstance(values, (list, tuple)):
        return [[values]]
    if not values:
        return []
    if isinstance(values[0], (list, tuple)):
        return [list(row) for row in values]
    return [list(values)]


def get_sheet_case_insensitive(workbook: xw.Book, sheet_name: str) -> Optional[xw.Sheet]:
    target = sheet_name.strip().lower()
    for sheet in workbook.sheets:
        if sheet.name.strip().lower() == target:
            return sheet
    return None


def find_anchor_cell(sheet: xw.Sheet, anchor_text: str = "max") -> Optional[Tuple[int, int, int, int, int, int]]:
    used_range = sheet.used_range
    used_values = to_2d(used_range.value)
    if not used_values:
        return None

    first_row = used_range.row
    first_col = used_range.column
    last_row = used_range.last_cell.row
    last_col = used_range.last_cell.column
    anchor_key = normalize_key(anchor_text)

    for row_offset, row_values in enumerate(used_values):
        for col_offset, cell_value in enumerate(row_values):
            if normalize_key(cell_value) == anchor_key:
                return (
                    first_row + row_offset,
                    first_col + col_offset,
                    first_row,
                    first_col,
                    last_row,
                    last_col,
                )
    return None


def build_header_map(sheet: xw.Sheet, header_row: int, start_col: int, end_col: int) -> Dict[str, int]:
    if start_col > end_col:
        return {}
    values = sheet.range((header_row, start_col), (header_row, end_col)).value
    if not isinstance(values, list):
        values = [values]

    header_map: Dict[str, int] = {}
    for index, raw_value in enumerate(values):
        key = normalize_key(raw_value)
        if key:
            header_map[key] = start_col + index
    return header_map


def resolve_column(header_map: Dict[str, int], aliases: Sequence[str]) -> Optional[int]:
    normalized_aliases = [normalize_key(alias) for alias in aliases]
    for alias in normalized_aliases:
        if alias in header_map:
            return header_map[alias]

    for alias in normalized_aliases:
        for header_key, column in header_map.items():
            if alias and (alias in header_key or header_key in alias):
                return column
    return None


def fallback_column(anchor_col: int, offset: int, min_col: int, max_col: int) -> Optional[int]:
    candidate = anchor_col + offset
    if min_col <= candidate <= max_col:
        return candidate
    return None


def close_workbook_safely(workbook: xw.Book) -> None:
    try:
        workbook.close(save=False)
        return
    except TypeError:
        pass
    except Exception:
        pass

    try:
        workbook.api.Close(False)
        return
    except Exception:
        pass

    try:
        workbook.close()
    except Exception:
        pass


def extract_empirical_rows(workbook: xw.Book, labels: FileLabels, source_file: str) -> List[Dict[str, Any]]:
    rows: List[Dict[str, Any]] = []
    sheet = get_sheet_case_insensitive(workbook, "Empirical Model")
    if sheet is None:
        print(f"Skipped empirical extraction in {source_file}: missing 'Empirical Model' sheet")
        return rows

    anchor = find_anchor_cell(sheet, "max")
    if anchor is None:
        print(f"Skipped empirical extraction in {source_file}: could not find 'max' anchor")
        return rows

    anchor_row, anchor_col, used_first_row, used_first_col, used_last_row, used_last_col = anchor
    _ = used_first_row  # documented for readability; retained for parity with anchor return tuple

    header_start_col = max(used_first_col, anchor_col - 20)
    header_end_col = min(used_last_col, anchor_col + 20)
    header_map = build_header_map(sheet, anchor_row, header_start_col, header_end_col)

    columns: Dict[str, Optional[int]] = {
        "forecast_max": anchor_col,
        "forecast_min": resolve_column(header_map, EMPIRICAL_ALIASES["forecast_min"]),
        "num_quarters_used": resolve_column(header_map, EMPIRICAL_ALIASES["num_quarters_used"]),
        "last_quarter_used": resolve_column(header_map, EMPIRICAL_ALIASES["last_quarter_used"]),
        "forecast_value": resolve_column(header_map, EMPIRICAL_ALIASES["forecast_value"]),
        "actual_value": resolve_column(header_map, EMPIRICAL_ALIASES["actual_value"]),
        "avg_penetration_pct": resolve_column(header_map, EMPIRICAL_ALIASES["avg_penetration_pct"]),
        "quarterly_sales": resolve_column(header_map, EMPIRICAL_ALIASES["quarterly_sales"]),
        "reported_sales": resolve_column(header_map, EMPIRICAL_ALIASES["reported_sales"]),
        "growth_rate_pct": resolve_column(header_map, EMPIRICAL_ALIASES["growth_rate_pct"]),
        "sales_captured_in_db_pct": resolve_column(
            header_map, EMPIRICAL_ALIASES["sales_captured_in_db_pct"]
        ),
    }

    fallback_offsets = {
        "forecast_min": 1,
        "num_quarters_used": -6,
        "last_quarter_used": -5,
        "forecast_value": -3,
        "actual_value": -2,
        "avg_penetration_pct": -4,
        "quarterly_sales": -9,
        "reported_sales": -8,
        "growth_rate_pct": -7,
        "sales_captured_in_db_pct": -1,
    }

    for key, offset in fallback_offsets.items():
        if columns.get(key) is None:
            columns[key] = fallback_column(anchor_col, offset, used_first_col, used_last_col)

    temp_avg_col = used_last_col + 3
    formula_rows: List[int] = []

    for idx in range(1, N_QUARTERS + 1):
        row = anchor_row + idx
        if columns["quarterly_sales"] is None and columns["reported_sales"] is None and columns["avg_penetration_pct"] is None:
            break

        if columns["quarterly_sales"] and columns["reported_sales"]:
            formula = (
                f"=IFERROR(RC{columns['quarterly_sales']}/RC{columns['reported_sales']},"
                f"RC{columns['avg_penetration_pct'] or columns['forecast_max']})"
            )
        elif columns["avg_penetration_pct"]:
            formula = f"=RC{columns['avg_penetration_pct']}"
        else:
            formula = f"=RC{columns['forecast_max']}"

        sheet.cells(row, temp_avg_col).formula2 = formula
        formula_rows.append(row)

    if formula_rows:
        workbook.app.calculate()

    for idx in range(1, N_QUARTERS + 1):
        row = anchor_row + idx
        forecast_value = sheet.cells(row, columns["forecast_value"]).value if columns["forecast_value"] else None
        forecast_max = sheet.cells(row, columns["forecast_max"]).value if columns["forecast_max"] else None
        forecast_min = sheet.cells(row, columns["forecast_min"]).value if columns["forecast_min"] else None
        actual_value = sheet.cells(row, columns["actual_value"]).value if columns["actual_value"] else None

        avg_penetration_pct = (
            sheet.cells(row, temp_avg_col).value if row in formula_rows else None
        )
        if avg_penetration_pct is None and columns["avg_penetration_pct"] is not None:
            avg_penetration_pct = sheet.cells(row, columns["avg_penetration_pct"]).value

        num_quarters_used = (
            sheet.cells(row, columns["num_quarters_used"]).value if columns["num_quarters_used"] else idx
        )
        if num_quarters_used is None:
            num_quarters_used = idx

        last_quarter_used = (
            sheet.cells(row, columns["last_quarter_used"]).value if columns["last_quarter_used"] else None
        )
        quarterly_sales = (
            sheet.cells(row, columns["quarterly_sales"]).value if columns["quarterly_sales"] else None
        )
        reported_sales = (
            sheet.cells(row, columns["reported_sales"]).value if columns["reported_sales"] else None
        )
        growth_rate_pct = (
            sheet.cells(row, columns["growth_rate_pct"]).value if columns["growth_rate_pct"] else None
        )
        sales_captured_in_db_pct = (
            sheet.cells(row, columns["sales_captured_in_db_pct"]).value
            if columns["sales_captured_in_db_pct"]
            else None
        )

        probe = [
            num_quarters_used,
            forecast_value,
            actual_value,
            forecast_max,
            forecast_min,
            avg_penetration_pct,
            quarterly_sales,
            reported_sales,
        ]
        if not has_content(probe):
            if rows:
                break
            continue

        rows.append(
            {
                "model": labels.model,
                "ticker": labels.ticker,
                "model_period": labels.model_period,
                "model_date": labels.model_date,
                "method": "empirical",
                "parameter_name": "avg_penetration_pct",
                "parameter_value": clean_value(avg_penetration_pct),
                "num_quarters_used": clean_value(num_quarters_used),
                "last_quarter_used": clean_value(last_quarter_used),
                "forecast_value": clean_value(forecast_value),
                "actual_value": clean_value(actual_value),
                "forecast_max": clean_value(forecast_max),
                "forecast_min": clean_value(forecast_min),
                "range_width": subtract_if_numeric(forecast_max, forecast_min),
                "avg_penetration_pct": clean_value(avg_penetration_pct),
                "quarterly_sales": clean_value(quarterly_sales),
                "reported_sales": clean_value(reported_sales),
                "growth_rate_pct": clean_value(growth_rate_pct),
                "sales_captured_in_db_pct": clean_value(sales_captured_in_db_pct),
                "source_file": source_file,
            }
        )

    return rows


def find_last_history_row(sheet: xw.Sheet, anchor_row: int, x_col: int, y_col: int) -> Optional[int]:
    for row in range(anchor_row - 1, 0, -1):
        x_value = sheet.cells(row, x_col).value
        y_value = sheet.cells(row, y_col).value
        if is_number(x_value) and is_number(y_value):
            return row
    return None


def extract_regression_rows(workbook: xw.Book, labels: FileLabels, source_file: str) -> List[Dict[str, Any]]:
    rows: List[Dict[str, Any]] = []
    sheet = get_sheet_case_insensitive(workbook, "Regression Model")
    if sheet is None:
        print(f"Skipped regression extraction in {source_file}: missing 'Regression Model' sheet")
        return rows

    anchor = find_anchor_cell(sheet, "max")
    if anchor is None:
        print(f"Skipped regression extraction in {source_file}: could not find 'max' anchor")
        return rows

    anchor_row, anchor_col, used_first_row, used_first_col, _, used_last_col = anchor
    _ = used_first_row
    y_col = anchor_col - 7
    x_col = anchor_col - 11

    header_start_col = max(used_first_col, anchor_col - 20)
    header_end_col = anchor_col + 20
    header_map = build_header_map(sheet, anchor_row, header_start_col, header_end_col)

    columns: Dict[str, Optional[int]] = {
        "forecast_max": anchor_col,
        "forecast_min": resolve_column(header_map, REGRESSION_ALIASES["forecast_min"]),
        "num_quarters_used": resolve_column(header_map, REGRESSION_ALIASES["num_quarters_used"]),
        "forecast_value": resolve_column(header_map, REGRESSION_ALIASES["forecast_value"]),
        "actual_value": resolve_column(header_map, REGRESSION_ALIASES["actual_value"]),
    }

    fallback_offsets = {
        "forecast_min": 1,
        "num_quarters_used": -6,
        "forecast_value": -3,
        "actual_value": -2,
    }
    for key, offset in fallback_offsets.items():
        if columns.get(key) is None:
            columns[key] = fallback_column(anchor_col, offset, used_first_col, header_end_col)

    last_history_row = find_last_history_row(sheet, anchor_row, x_col, y_col)
    temp_intercept_col = used_last_col + 3
    temp_slope_col = used_last_col + 4
    formula_rows: List[Tuple[int, int]] = []

    for idx in range(1, N_QUARTERS + 1):
        row = anchor_row + idx
        num_quarters = (
            sheet.cells(row, columns["num_quarters_used"]).value if columns["num_quarters_used"] else idx
        )
        num_quarters_int = to_int(num_quarters) or idx
        if last_history_row is None:
            continue
        if num_quarters_int <= 0:
            continue

        start_row = max(1, last_history_row - num_quarters_int + 1)
        intercept_formula = (
            f"=IFERROR(INTERCEPT(R{start_row}C{y_col}:R{last_history_row}C{y_col},"
            f"R{start_row}C{x_col}:R{last_history_row}C{x_col}),\"\")"
        )
        slope_formula = (
            f"=IFERROR(SLOPE(R{start_row}C{y_col}:R{last_history_row}C{y_col},"
            f"R{start_row}C{x_col}:R{last_history_row}C{x_col}),\"\")"
        )

        sheet.cells(row, temp_intercept_col).formula2 = intercept_formula
        sheet.cells(row, temp_slope_col).formula2 = slope_formula
        formula_rows.append((row, num_quarters_int))

    if formula_rows:
        workbook.app.calculate()

    previous_signature: Optional[Tuple[Any, ...]] = None

    for idx in range(1, N_QUARTERS + 1):
        row = anchor_row + idx
        num_quarters_used = (
            sheet.cells(row, columns["num_quarters_used"]).value if columns["num_quarters_used"] else idx
        )
        if num_quarters_used is None:
            num_quarters_used = idx

        forecast_value = sheet.cells(row, columns["forecast_value"]).value if columns["forecast_value"] else None
        actual_value = sheet.cells(row, columns["actual_value"]).value if columns["actual_value"] else None
        forecast_max = sheet.cells(row, columns["forecast_max"]).value if columns["forecast_max"] else None
        forecast_min = sheet.cells(row, columns["forecast_min"]).value if columns["forecast_min"] else None
        intercept = sheet.cells(row, temp_intercept_col).value if formula_rows else None
        slope = sheet.cells(row, temp_slope_col).value if formula_rows else None

        probe = [num_quarters_used, forecast_value, forecast_max, forecast_min, intercept, slope]
        if not has_content(probe):
            if rows:
                break
            continue

        signature = (
            normalize_signature_value(num_quarters_used),
            normalize_signature_value(forecast_value),
            normalize_signature_value(forecast_max),
            normalize_signature_value(forecast_min),
            normalize_signature_value(intercept),
            normalize_signature_value(slope),
        )
        if previous_signature is not None and signature == previous_signature:
            break

        rows.append(
            {
                "model": labels.model,
                "ticker": labels.ticker,
                "model_period": labels.model_period,
                "model_date": labels.model_date,
                "method": "regression",
                "parameter_name": "num_quarters_used",
                "parameter_value": clean_value(num_quarters_used),
                "num_quarters_used": clean_value(num_quarters_used),
                "forecast_value": clean_value(forecast_value),
                "actual_value": clean_value(actual_value),
                "forecast_max": clean_value(forecast_max),
                "forecast_min": clean_value(forecast_min),
                "range_width": subtract_if_numeric(forecast_max, forecast_min),
                "intercept": clean_value(intercept),
                "slope": clean_value(slope),
                "source_file": source_file,
            }
        )
        previous_signature = signature

    return rows


def set_reasonable_column_widths(worksheet: Worksheet, headers: Sequence[str], max_width: int = 60) -> None:
    for col_idx, header in enumerate(headers, start=1):
        max_len = len(header)
        for row_idx in range(2, worksheet.max_row + 1):
            value = worksheet.cell(row=row_idx, column=col_idx).value
            if value is None:
                continue
            as_text = str(value)
            if len(as_text) > max_len:
                max_len = len(as_text)
        worksheet.column_dimensions[get_column_letter(col_idx)].width = min(max_len + 2, max_width)


def write_sheet(worksheet: Worksheet, headers: Sequence[str], data_rows: Sequence[Dict[str, Any]]) -> None:
    worksheet.append(list(headers))
    for row_data in data_rows:
        worksheet.append([row_data.get(header) for header in headers])

    for col_idx in range(1, len(headers) + 1):
        worksheet.cell(row=1, column=col_idx).font = Font(bold=True)

    worksheet.freeze_panes = "A2"
    worksheet.auto_filter.ref = worksheet.dimensions
    set_reasonable_column_widths(worksheet, headers)


def write_output_workbook(
    output_path: Path,
    empirical_rows: Sequence[Dict[str, Any]],
    regression_rows: Sequence[Dict[str, Any]],
) -> None:
    workbook = Workbook()
    default_sheet = workbook.active
    workbook.remove(default_sheet)

    empirical_sheet = workbook.create_sheet("empirical_candidates")
    write_sheet(empirical_sheet, EMPIRICAL_COLUMNS, empirical_rows)

    regression_sheet = workbook.create_sheet("regression_candidates")
    write_sheet(regression_sheet, REGRESSION_COLUMNS, regression_rows)

    workbook.save(output_path)


def main() -> None:
    if xw is None:
        raise ModuleNotFoundError(
            "xlwings is required to run this script. Install dependencies with "
            "'pip install xlwings openpyxl'."
        ) from _XLWINGS_IMPORT_ERROR

    if not input_dir.exists():
        raise FileNotFoundError(f"Input directory does not exist: {input_dir}")

    output_path = next_output_path(input_dir, output_dir)
    empirical_rows: List[Dict[str, Any]] = []
    regression_rows: List[Dict[str, Any]] = []
    processed_files = 0

    app = xw.App(visible=False, add_book=False)
    app.display_alerts = False
    app.screen_updating = False
    try:
        try:
            app.calculation = "manual"
        except Exception:
            pass

        for file_path in sorted(input_dir.iterdir()):
            if file_path.is_dir():
                print(f"Skipped {file_path.name}: directory")
                continue
            if file_path.name.startswith("~"):
                print(f"Skipped {file_path.name}: temporary file")
                continue
            if file_path.suffix.lower() != ".xlsx":
                print(f"Skipped {file_path.name}: not an .xlsx file")
                continue

            print(f"Processing {file_path.name}")
            workbook: Optional[xw.Book] = None
            try:
                labels = parse_file_labels(file_path)
                workbook = app.books.open(str(file_path), update_links=False)

                empirical_rows.extend(extract_empirical_rows(workbook, labels, file_path.name))
                regression_rows.extend(extract_regression_rows(workbook, labels, file_path.name))
                processed_files += 1
            except Exception as exc:
                print(f"Skipped {file_path.name}: processing error ({exc})")
            finally:
                if workbook is not None:
                    close_workbook_safely(workbook)

        write_output_workbook(output_path, empirical_rows, regression_rows)
    finally:
        app.quit()

    print(f"Output path: {output_path}")
    print(f"Files processed: {processed_files}")
    print(f"Empirical rows: {len(empirical_rows)}")
    print(f"Regression rows: {len(regression_rows)}")


if __name__ == "__main__":
    main()
