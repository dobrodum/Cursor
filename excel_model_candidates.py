#!/usr/bin/env python3
"""
Extract empirical and regression candidates from model workbooks.

The script scans all .xlsx files in input_dir, processes each source workbook once,
and writes one combined output workbook with:
  - empirical_candidates
  - regression_candidates
"""

from __future__ import annotations

from datetime import date, datetime
from pathlib import Path
import re
from typing import Any, Dict, Iterable, List, Optional, Sequence, Tuple

import xlwings as xw
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter


# ---------------------------------------------------------------------------
# Configure paths here.
# ---------------------------------------------------------------------------
input_dir = "/path/to/input"
output_dir = "/path/to/output"


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


PERIOD_DAY = {"Early": 5, "Mid": 15, "Late": 25}


def to_2d(values: Any) -> List[List[Any]]:
    """Normalize xlwings return values into 2D lists."""
    if values is None:
        return []
    if isinstance(values, list):
        if not values:
            return []
        if isinstance(values[0], list):
            return values
        return [values]
    return [[values]]


def normalize_text(value: Any) -> str:
    if value is None:
        return ""
    return str(value).strip()


def normalize_header(value: Any) -> str:
    text = normalize_text(value).lower()
    text = re.sub(r"[^a-z0-9]+", "_", text)
    return text.strip("_")


def is_blank(value: Any) -> bool:
    return value is None or (isinstance(value, str) and value.strip() == "")


def numeric(value: Any) -> Optional[float]:
    if is_blank(value):
        return None
    try:
        return float(value)
    except (TypeError, ValueError):
        return None


def calc_range_width(max_value: Any, min_value: Any) -> Optional[float]:
    max_num = numeric(max_value)
    min_num = numeric(min_value)
    if max_num is None or min_num is None:
        return None
    return max_num - min_num


def parse_month(month_token: str) -> Tuple[Optional[int], Optional[str]]:
    token = month_token.strip()
    if not token:
        return None, None

    attempts = [token, token[:3]]
    for candidate in attempts:
        for fmt in ("%b", "%B"):
            try:
                parsed = datetime.strptime(candidate.title(), fmt)
                month_num = parsed.month
                month_abbrev = date(2000, month_num, 1).strftime("%b")
                return month_num, month_abbrev
            except ValueError:
                continue
    return None, None


def parse_file_label(file_path: Path) -> Optional[Dict[str, str]]:
    """
    Parse filename like:
      MedMiner_Model - AORT - MidJan2026_Send.xlsx
    """
    stem = file_path.stem
    parts = [part.strip() for part in stem.split(" - ")]
    if len(parts) < 3:
        return None

    ticker = parts[1]
    period_chunk = parts[2]
    period_match = re.search(r"(Early|Mid|Late)\s*([A-Za-z]+)\s*(\d{4})", period_chunk, re.IGNORECASE)
    if not period_match:
        return None

    period_label = period_match.group(1).title()
    month_token = period_match.group(2)
    year_text = period_match.group(3)

    month_num, month_abbrev = parse_month(month_token)
    if month_num is None or month_abbrev is None:
        return None

    year_num = int(year_text)
    model_period = f"{period_label}{month_abbrev}_{year_num}"
    model_date = date(year_num, month_num, PERIOD_DAY[period_label]).isoformat()
    model = f"{ticker}_{model_period}"

    return {
        "model": model,
        "ticker": ticker,
        "model_period": model_period,
        "model_date": model_date,
    }


def discover_source_files(input_path: Path) -> Tuple[List[Path], List[str]]:
    eligible_files: List[Path] = []
    skipped: List[str] = []

    if not input_path.exists() or not input_path.is_dir():
        skipped.append(f"Skipping input directory: {input_path} (missing or not a directory)")
        return eligible_files, skipped

    for file_path in sorted(input_path.iterdir()):
        if not file_path.is_file():
            continue
        if file_path.name.startswith("~"):
            skipped.append(f"Skipping file: {file_path.name} (temporary workbook)")
            continue
        if file_path.suffix.lower() != ".xlsx":
            skipped.append(f"Skipping file: {file_path.name} (not .xlsx)")
            continue
        eligible_files.append(file_path)

    return eligible_files, skipped


def get_unique_output_path(input_path: Path, output_path: Path) -> Path:
    base_name = f"{input_path.name}_PARAM"
    candidate = output_path / f"{base_name}.xlsx"
    if not candidate.exists():
        return candidate

    suffix = 1
    while True:
        candidate = output_path / f"{base_name}.{suffix}.xlsx"
        if not candidate.exists():
            return candidate
        suffix += 1


def get_sheet_or_none(workbook: xw.Book, sheet_name: str) -> Optional[xw.Sheet]:
    try:
        return workbook.sheets[sheet_name]
    except Exception:
        return None


def find_anchor_cell(sheet: xw.Sheet, label: str = "max") -> Optional[Tuple[int, int]]:
    used = sheet.used_range
    values = to_2d(used.value)
    if not values:
        return None

    start_row = used.row
    start_col = used.column
    target = label.strip().lower()

    for row_offset, row_values in enumerate(values):
        for col_offset, cell_value in enumerate(row_values):
            if isinstance(cell_value, str) and cell_value.strip().lower() == target:
                return start_row + row_offset, start_col + col_offset
    return None


def header_map_near_anchor(sheet: xw.Sheet, anchor_row: int, anchor_col: int, radius: int = 20) -> Dict[str, int]:
    start_col = max(1, anchor_col - radius)
    end_col = anchor_col + radius
    row_values = to_2d(sheet.range((anchor_row, start_col), (anchor_row, end_col)).value)
    if not row_values:
        return {}

    mapping: Dict[str, int] = {}
    for idx, value in enumerate(row_values[0]):
        header = normalize_header(value)
        if header:
            mapping[header] = start_col + idx
    return mapping


def pick_column(mapping: Dict[str, int], token_groups: Sequence[Sequence[str]], fallback: int) -> int:
    for tokens in token_groups:
        for header_text, column_number in mapping.items():
            if all(token in header_text for token in tokens):
                return column_number
    return fallback


def set_r1c1_formula2(cell: xw.Range, formula: str) -> None:
    """
    Try Formula2 R1C1 first, then fallback.
    This keeps formulas in R1C1 style while supporting more environments.
    """
    try:
        cell.api.Formula2R1C1 = formula
        return
    except Exception:
        pass

    try:
        cell.formula2 = formula
        return
    except Exception:
        pass

    cell.formula = formula


def safe_close_workbook(workbook: xw.Book) -> None:
    try:
        workbook.close(save=False)
        return
    except TypeError:
        pass
    except Exception:
        pass

    try:
        workbook.close(SaveChanges=False)
        return
    except Exception:
        pass

    try:
        workbook.api.Close(SaveChanges=False)
        return
    except Exception:
        pass

    try:
        workbook.close()
    except Exception:
        pass


def build_empirical_row(
    metadata: Dict[str, str],
    source_file: str,
    *,
    num_quarters_used: Any,
    last_quarter_used: Any,
    forecast_value: Any,
    actual_value: Any,
    forecast_max: Any,
    forecast_min: Any,
    avg_penetration_pct: Any,
    quarterly_sales: Any,
    reported_sales: Any,
    growth_rate_pct: Any,
    sales_captured_in_db_pct: Any,
) -> Dict[str, Any]:
    return {
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
        "range_width": calc_range_width(forecast_max, forecast_min),
        "avg_penetration_pct": avg_penetration_pct,
        "quarterly_sales": quarterly_sales,
        "reported_sales": reported_sales,
        "growth_rate_pct": growth_rate_pct,
        "sales_captured_in_db_pct": sales_captured_in_db_pct,
        "source_file": source_file,
    }


def extract_empirical_rows(
    workbook: xw.Book, metadata: Dict[str, str], source_file: str, n_quarters: int = 10
) -> Tuple[List[Dict[str, Any]], Optional[str]]:
    sheet = get_sheet_or_none(workbook, "Empirical Model")
    if sheet is None:
        return [], "missing sheet 'Empirical Model'"

    anchor = find_anchor_cell(sheet, "max")
    if anchor is None:
        return [], "missing 'max' anchor in 'Empirical Model'"

    anchor_row, anchor_col = anchor
    headers = header_map_near_anchor(sheet, anchor_row, anchor_col)

    num_quarters_col = pick_column(headers, [("num", "quarter"), ("n", "quarter")], anchor_col - 6)
    last_quarter_col = pick_column(headers, [("last", "quarter")], anchor_col - 5)
    forecast_col = pick_column(
        headers,
        [
            ("estimated", "total", "sold"),
            ("forecast", "value"),
            ("forecast", "total"),
        ],
        anchor_col - 4,
    )
    actual_col = pick_column(headers, [("actual",), ("reported", "sales")], anchor_col - 3)
    min_col = pick_column(headers, [("min",)], anchor_col + 1)
    avg_pen_col = pick_column(headers, [("avg", "penetration"), ("average", "penetration")], anchor_col - 2)
    quarterly_sales_col = pick_column(headers, [("quarterly", "sales")], anchor_col - 9)
    reported_sales_col = pick_column(headers, [("reported", "sales")], actual_col)
    growth_col = pick_column(headers, [("growth", "rate")], anchor_col - 7)
    captured_col = pick_column(
        headers,
        [
            ("captured", "db"),
            ("captured", "database"),
            ("sales", "captured"),
        ],
        anchor_col - 1,
    )

    data_end_row = anchor_row - 1
    formulas_written = False

    for n in range(1, n_quarters + 1):
        data_start_row = data_end_row - n + 1
        if data_start_row < 1:
            break
        target_row = anchor_row + n
        formula = f'=IFERROR(AVERAGE(R{data_start_row}C{captured_col}:R{data_end_row}C{captured_col}),"")'
        set_r1c1_formula2(sheet.range((target_row, avg_pen_col)), formula)
        formulas_written = True

    if formulas_written:
        workbook.app.calculate()

    rows: List[Dict[str, Any]] = []
    for n in range(1, n_quarters + 1):
        row = anchor_row + n
        num_quarters_used = sheet.range((row, num_quarters_col)).value
        if is_blank(num_quarters_used):
            num_quarters_used = n

        last_quarter_used = sheet.range((row, last_quarter_col)).value
        forecast_value = sheet.range((row, forecast_col)).value
        actual_value = sheet.range((row, actual_col)).value
        forecast_max = sheet.range((row, anchor_col)).value
        forecast_min = sheet.range((row, min_col)).value
        avg_penetration_pct = sheet.range((row, avg_pen_col)).value
        quarterly_sales = sheet.range((row, quarterly_sales_col)).value
        reported_sales = sheet.range((row, reported_sales_col)).value
        growth_rate_pct = sheet.range((row, growth_col)).value
        sales_captured_in_db_pct = sheet.range((row, captured_col)).value

        if all(
            is_blank(value)
            for value in (
                forecast_value,
                forecast_max,
                forecast_min,
                avg_penetration_pct,
            )
        ):
            continue

        rows.append(
            build_empirical_row(
                metadata,
                source_file,
                num_quarters_used=num_quarters_used,
                last_quarter_used=last_quarter_used,
                forecast_value=forecast_value,
                actual_value=actual_value,
                forecast_max=forecast_max,
                forecast_min=forecast_min,
                avg_penetration_pct=avg_penetration_pct,
                quarterly_sales=quarterly_sales,
                reported_sales=reported_sales,
                growth_rate_pct=growth_rate_pct,
                sales_captured_in_db_pct=sales_captured_in_db_pct,
            )
        )

    return rows, None


def build_regression_row(
    metadata: Dict[str, str],
    source_file: str,
    *,
    num_quarters_used: Any,
    forecast_value: Any,
    forecast_max: Any,
    forecast_min: Any,
    intercept: Any,
    slope: Any,
    actual_value: Any = None,
) -> Dict[str, Any]:
    return {
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
        "range_width": calc_range_width(forecast_max, forecast_min),
        "intercept": intercept,
        "slope": slope,
        "source_file": source_file,
    }


def row_signature_for_dedup(values: Iterable[Any]) -> Tuple[Any, ...]:
    normalized: List[Any] = []
    for value in values:
        num = numeric(value)
        if num is None:
            normalized.append(normalize_text(value) or None)
        else:
            normalized.append(round(num, 10))
    return tuple(normalized)


def extract_regression_rows(
    workbook: xw.Book, metadata: Dict[str, str], source_file: str, n_quarters: int = 10
) -> Tuple[List[Dict[str, Any]], Optional[str]]:
    sheet = get_sheet_or_none(workbook, "Regression Model")
    if sheet is None:
        return [], "missing sheet 'Regression Model'"

    anchor = find_anchor_cell(sheet, "max")
    if anchor is None:
        return [], "missing 'max' anchor in 'Regression Model'"

    anchor_row, anchor_col = anchor
    headers = header_map_near_anchor(sheet, anchor_row, anchor_col)

    num_quarters_col = pick_column(headers, [("num", "quarter"), ("n", "quarter")], anchor_col - 8)
    forecast_col = pick_column(
        headers,
        [
            ("tot", "fcst", "w", "sa"),
            ("tot", "forecast", "without", "sa"),
            ("forecast", "total"),
        ],
        anchor_col - 1,
    )
    min_col = pick_column(headers, [("min",)], anchor_col + 1)
    actual_col = pick_column(headers, [("actual",), ("reported", "sales")], -1)

    # Required by specification.
    y_col = anchor_col - 7
    x_col = anchor_col - 11

    used_last_row = sheet.used_range.last_cell.row
    used_last_col = sheet.used_range.last_cell.column
    helper_row_start = used_last_row + 2
    helper_intercept_col = used_last_col + 2
    helper_slope_col = helper_intercept_col + 1

    data_end_row = anchor_row - 1
    helper_rows: List[Tuple[int, int]] = []

    for n in range(1, n_quarters + 1):
        data_start_row = data_end_row - n + 1
        if data_start_row < 1:
            break

        helper_row = helper_row_start + n - 1
        intercept_formula = (
            f'=IFERROR(INTERCEPT(R{data_start_row}C{y_col}:R{data_end_row}C{y_col},'
            f'R{data_start_row}C{x_col}:R{data_end_row}C{x_col}),"")'
        )
        slope_formula = (
            f'=IFERROR(SLOPE(R{data_start_row}C{y_col}:R{data_end_row}C{y_col},'
            f'R{data_start_row}C{x_col}:R{data_end_row}C{x_col}),"")'
        )
        set_r1c1_formula2(sheet.range((helper_row, helper_intercept_col)), intercept_formula)
        set_r1c1_formula2(sheet.range((helper_row, helper_slope_col)), slope_formula)
        helper_rows.append((n, helper_row))

    if helper_rows:
        workbook.app.calculate()

    rows: List[Dict[str, Any]] = []
    prev_signature: Optional[Tuple[Any, ...]] = None

    for n, helper_row in helper_rows:
        candidate_row = anchor_row + n
        num_quarters_used = sheet.range((candidate_row, num_quarters_col)).value
        if is_blank(num_quarters_used):
            num_quarters_used = n

        forecast_value = sheet.range((candidate_row, forecast_col)).value
        forecast_max = sheet.range((candidate_row, anchor_col)).value
        forecast_min = sheet.range((candidate_row, min_col)).value
        actual_value = sheet.range((candidate_row, actual_col)).value if actual_col > 0 else None
        intercept_value = sheet.range((helper_row, helper_intercept_col)).value
        slope_value = sheet.range((helper_row, helper_slope_col)).value

        if all(is_blank(value) for value in (forecast_value, forecast_max, forecast_min, intercept_value, slope_value)):
            continue

        signature = row_signature_for_dedup(
            (
                num_quarters_used,
                forecast_value,
                forecast_max,
                forecast_min,
                intercept_value,
                slope_value,
            )
        )
        if prev_signature is not None and signature == prev_signature:
            continue
        prev_signature = signature

        rows.append(
            build_regression_row(
                metadata,
                source_file,
                num_quarters_used=num_quarters_used,
                forecast_value=forecast_value,
                forecast_max=forecast_max,
                forecast_min=forecast_min,
                intercept=intercept_value,
                slope=slope_value,
                actual_value=actual_value,
            )
        )

    return rows, None


def set_sheet_formatting(worksheet: Any, headers: List[str], rows: List[Dict[str, Any]]) -> None:
    worksheet.freeze_panes = "A2"
    worksheet.auto_filter.ref = worksheet.dimensions
    for cell in worksheet[1]:
        cell.font = Font(bold=True)

    for index, header in enumerate(headers, start=1):
        max_len = len(header)
        for row in rows:
            value = row.get(header)
            text = "" if value is None else str(value)
            if len(text) > max_len:
                max_len = len(text)
        worksheet.column_dimensions[get_column_letter(index)].width = min(max_len + 2, 45)


def write_output_workbook(
    empirical_rows: List[Dict[str, Any]],
    regression_rows: List[Dict[str, Any]],
    output_path: Path,
) -> None:
    workbook = Workbook()
    empirical_sheet = workbook.active
    empirical_sheet.title = "empirical_candidates"
    regression_sheet = workbook.create_sheet("regression_candidates")

    empirical_sheet.append(EMPIRICAL_COLUMNS)
    for row in empirical_rows:
        empirical_sheet.append([row.get(column) for column in EMPIRICAL_COLUMNS])

    regression_sheet.append(REGRESSION_COLUMNS)
    for row in regression_rows:
        regression_sheet.append([row.get(column) for column in REGRESSION_COLUMNS])

    set_sheet_formatting(empirical_sheet, EMPIRICAL_COLUMNS, empirical_rows)
    set_sheet_formatting(regression_sheet, REGRESSION_COLUMNS, regression_rows)

    output_path.parent.mkdir(parents=True, exist_ok=True)
    workbook.save(output_path)


def run() -> None:
    input_path = Path(input_dir).expanduser().resolve()
    output_path_root = Path(output_dir).expanduser().resolve()
    output_file = get_unique_output_path(input_path, output_path_root)

    source_files, skipped_messages = discover_source_files(input_path)
    for message in skipped_messages:
        print(message)

    empirical_rows: List[Dict[str, Any]] = []
    regression_rows: List[Dict[str, Any]] = []
    processed_files = 0

    app = xw.App(visible=False, add_book=False)
    app.display_alerts = False
    app.screen_updating = False

    try:
        for file_path in source_files:
            metadata = parse_file_label(file_path)
            if metadata is None:
                print(f"Skipping file: {file_path.name} (filename label parsing failed)")
                continue

            print(f"Processing file: {file_path.name}")
            workbook = None
            try:
                workbook = app.books.open(str(file_path), update_links=False)
            except Exception as exc:
                print(f"Skipping file: {file_path.name} (open failed: {exc})")
                continue

            try:
                file_empirical_rows, empirical_error = extract_empirical_rows(workbook, metadata, file_path.name)
                if empirical_error:
                    print(f"Skipping empirical in {file_path.name}: {empirical_error}")
                empirical_rows.extend(file_empirical_rows)

                file_regression_rows, regression_error = extract_regression_rows(workbook, metadata, file_path.name)
                if regression_error:
                    print(f"Skipping regression in {file_path.name}: {regression_error}")
                regression_rows.extend(file_regression_rows)

                processed_files += 1
            except Exception as exc:
                print(f"Skipping file: {file_path.name} (processing error: {exc})")
            finally:
                if workbook is not None:
                    safe_close_workbook(workbook)
    finally:
        app.quit()

    write_output_workbook(empirical_rows, regression_rows, output_file)

    print(f"Output path: {output_file}")
    print(f"Files processed: {processed_files}")
    print(f"Empirical rows: {len(empirical_rows)}")
    print(f"Regression rows: {len(regression_rows)}")


if __name__ == "__main__":
    run()
