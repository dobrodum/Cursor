#!/usr/bin/env python3
"""Extract empirical and regression model candidates from .xlsx workbooks."""

from __future__ import annotations

from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path
import math
import re
from typing import Any, Dict, Iterable, List, Optional, Sequence, Tuple

import xlwings as xw
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter


# =========================
# Required user configuration
# =========================
input_dir = Path("input")
output_dir = Path("output")


N_QUARTERS = 10
EMPIRICAL_SHEET_NAME = "Empirical Model"
REGRESSION_SHEET_NAME = "Regression Model"

EMPIRICAL_OUTPUT_COLUMNS = [
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

REGRESSION_OUTPUT_COLUMNS = [
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

# Anchor-based defaults. If header names are found in the anchor row,
# header matching overrides these defaults.
EMPIRICAL_COLUMN_OFFSETS = {
    "num_quarters_used": -8,
    "last_quarter_used": -7,
    "forecast_value": -6,
    "actual_value": -5,
    "avg_penetration_pct": -4,
    "quarterly_sales": -3,
    "reported_sales": -2,
    "growth_rate_pct": -1,
    "forecast_max": 0,
    "forecast_min": 1,
    "sales_captured_in_db_pct": 2,
}

REGRESSION_COLUMN_OFFSETS = {
    "num_quarters_used": -10,
    "forecast_value": -4,
    "actual_value": -3,
    "forecast_max": 0,
    "forecast_min": 1,
}

EMPIRICAL_HEADER_KEYWORDS = {
    "num_quarters_used": ("num quarter", "quarters used", "quarter count"),
    "last_quarter_used": ("last quarter",),
    "forecast_value": ("estimated total sold", "forecast value", "tot fcst"),
    "actual_value": ("actual", "reported sales"),
    "avg_penetration_pct": ("avg penetration", "average penetration"),
    "quarterly_sales": ("quarterly sales",),
    "reported_sales": ("reported sales", "reported sale"),
    "growth_rate_pct": ("growth rate",),
    "forecast_max": ("max",),
    "forecast_min": ("min",),
    "sales_captured_in_db_pct": ("captured in db", "sales captured", "db pct"),
}

REGRESSION_HEADER_KEYWORDS = {
    "num_quarters_used": ("num quarter", "quarters used", "quarter count"),
    "forecast_value": ("tot fcst w/o sa", "tot fcst wo sa", "forecast", "forecast value"),
    "actual_value": ("actual", "reported sales"),
    "forecast_max": ("max",),
    "forecast_min": ("min",),
}

FILE_OUTPUT_PATTERN = re.compile(r"_param(?:\.\d+)?\.xlsx$", flags=re.IGNORECASE)


@dataclass(frozen=True)
class FileLabels:
    model: str
    ticker: str
    model_period: str
    model_date: str


def normalize_text(value: Any) -> str:
    if value is None:
        return ""
    text = str(value).strip().lower()
    text = re.sub(r"[^a-z0-9]+", " ", text)
    return re.sub(r"\s+", " ", text).strip()


def month_token_to_number(month_token: str) -> int:
    month_short = month_token[:3].title()
    return datetime.strptime(month_short, "%b").month


def parse_file_labels(file_name: str) -> FileLabels:
    stem = Path(file_name).stem
    parts = [part.strip() for part in stem.split(" - ") if part.strip()]
    ticker = parts[1].upper() if len(parts) >= 2 else ""

    period_source = parts[2] if len(parts) >= 3 else stem
    period_source = period_source.split("_")[0]
    period_match = re.search(
        r"(Early|Mid|Late)\s*([A-Za-z]{3,9})\s*[_-]?(\d{4})",
        period_source,
        flags=re.IGNORECASE,
    )
    if not period_match:
        period_match = re.search(
            r"(Early|Mid|Late)\s*([A-Za-z]{3,9})\s*[_-]?(\d{4})",
            stem,
            flags=re.IGNORECASE,
        )
    if not period_match:
        raise ValueError("could not parse model period (Early/Mid/Late + month + year)")

    timing_raw, month_raw, year_raw = period_match.groups()
    timing = timing_raw.title()
    month_short = month_raw[:3].title()
    year = int(year_raw)
    day_map = {"Early": 5, "Mid": 15, "Late": 25}
    month_number = month_token_to_number(month_short)
    model_period = f"{timing}{month_short}_{year}"
    model_date = date(year, month_number, day_map[timing]).isoformat()

    if not ticker:
        ticker_match = re.search(r"\b([A-Z]{2,6})\b", stem)
        ticker = ticker_match.group(1) if ticker_match else "UNKNOWN"

    model = f"{ticker}_{model_period}"
    return FileLabels(model=model, ticker=ticker, model_period=model_period, model_date=model_date)


def is_blank(value: Any) -> bool:
    if value is None:
        return True
    if isinstance(value, str) and value.strip() == "":
        return True
    return False


def to_float(value: Any) -> Optional[float]:
    if value is None:
        return None
    if isinstance(value, bool):
        return float(value)
    if isinstance(value, (int, float)):
        if isinstance(value, float) and math.isnan(value):
            return None
        return float(value)
    if isinstance(value, str):
        cleaned = value.strip().replace(",", "")
        if cleaned.endswith("%"):
            cleaned = cleaned[:-1].strip()
            if cleaned == "":
                return None
            try:
                return float(cleaned) / 100.0
            except ValueError:
                return None
        if cleaned == "":
            return None
        try:
            return float(cleaned)
        except ValueError:
            return None
    return None


def to_int(value: Any) -> Optional[int]:
    as_float = to_float(value)
    if as_float is None:
        return None
    return int(round(as_float))


def subtract_numeric(left: Any, right: Any) -> Optional[float]:
    left_num = to_float(left)
    right_num = to_float(right)
    if left_num is None or right_num is None:
        return None
    return left_num - right_num


def values_equal(left: Any, right: Any, tolerance: float = 1e-9) -> bool:
    left_num = to_float(left)
    right_num = to_float(right)
    if left_num is not None and right_num is not None:
        return abs(left_num - right_num) <= tolerance
    return str(left) == str(right)


def get_output_path(source_dir: Path, target_dir: Path) -> Path:
    target_dir.mkdir(parents=True, exist_ok=True)
    base_name = f"{source_dir.name}_PARAM"
    candidate = target_dir / f"{base_name}.xlsx"
    index = 1
    while candidate.exists():
        candidate = target_dir / f"{base_name}.{index}.xlsx"
        index += 1
    return candidate


def safe_close_workbook(workbook: Any) -> None:
    try:
        workbook.close(save=False)
        return
    except TypeError:
        pass
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


def get_sheet(workbook: Any, target_name: str) -> Optional[Any]:
    lookup = {sheet.name.strip().lower(): sheet for sheet in workbook.sheets}
    return lookup.get(target_name.strip().lower())


def get_used_matrix(sheet: Any) -> Tuple[int, int, List[List[Any]], int, int]:
    used = sheet.used_range
    start_row = used.row
    start_col = used.column
    end_row = used.last_cell.row
    end_col = used.last_cell.column
    raw_values = used.value

    if raw_values is None:
        return start_row, start_col, [], end_row, end_col
    if not isinstance(raw_values, list):
        matrix = [[raw_values]]
    elif raw_values and not isinstance(raw_values[0], list):
        matrix = [raw_values]
    else:
        matrix = raw_values
    return start_row, start_col, matrix, end_row, end_col


def find_max_anchor(matrix: Sequence[Sequence[Any]], start_row: int, start_col: int) -> Optional[Tuple[int, int]]:
    for row_index, row_values in enumerate(matrix):
        for col_index, value in enumerate(row_values):
            if normalize_text(value) == "max":
                return start_row + row_index, start_col + col_index
    return None


def map_headers_from_anchor_row(
    matrix: Sequence[Sequence[Any]],
    start_row: int,
    start_col: int,
    anchor_row: int,
    header_keywords: Dict[str, Sequence[str]],
) -> Dict[str, int]:
    matrix_row_index = anchor_row - start_row
    if matrix_row_index < 0 or matrix_row_index >= len(matrix):
        return {}

    header_row_values = matrix[matrix_row_index]
    mapped: Dict[str, int] = {}
    for field_name, keywords in header_keywords.items():
        for offset, cell_value in enumerate(header_row_values):
            normalized = normalize_text(cell_value)
            if not normalized:
                continue
            if any(keyword in normalized for keyword in keywords):
                mapped[field_name] = start_col + offset
                break
    return mapped


def build_column_map(
    anchor_col: int,
    default_offsets: Dict[str, int],
    header_mapped_columns: Dict[str, int],
) -> Dict[str, int]:
    column_map = {field: max(1, anchor_col + offset) for field, offset in default_offsets.items()}
    column_map.update(header_mapped_columns)
    return column_map


def set_formula2_r1c1(cell: Any, formula_r1c1: str) -> None:
    try:
        cell.formula2 = formula_r1c1
        return
    except Exception:
        pass
    cell.api.Formula2R1C1 = formula_r1c1


def get_cell_value(sheet: Any, row: int, col: int) -> Any:
    if row < 1 or col < 1:
        return None
    return sheet.cells(row, col).value


def extract_empirical_rows(sheet: Any, labels: FileLabels, source_file: str) -> List[Dict[str, Any]]:
    start_row, start_col, matrix, _last_row, last_col = get_used_matrix(sheet)
    if not matrix:
        return []

    anchor = find_max_anchor(matrix, start_row, start_col)
    if anchor is None:
        return []
    anchor_row, anchor_col = anchor

    header_map = map_headers_from_anchor_row(
        matrix=matrix,
        start_row=start_row,
        start_col=start_col,
        anchor_row=anchor_row,
        header_keywords=EMPIRICAL_HEADER_KEYWORDS,
    )
    column_map = build_column_map(anchor_col, EMPIRICAL_COLUMN_OFFSETS, header_map)

    helper_col = max(last_col, anchor_col) + 2
    data_start_row = anchor_row + 1
    data_end_row = data_start_row + N_QUARTERS - 1

    formula_rows: List[int] = []
    for row in range(data_start_row, data_end_row + 1):
        avg_col = column_map.get("avg_penetration_pct")
        quarterly_col = column_map.get("quarterly_sales")
        reported_col = column_map.get("reported_sales")
        if avg_col and quarterly_col and reported_col:
            formula = (
                f"=IFERROR(R{row}C{quarterly_col}/R{row}C{reported_col},"
                f"R{row}C{avg_col})"
            )
        elif quarterly_col and reported_col:
            formula = f"=IFERROR(R{row}C{quarterly_col}/R{row}C{reported_col},\"\")"
        elif avg_col:
            formula = f"=R{row}C{avg_col}"
        else:
            formula = "=\"\""
        set_formula2_r1c1(sheet.cells(row, helper_col), formula)
        formula_rows.append(row)

    # Calculate once after all empirical formulas are written.
    sheet.book.app.calculate()

    extracted_rows: List[Dict[str, Any]] = []
    try:
        for row_offset, row in enumerate(range(data_start_row, data_end_row + 1)):
            num_quarters_used = get_cell_value(sheet, row, column_map["num_quarters_used"])
            if is_blank(num_quarters_used):
                num_quarters_used = N_QUARTERS - row_offset

            last_quarter_used = get_cell_value(sheet, row, column_map["last_quarter_used"])
            forecast_value = get_cell_value(sheet, row, column_map["forecast_value"])
            actual_value = get_cell_value(sheet, row, column_map["actual_value"])
            forecast_max = get_cell_value(sheet, row, column_map["forecast_max"])
            forecast_min = get_cell_value(sheet, row, column_map["forecast_min"])
            avg_penetration_pct = get_cell_value(sheet, row, helper_col)
            quarterly_sales = get_cell_value(sheet, row, column_map["quarterly_sales"])
            reported_sales = get_cell_value(sheet, row, column_map["reported_sales"])
            growth_rate_pct = get_cell_value(sheet, row, column_map["growth_rate_pct"])
            sales_captured_in_db_pct = get_cell_value(
                sheet, row, column_map["sales_captured_in_db_pct"]
            )

            row_values_to_check = (
                num_quarters_used,
                last_quarter_used,
                forecast_value,
                actual_value,
                forecast_max,
                forecast_min,
                avg_penetration_pct,
                quarterly_sales,
                reported_sales,
            )
            if all(is_blank(value) for value in row_values_to_check):
                continue

            extracted_rows.append(
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
                    "actual_value": actual_value,
                    "forecast_max": forecast_max,
                    "forecast_min": forecast_min,
                    "range_width": subtract_numeric(forecast_max, forecast_min),
                    "avg_penetration_pct": avg_penetration_pct,
                    "quarterly_sales": quarterly_sales,
                    "reported_sales": reported_sales,
                    "growth_rate_pct": growth_rate_pct,
                    "sales_captured_in_db_pct": sales_captured_in_db_pct,
                    "source_file": source_file,
                }
            )
    finally:
        if formula_rows:
            sheet.range((formula_rows[0], helper_col), (formula_rows[-1], helper_col)).clear_contents()

    return extracted_rows


def extract_regression_rows(sheet: Any, labels: FileLabels, source_file: str) -> List[Dict[str, Any]]:
    start_row, start_col, matrix, last_row, last_col = get_used_matrix(sheet)
    if not matrix:
        return []

    anchor = find_max_anchor(matrix, start_row, start_col)
    if anchor is None:
        return []
    anchor_row, anchor_col = anchor

    header_map = map_headers_from_anchor_row(
        matrix=matrix,
        start_row=start_row,
        start_col=start_col,
        anchor_row=anchor_row,
        header_keywords=REGRESSION_HEADER_KEYWORDS,
    )
    column_map = build_column_map(anchor_col, REGRESSION_COLUMN_OFFSETS, header_map)

    y_col = anchor_col - 7
    x_col = anchor_col - 11
    intercept_col = max(last_col, anchor_col) + 2
    slope_col = intercept_col + 1

    data_start_row = anchor_row + 1
    data_end_row = data_start_row + N_QUARTERS - 1
    formula_rows: List[int] = []

    for row_offset, row in enumerate(range(data_start_row, data_end_row + 1)):
        num_quarters = get_cell_value(sheet, row, column_map["num_quarters_used"])
        if is_blank(num_quarters):
            num_quarters = N_QUARTERS - row_offset
        n = max(2, to_int(num_quarters) or 2)

        y_start = row
        y_end = min(last_row, y_start + n - 1)
        if y_end - y_start + 1 < 2:
            intercept_formula = "=\"\""
            slope_formula = "=\"\""
        else:
            intercept_formula = (
                f"=IFERROR(INTERCEPT(R{y_start}C{y_col}:R{y_end}C{y_col},"
                f"R{y_start}C{x_col}:R{y_end}C{x_col}),\"\")"
            )
            slope_formula = (
                f"=IFERROR(SLOPE(R{y_start}C{y_col}:R{y_end}C{y_col},"
                f"R{y_start}C{x_col}:R{y_end}C{x_col}),\"\")"
            )

        set_formula2_r1c1(sheet.cells(row, intercept_col), intercept_formula)
        set_formula2_r1c1(sheet.cells(row, slope_col), slope_formula)
        formula_rows.append(row)

    # Calculate once after all regression formulas are written.
    sheet.book.app.calculate()

    extracted_rows: List[Dict[str, Any]] = []
    try:
        for row_offset, row in enumerate(range(data_start_row, data_end_row + 1)):
            num_quarters_used = get_cell_value(sheet, row, column_map["num_quarters_used"])
            if is_blank(num_quarters_used):
                num_quarters_used = N_QUARTERS - row_offset

            forecast_value = get_cell_value(sheet, row, column_map["forecast_value"])
            actual_value = get_cell_value(sheet, row, column_map["actual_value"])
            forecast_max = get_cell_value(sheet, row, column_map["forecast_max"])
            forecast_min = get_cell_value(sheet, row, column_map["forecast_min"])
            intercept = get_cell_value(sheet, row, intercept_col)
            slope = get_cell_value(sheet, row, slope_col)

            row_values_to_check = (
                num_quarters_used,
                forecast_value,
                forecast_max,
                forecast_min,
                intercept,
                slope,
            )
            if all(is_blank(value) for value in row_values_to_check):
                continue

            candidate_row = {
                "model": labels.model,
                "ticker": labels.ticker,
                "model_period": labels.model_period,
                "model_date": labels.model_date,
                "method": "regression",
                "parameter_name": "num_quarters_used",
                "parameter_value": num_quarters_used,
                "num_quarters_used": num_quarters_used,
                "forecast_value": forecast_value,
                "actual_value": actual_value if not is_blank(actual_value) else "",
                "forecast_max": forecast_max,
                "forecast_min": forecast_min,
                "range_width": subtract_numeric(forecast_max, forecast_min),
                "intercept": intercept,
                "slope": slope,
                "source_file": source_file,
            }

            if extracted_rows:
                previous_row = extracted_rows[-1]
                duplicate = (
                    values_equal(previous_row["num_quarters_used"], candidate_row["num_quarters_used"])
                    and values_equal(previous_row["forecast_value"], candidate_row["forecast_value"])
                    and values_equal(previous_row["forecast_max"], candidate_row["forecast_max"])
                    and values_equal(previous_row["forecast_min"], candidate_row["forecast_min"])
                    and values_equal(previous_row["intercept"], candidate_row["intercept"])
                    and values_equal(previous_row["slope"], candidate_row["slope"])
                )
                if duplicate:
                    continue

            extracted_rows.append(candidate_row)
    finally:
        if formula_rows:
            sheet.range((formula_rows[0], intercept_col), (formula_rows[-1], slope_col)).clear_contents()

    return extracted_rows


def write_rows(worksheet: Any, columns: Sequence[str], rows: Iterable[Dict[str, Any]]) -> None:
    worksheet.append(list(columns))
    for row in rows:
        worksheet.append([row.get(column, "") for column in columns])

    for cell in worksheet[1]:
        cell.font = Font(bold=True)
    worksheet.freeze_panes = "A2"
    worksheet.auto_filter.ref = worksheet.dimensions

    for col_index, column_name in enumerate(columns, start=1):
        max_length = len(column_name)
        for cell in worksheet.iter_cols(
            min_col=col_index,
            max_col=col_index,
            min_row=2,
            max_row=worksheet.max_row,
            values_only=True,
        ):
            for value in cell:
                if value is None:
                    continue
                max_length = max(max_length, len(str(value)))
        worksheet.column_dimensions[get_column_letter(col_index)].width = min(max(10, max_length + 2), 42)


def should_skip_file(file_path: Path) -> Optional[str]:
    if not file_path.is_file():
        return "not a file"
    if file_path.name.startswith("~"):
        return "temporary file"
    if file_path.suffix.lower() != ".xlsx":
        return "not an .xlsx file"
    if FILE_OUTPUT_PATTERN.search(file_path.name):
        return "looks like a generated PARAM output file"
    return None


def main() -> None:
    if not input_dir.exists():
        print(f"Input directory does not exist: {input_dir}")
        return

    output_file_path = get_output_path(input_dir, output_dir)
    empirical_rows: List[Dict[str, Any]] = []
    regression_rows: List[Dict[str, Any]] = []
    files_processed = 0

    app = xw.App(visible=False, add_book=False)
    app.display_alerts = False
    app.screen_updating = False

    original_calculation = app.calculation
    app.calculation = "manual"

    try:
        for file_path in sorted(input_dir.iterdir(), key=lambda p: p.name.lower()):
            skip_reason = should_skip_file(file_path)
            if skip_reason:
                print(f"Skipped: {file_path.name} ({skip_reason})")
                continue

            workbook = None
            try:
                labels = parse_file_labels(file_path.name)
            except ValueError as parse_error:
                print(f"Skipped: {file_path.name} ({parse_error})")
                continue

            try:
                workbook = app.books.open(str(file_path), update_links=False)
                empirical_sheet = get_sheet(workbook, EMPIRICAL_SHEET_NAME)
                regression_sheet = get_sheet(workbook, REGRESSION_SHEET_NAME)

                if empirical_sheet is None and regression_sheet is None:
                    print(f"Skipped: {file_path.name} (missing both model sheets)")
                    continue

                if empirical_sheet is not None:
                    empirical_rows.extend(
                        extract_empirical_rows(empirical_sheet, labels=labels, source_file=file_path.name)
                    )
                if regression_sheet is not None:
                    regression_rows.extend(
                        extract_regression_rows(regression_sheet, labels=labels, source_file=file_path.name)
                    )

                files_processed += 1
                print(f"Processed: {file_path.name}")
            except Exception as run_error:
                print(f"Skipped: {file_path.name} (processing error: {run_error})")
            finally:
                if workbook is not None:
                    safe_close_workbook(workbook)
    finally:
        try:
            app.calculation = original_calculation
        except Exception:
            pass
        app.quit()

    output_workbook = Workbook()
    empirical_ws = output_workbook.active
    empirical_ws.title = "empirical_candidates"
    regression_ws = output_workbook.create_sheet("regression_candidates")

    write_rows(empirical_ws, EMPIRICAL_OUTPUT_COLUMNS, empirical_rows)
    write_rows(regression_ws, REGRESSION_OUTPUT_COLUMNS, regression_rows)

    output_workbook.save(output_file_path)

    print(f"Output path: {output_file_path}")
    print(f"Files processed: {files_processed}")
    print(f"Empirical rows: {len(empirical_rows)}")
    print(f"Regression rows: {len(regression_rows)}")


if __name__ == "__main__":
    main()
