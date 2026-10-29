#!/usr/bin/env python3
"""Extract empirical/regression candidate rows from Excel model workbooks.

Runtime goals:
- One hidden Excel app for the whole run.
- Open each source workbook once.
- Process both model sheets while open, then close without saving.
"""

from __future__ import annotations

import math
import re
from datetime import date
from pathlib import Path
from typing import Any, Dict, List, Optional, Sequence, Tuple

import xlwings as xw
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

# ---------------------------------------------------------------------------
# User-configurable paths
# ---------------------------------------------------------------------------
input_dir = Path("/path/to/input")
output_dir = Path("/path/to/output")


EMPIRICAL_MODEL_SHEET = "Empirical Model"
REGRESSION_MODEL_SHEET = "Regression Model"
N_QUARTERS = 10

EMPIRICAL_HEADERS = [
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

REGRESSION_HEADERS = [
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

PERIOD_DAY_MAP = {"early": 5, "mid": 15, "late": 25}


def normalize_text(value: Any) -> str:
    if value is None:
        return ""
    text = str(value).strip().lower()
    return re.sub(r"\s+", " ", text)


def is_blank(value: Any) -> bool:
    return normalize_text(value) == ""


def to_float(value: Any) -> Optional[float]:
    if value is None:
        return None
    if isinstance(value, bool):
        return float(value)
    if isinstance(value, (int, float)):
        if isinstance(value, float) and math.isnan(value):
            return None
        return float(value)
    text = str(value).strip().replace(",", "")
    if not text:
        return None
    if text.endswith("%"):
        text = text[:-1].strip()
    try:
        return float(text)
    except ValueError:
        return None


def to_int(value: Any) -> Optional[int]:
    as_float = to_float(value)
    if as_float is None:
        return None
    try:
        return int(round(as_float))
    except (TypeError, ValueError):
        return None


def safe_numeric_delta(max_value: Any, min_value: Any) -> Optional[float]:
    max_num = to_float(max_value)
    min_num = to_float(min_value)
    if max_num is None or min_num is None:
        return None
    return max_num - min_num


def parse_file_metadata(file_path: Path) -> Dict[str, str]:
    stem = file_path.stem
    parts = [part.strip() for part in stem.split("-")]

    ticker = "UNKNOWN"
    if len(parts) >= 2 and parts[1]:
        ticker_candidate = re.split(r"[_\s]+", parts[1])[0].strip()
        if ticker_candidate:
            ticker = ticker_candidate.upper()
    else:
        ticker_match = re.search(r"\b[A-Z]{2,8}\b", stem)
        if ticker_match:
            ticker = ticker_match.group(0)

    period_match = re.search(r"(Early|Mid|Late)([A-Za-z]{3,9})(\d{4})", stem, re.IGNORECASE)
    if period_match:
        period_prefix = period_match.group(1).title()
        month_token = period_match.group(2)[:3].title()
        year_token = int(period_match.group(3))

        month_num = MONTH_MAP.get(month_token.lower())
        if month_num is None:
            model_period = f"{period_prefix}{month_token}_{year_token}"
            model_date = ""
        else:
            day_num = PERIOD_DAY_MAP[period_prefix.lower()]
            model_period = f"{period_prefix}{month_token}_{year_token}"
            model_date = date(year_token, month_num, day_num).isoformat()
    else:
        model_period = "Unknown_0000"
        model_date = ""

    model = f"{ticker}_{model_period}" if ticker != "UNKNOWN" else model_period

    return {
        "model": model,
        "ticker": ticker,
        "model_period": model_period,
        "model_date": model_date,
    }


def next_output_path(source_input_dir: Path, destination_output_dir: Path) -> Path:
    destination_output_dir.mkdir(parents=True, exist_ok=True)
    base_name = f"{source_input_dir.name}_PARAM"
    candidate = destination_output_dir / f"{base_name}.xlsx"
    suffix = 1
    while candidate.exists():
        candidate = destination_output_dir / f"{base_name}.{suffix}.xlsx"
        suffix += 1
    return candidate


def to_matrix(values: Any) -> List[List[Any]]:
    if values is None:
        return []
    if not isinstance(values, list):
        return [[values]]
    if values and isinstance(values[0], list):
        return values
    return [values]


def find_anchor_cell(sheet: xw.Sheet, anchor_value: str = "max") -> Optional[Tuple[int, int]]:
    used = sheet.used_range
    matrix = to_matrix(used.value)
    if not matrix:
        return None

    anchor_key = normalize_text(anchor_value)
    first_row = used.row
    first_col = used.column

    for r_idx, row_values in enumerate(matrix):
        for c_idx, cell_value in enumerate(row_values):
            if normalize_text(cell_value) == anchor_key:
                return first_row + r_idx, first_col + c_idx
    return None


def find_header_col(
    sheet: xw.Sheet,
    header_row: int,
    anchor_col: int,
    search_terms: Sequence[str],
    default_col: int,
    window: int = 24,
    exclude_terms: Optional[Sequence[str]] = None,
) -> int:
    exclude_terms = tuple(normalize_text(term) for term in (exclude_terms or ()))
    lowered_terms = tuple(normalize_text(term) for term in search_terms)

    min_col = max(1, anchor_col - window)
    max_col = anchor_col + window
    for col in range(min_col, max_col + 1):
        header_value = normalize_text(sheet.cells(header_row, col).value)
        if not header_value:
            continue
        if exclude_terms and any(excl in header_value for excl in exclude_terms):
            continue
        if any(term in header_value for term in lowered_terms):
            return col
    return max(1, default_col)


def close_source_workbook(wb: xw.Book, source_name: str) -> None:
    try:
        wb.close(save=False)
        return
    except TypeError:
        # Some backends only support close() with no keyword args.
        pass
    except Exception as exc:
        print(f"Warning: close(save=False) failed for {source_name}: {exc}")

    try:
        wb.api.Close(SaveChanges=False)
        return
    except Exception:
        pass

    try:
        wb.close()
    except Exception as exc:
        print(f"Warning: could not close workbook safely for {source_name}: {exc}")


def build_empirical_rows(
    wb: xw.Book, sheet: xw.Sheet, metadata: Dict[str, str], source_file: str
) -> List[Dict[str, Any]]:
    anchor = find_anchor_cell(sheet, "max")
    if anchor is None:
        print(f"Skipped empirical extraction for {source_file}: 'max' anchor not found.")
        return []

    anchor_row, anchor_col = anchor
    data_start_row = anchor_row + 1

    num_quarters_col = find_header_col(
        sheet,
        anchor_row,
        anchor_col,
        search_terms=("num quarters", "n quarters", "# quarters", "quarters used"),
        default_col=anchor_col - 8,
    )
    last_quarter_col = find_header_col(
        sheet,
        anchor_row,
        anchor_col,
        search_terms=("last quarter", "quarter used"),
        default_col=anchor_col - 7,
    )
    forecast_value_col = find_header_col(
        sheet,
        anchor_row,
        anchor_col,
        search_terms=("estimated total sold", "forecast", "fcst"),
        default_col=anchor_col - 3,
    )
    actual_value_col = find_header_col(
        sheet,
        anchor_row,
        anchor_col,
        search_terms=("actual", "reported sales"),
        default_col=anchor_col - 2,
    )
    forecast_min_col = find_header_col(
        sheet,
        anchor_row,
        anchor_col,
        search_terms=("min",),
        default_col=anchor_col + 1,
        window=4,
    )
    quarterly_sales_col = find_header_col(
        sheet,
        anchor_row,
        anchor_col,
        search_terms=("quarterly sales",),
        default_col=anchor_col - 6,
    )
    reported_sales_col = find_header_col(
        sheet,
        anchor_row,
        anchor_col,
        search_terms=("reported sales",),
        default_col=anchor_col - 5,
    )
    growth_rate_col = find_header_col(
        sheet,
        anchor_row,
        anchor_col,
        search_terms=("growth", "growth rate"),
        default_col=anchor_col - 4,
    )
    sales_captured_col = find_header_col(
        sheet,
        anchor_row,
        anchor_col,
        search_terms=("sales captured", "captured in db"),
        default_col=anchor_col - 1,
    )
    penetration_source_col = find_header_col(
        sheet,
        anchor_row,
        anchor_col,
        search_terms=("penetration",),
        default_col=anchor_col - 7,
        exclude_terms=("avg", "average"),
    )

    helper_col = max(
        anchor_col,
        forecast_min_col,
        forecast_value_col,
        actual_value_col,
        sales_captured_col,
    ) + 5

    # Build all helper formulas first, then calculate once.
    for i in range(N_QUARTERS):
        row = data_start_row + i
        num_quarters = to_int(sheet.cells(row, num_quarters_col).value) or (i + 1)
        num_quarters = max(1, min(N_QUARTERS, num_quarters))
        helper_cell = sheet.cells(row, helper_col)
        source_offset = penetration_source_col - helper_col
        start_offset = -(num_quarters - 1)
        helper_cell.formula2 = (
            f'=IFERROR(AVERAGE(R[{start_offset}]C[{source_offset}]:RC[{source_offset}]),"")'
        )

    wb.app.calculate()

    empirical_rows: List[Dict[str, Any]] = []
    signal_cols = (
        num_quarters_col,
        forecast_value_col,
        anchor_col,
        forecast_min_col,
        quarterly_sales_col,
        reported_sales_col,
    )

    for i in range(N_QUARTERS):
        row = data_start_row + i
        if all(is_blank(sheet.cells(row, col).value) for col in signal_cols):
            continue

        num_quarters = to_int(sheet.cells(row, num_quarters_col).value) or (i + 1)
        avg_penetration = to_float(sheet.cells(row, helper_col).value)
        forecast_value = to_float(sheet.cells(row, forecast_value_col).value)
        actual_value = to_float(sheet.cells(row, actual_value_col).value)
        forecast_max = to_float(sheet.cells(row, anchor_col).value)
        forecast_min = to_float(sheet.cells(row, forecast_min_col).value)
        reported_sales = to_float(sheet.cells(row, reported_sales_col).value)

        if actual_value is None:
            actual_value = reported_sales

        empirical_rows.append(
            {
                "model": metadata["model"],
                "ticker": metadata["ticker"],
                "model_period": metadata["model_period"],
                "model_date": metadata["model_date"],
                "method": "empirical",
                "parameter_name": "avg_penetration_pct",
                "parameter_value": avg_penetration,
                "num_quarters_used": num_quarters,
                "last_quarter_used": sheet.cells(row, last_quarter_col).value,
                "forecast_value": forecast_value,
                "actual_value": actual_value,
                "forecast_max": forecast_max,
                "forecast_min": forecast_min,
                "range_width": safe_numeric_delta(forecast_max, forecast_min),
                "avg_penetration_pct": avg_penetration,
                "quarterly_sales": to_float(sheet.cells(row, quarterly_sales_col).value),
                "reported_sales": reported_sales,
                "growth_rate_pct": to_float(sheet.cells(row, growth_rate_col).value),
                "sales_captured_in_db_pct": to_float(sheet.cells(row, sales_captured_col).value),
                "source_file": source_file,
            }
        )

    return empirical_rows


def build_regression_rows(
    wb: xw.Book, sheet: xw.Sheet, metadata: Dict[str, str], source_file: str
) -> List[Dict[str, Any]]:
    anchor = find_anchor_cell(sheet, "max")
    if anchor is None:
        print(f"Skipped regression extraction for {source_file}: 'max' anchor not found.")
        return []

    anchor_row, anchor_col = anchor
    data_start_row = anchor_row + 1

    y_col = anchor_col - 7
    x_col = anchor_col - 11

    num_quarters_col = find_header_col(
        sheet,
        anchor_row,
        anchor_col,
        search_terms=("num quarters", "n quarters", "# quarters", "quarters used"),
        default_col=anchor_col - 9,
    )
    forecast_col = find_header_col(
        sheet,
        anchor_row,
        anchor_col,
        search_terms=("tot fcst w/o sa", "tot fcst", "forecast"),
        default_col=anchor_col - 2,
    )
    min_col = find_header_col(
        sheet,
        anchor_row,
        anchor_col,
        search_terms=("min",),
        default_col=anchor_col + 1,
        window=4,
    )
    actual_col = find_header_col(
        sheet,
        anchor_row,
        anchor_col,
        search_terms=("actual", "reported"),
        default_col=anchor_col - 1,
    )

    helper_intercept_col = max(anchor_col, min_col, forecast_col, num_quarters_col) + 5
    helper_slope_col = helper_intercept_col + 1

    for i in range(N_QUARTERS):
        row = data_start_row + i
        num_quarters = to_int(sheet.cells(row, num_quarters_col).value) or (i + 1)
        num_quarters = max(2, min(N_QUARTERS, num_quarters))
        start_offset = -(num_quarters - 1)

        intercept_cell = sheet.cells(row, helper_intercept_col)
        slope_cell = sheet.cells(row, helper_slope_col)

        y_offset_for_intercept = y_col - helper_intercept_col
        x_offset_for_intercept = x_col - helper_intercept_col
        intercept_cell.formula2 = (
            f'=IFERROR(INTERCEPT('
            f'R[{start_offset}]C[{y_offset_for_intercept}]:RC[{y_offset_for_intercept}],'
            f'R[{start_offset}]C[{x_offset_for_intercept}]:RC[{x_offset_for_intercept}]'
            f'),"")'
        )

        y_offset_for_slope = y_col - helper_slope_col
        x_offset_for_slope = x_col - helper_slope_col
        slope_cell.formula2 = (
            f'=IFERROR(SLOPE('
            f'R[{start_offset}]C[{y_offset_for_slope}]:RC[{y_offset_for_slope}],'
            f'R[{start_offset}]C[{x_offset_for_slope}]:RC[{x_offset_for_slope}]'
            f'),"")'
        )

    wb.app.calculate()

    regression_rows: List[Dict[str, Any]] = []
    last_signature: Optional[Tuple[Any, ...]] = None
    signal_cols = (num_quarters_col, forecast_col, anchor_col, min_col)

    for i in range(N_QUARTERS):
        row = data_start_row + i
        if all(is_blank(sheet.cells(row, col).value) for col in signal_cols):
            continue

        num_quarters = to_int(sheet.cells(row, num_quarters_col).value) or (i + 1)
        forecast_value = to_float(sheet.cells(row, forecast_col).value)
        forecast_max = to_float(sheet.cells(row, anchor_col).value)
        forecast_min = to_float(sheet.cells(row, min_col).value)
        intercept = to_float(sheet.cells(row, helper_intercept_col).value)
        slope = to_float(sheet.cells(row, helper_slope_col).value)

        signature = (
            num_quarters,
            round(intercept, 12) if intercept is not None else None,
            round(slope, 12) if slope is not None else None,
            round(forecast_value, 12) if forecast_value is not None else None,
            round(forecast_max, 12) if forecast_max is not None else None,
            round(forecast_min, 12) if forecast_min is not None else None,
        )
        if signature == last_signature:
            continue
        last_signature = signature

        regression_rows.append(
            {
                "model": metadata["model"],
                "ticker": metadata["ticker"],
                "model_period": metadata["model_period"],
                "model_date": metadata["model_date"],
                "method": "regression",
                "parameter_name": "num_quarters_used",
                "parameter_value": num_quarters,
                "num_quarters_used": num_quarters,
                "forecast_value": forecast_value,
                "actual_value": to_float(sheet.cells(row, actual_col).value),
                "forecast_max": forecast_max,
                "forecast_min": forecast_min,
                "range_width": safe_numeric_delta(forecast_max, forecast_min),
                "intercept": intercept,
                "slope": slope,
                "source_file": source_file,
            }
        )

    return regression_rows


def write_output_sheet(ws, headers: Sequence[str], rows: Sequence[Dict[str, Any]]) -> None:
    ws.append(list(headers))
    for row in rows:
        ws.append([row.get(header) for header in headers])

    for cell in ws[1]:
        cell.font = Font(bold=True)

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions

    for col_idx, header in enumerate(headers, start=1):
        max_len = len(header)
        for row_idx in range(2, ws.max_row + 1):
            cell_value = ws.cell(row=row_idx, column=col_idx).value
            if cell_value is None:
                continue
            max_len = max(max_len, len(str(cell_value)))
        ws.column_dimensions[get_column_letter(col_idx)].width = min(max(12, max_len + 2), 44)


def main() -> None:
    if not input_dir.exists() or not input_dir.is_dir():
        raise FileNotFoundError(f"input_dir does not exist or is not a directory: {input_dir}")

    source_entries = sorted(input_dir.iterdir(), key=lambda p: p.name.lower())
    empirical_rows: List[Dict[str, Any]] = []
    regression_rows: List[Dict[str, Any]] = []
    files_processed = 0

    app = xw.App(visible=False, add_book=False)
    app.display_alerts = False
    app.screen_updating = False

    try:
        for file_path in source_entries:
            if not file_path.is_file():
                print(f"Skipped file: {file_path.name} (not a file)")
                continue
            if file_path.name.startswith("~"):
                print(f"Skipped file: {file_path.name} (temporary file)")
                continue
            if file_path.suffix.lower() != ".xlsx":
                print(f"Skipped file: {file_path.name} (not an .xlsx file)")
                continue

            print(f"Processed file: {file_path.name}")
            wb: Optional[xw.Book] = None

            try:
                wb = app.books.open(str(file_path), update_links=False)
                files_processed += 1
                metadata = parse_file_metadata(file_path)

                try:
                    empirical_sheet = wb.sheets[EMPIRICAL_MODEL_SHEET]
                    empirical_rows.extend(
                        build_empirical_rows(wb, empirical_sheet, metadata, file_path.name)
                    )
                except Exception as exc:
                    print(
                        f"Skipped file: {file_path.name} "
                        f"(empirical sheet error: {exc})"
                    )

                try:
                    regression_sheet = wb.sheets[REGRESSION_MODEL_SHEET]
                    regression_rows.extend(
                        build_regression_rows(wb, regression_sheet, metadata, file_path.name)
                    )
                except Exception as exc:
                    print(
                        f"Skipped file: {file_path.name} "
                        f"(regression sheet error: {exc})"
                    )

            except Exception as exc:
                print(f"Skipped file: {file_path.name} (open failed: {exc})")
            finally:
                if wb is not None:
                    close_source_workbook(wb, file_path.name)
    finally:
        app.quit()

    output_path = next_output_path(input_dir, output_dir)

    out_wb = Workbook()
    empirical_ws = out_wb.active
    empirical_ws.title = "empirical_candidates"
    write_output_sheet(empirical_ws, EMPIRICAL_HEADERS, empirical_rows)

    regression_ws = out_wb.create_sheet("regression_candidates")
    write_output_sheet(regression_ws, REGRESSION_HEADERS, regression_rows)

    out_wb.save(output_path)

    print(f"Output path: {output_path}")
    print(f"Files processed: {files_processed}")
    print(f"Empirical rows: {len(empirical_rows)}")
    print(f"Regression rows: {len(regression_rows)}")


if __name__ == "__main__":
    main()
