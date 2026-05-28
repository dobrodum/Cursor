#!/usr/bin/env python3
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
input_dir = Path("./input").resolve()
output_dir = Path("./output").resolve()

EMPIRICAL_SHEET = "Empirical Model"
REGRESSION_SHEET = "Regression Model"

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

DAY_BY_PHASE = {"early": 5, "mid": 15, "late": 25}
MONTH_BY_ABBR = {
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


def to_2d(values: Any) -> List[List[Any]]:
    if values is None:
        return []
    if not isinstance(values, (list, tuple)):
        return [[values]]
    if not values:
        return []
    first = values[0]
    if isinstance(first, (list, tuple)):
        rows: List[List[Any]] = []
        for row in values:
            if isinstance(row, (list, tuple)):
                rows.append(list(row))
            else:
                rows.append([row])
        return rows
    return [list(values)]


def build_sheet_snapshot(sheet: xw.Sheet) -> Dict[str, Any]:
    used = sheet.used_range
    matrix = to_2d(used.value)
    if not matrix:
        return {
            "start_row": used.row,
            "start_col": used.column,
            "end_row": used.row,
            "end_col": used.column,
            "matrix": [],
        }

    max_cols = max(len(row) for row in matrix)
    normalized = [list(row) + [None] * (max_cols - len(row)) for row in matrix]
    return {
        "start_row": used.row,
        "start_col": used.column,
        "end_row": used.row + len(normalized) - 1,
        "end_col": used.column + max_cols - 1,
        "matrix": normalized,
    }


def find_max_anchor(snapshot: Dict[str, Any]) -> Optional[Tuple[int, int]]:
    matrix: List[List[Any]] = snapshot["matrix"]
    start_row: int = snapshot["start_row"]
    start_col: int = snapshot["start_col"]

    partial_hits: List[Tuple[int, int]] = []
    for r_idx, row in enumerate(matrix):
        for c_idx, value in enumerate(row):
            if not isinstance(value, str):
                continue
            text = value.strip().lower()
            if text == "max":
                return start_row + r_idx, start_col + c_idx
            if re.search(r"\bmax\b", text):
                partial_hits.append((start_row + r_idx, start_col + c_idx))

    return partial_hits[0] if partial_hits else None


def safe_value(value: Any) -> Any:
    if value == "":
        return None
    return value


def as_number(value: Any) -> Optional[float]:
    if value is None or value == "":
        return None
    if isinstance(value, bool):
        return float(value)
    if isinstance(value, (int, float)):
        number = float(value)
        if math.isnan(number):
            return None
        return number
    if isinstance(value, str):
        text = value.strip().replace(",", "")
        if not text:
            return None
        is_pct = text.endswith("%")
        if is_pct:
            text = text[:-1].strip()
        try:
            parsed = float(text)
        except ValueError:
            return None
        return parsed / 100 if is_pct else parsed
    return None


def as_int(value: Any) -> Optional[int]:
    parsed = as_number(value)
    if parsed is None:
        return None
    return int(round(parsed))


def compute_range_width(max_value: Any, min_value: Any) -> Optional[float]:
    max_num = as_number(max_value)
    min_num = as_number(min_value)
    if max_num is None or min_num is None:
        return None
    return max_num - min_num


def set_formula2(cell: xw.Range, formula: str) -> None:
    try:
        cell.formula2 = formula
    except Exception:
        cell.formula = formula


def recalculate(app: xw.App) -> None:
    try:
        app.calculate()
    except Exception:
        app.api.Calculate()


def safe_close_workbook(workbook: xw.Book) -> None:
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
    except Exception as exc:
        print(f"Warning: unable to close workbook safely ({workbook.name}): {exc}")


def read_cell_by_offset(sheet: xw.Sheet, row: int, anchor_col: int, offset: int) -> Any:
    col = anchor_col + offset
    if col < 1:
        return None
    return safe_value(sheet.range((row, col)).value)


def parse_file_labels(file_path: Path) -> Dict[str, str]:
    stem = file_path.stem
    parts = [segment.strip() for segment in stem.split(" - ")]

    ticker_raw = parts[1] if len(parts) >= 2 else parts[0]
    ticker = re.sub(r"[^A-Za-z0-9_]", "", ticker_raw).upper() or "UNKNOWN"

    period_source = parts[2] if len(parts) >= 3 else stem
    period_match = re.search(
        r"(Early|Mid|Late)(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)(\d{4})",
        period_source,
        flags=re.IGNORECASE,
    )

    model_period = ""
    model_date = ""
    if period_match:
        phase = period_match.group(1).lower()
        month = period_match.group(2).lower()
        year = int(period_match.group(3))
        model_period = f"{phase.title()}{month.title()}_{year}"
        model_date = date(year, MONTH_BY_ABBR[month], DAY_BY_PHASE[phase]).isoformat()

    model = f"{ticker}_{model_period}" if model_period else ticker
    return {
        "model": model,
        "ticker": ticker,
        "model_period": model_period,
        "model_date": model_date,
    }


def contiguous_xy_rows(
    sheet: xw.Sheet,
    x_col: int,
    y_col: int,
    anchor_row: int,
) -> List[int]:
    rows: List[int] = []
    row = anchor_row - 1
    while row >= 1:
        x_value = as_number(sheet.range((row, x_col)).value)
        y_value = as_number(sheet.range((row, y_col)).value)
        if x_value is not None and y_value is not None:
            rows.append(row)
            row -= 1
            continue
        if rows:
            break
        row -= 1
    rows.reverse()
    return rows


def rounded_signature(values: Sequence[Any], precision: int = 10) -> Tuple[Any, ...]:
    signature: List[Any] = []
    for value in values:
        numeric = as_number(value)
        if numeric is not None:
            signature.append(round(numeric, precision))
        else:
            signature.append(value)
    return tuple(signature)


def extract_empirical_rows(
    workbook: xw.Book,
    labels: Dict[str, str],
    source_file: str,
) -> List[Dict[str, Any]]:
    try:
        sheet = workbook.sheets[EMPIRICAL_SHEET]
    except Exception:
        print(f"Skipped empirical extraction ({source_file}): '{EMPIRICAL_SHEET}' sheet not found")
        return []

    snapshot = build_sheet_snapshot(sheet)
    anchor = find_max_anchor(snapshot)
    if anchor is None:
        print(f"Skipped empirical extraction ({source_file}): 'max' anchor not found")
        return []

    anchor_row, anchor_col = anchor
    offsets = {
        "sales_captured_in_db_pct": -8,
        "growth_rate_pct": -7,
        "quarterly_sales": -6,
        "num_quarters_used": -5,
        "last_quarter_used": -4,
        "avg_penetration_pct": -3,
        "reported_sales": -2,
        "forecast_value": -1,
        "forecast_max": 0,
        "forecast_min": 1,
    }

    n_quarters = 10
    first_data_row = anchor_row + 1
    penetration_source_col = anchor_col + offsets["sales_captured_in_db_pct"]
    if penetration_source_col < 1:
        print(
            f"Skipped empirical extraction ({source_file}): invalid anchor offset "
            "for penetration source column"
        )
        return []

    helper_col = max(anchor_col + 2, snapshot["end_col"] + 2)
    source_rel_col = penetration_source_col - helper_col

    for idx in range(n_quarters):
        row = first_data_row + idx
        formula = f'=IFERROR(AVERAGE(R[{-idx}]C[{source_rel_col}]:RC[{source_rel_col}]), "")'
        set_formula2(sheet.range((row, helper_col)), formula)

    recalculate(workbook.app)
    helper_values = [
        safe_value(sheet.range((first_data_row + idx, helper_col)).value) for idx in range(n_quarters)
    ]
    sheet.range((first_data_row, helper_col), (first_data_row + n_quarters - 1, helper_col)).clear_contents()

    rows: List[Dict[str, Any]] = []
    for idx in range(n_quarters):
        row_idx = first_data_row + idx
        num_quarters_used = read_cell_by_offset(
            sheet, row_idx, anchor_col, offsets["num_quarters_used"]
        )
        last_quarter_used = read_cell_by_offset(
            sheet, row_idx, anchor_col, offsets["last_quarter_used"]
        )
        forecast_value = read_cell_by_offset(sheet, row_idx, anchor_col, offsets["forecast_value"])
        reported_sales = read_cell_by_offset(sheet, row_idx, anchor_col, offsets["reported_sales"])
        forecast_max = read_cell_by_offset(sheet, row_idx, anchor_col, offsets["forecast_max"])
        forecast_min = read_cell_by_offset(sheet, row_idx, anchor_col, offsets["forecast_min"])
        quarterly_sales = read_cell_by_offset(sheet, row_idx, anchor_col, offsets["quarterly_sales"])
        growth_rate_pct = read_cell_by_offset(sheet, row_idx, anchor_col, offsets["growth_rate_pct"])
        sales_captured_in_db_pct = read_cell_by_offset(
            sheet,
            row_idx,
            anchor_col,
            offsets["sales_captured_in_db_pct"],
        )

        avg_penetration_existing = read_cell_by_offset(
            sheet,
            row_idx,
            anchor_col,
            offsets["avg_penetration_pct"],
        )
        avg_penetration_pct = (
            avg_penetration_existing
            if avg_penetration_existing is not None
            else helper_values[idx]
        )

        if avg_penetration_pct is None:
            quarterly_sales_num = as_number(quarterly_sales)
            reported_sales_num = as_number(reported_sales)
            if quarterly_sales_num is not None and reported_sales_num not in (None, 0):
                avg_penetration_pct = quarterly_sales_num / reported_sales_num

        if all(
            value is None
            for value in (
                forecast_value,
                reported_sales,
                forecast_max,
                forecast_min,
                avg_penetration_pct,
                quarterly_sales,
            )
        ):
            continue

        num_quarters_int = as_int(num_quarters_used) or (idx + 1)
        rows.append(
            {
                "model": labels["model"],
                "ticker": labels["ticker"],
                "model_period": labels["model_period"],
                "model_date": labels["model_date"],
                "method": "empirical",
                "parameter_name": "avg_penetration_pct",
                "parameter_value": avg_penetration_pct,
                "num_quarters_used": num_quarters_int,
                "last_quarter_used": last_quarter_used,
                "forecast_value": forecast_value,
                "actual_value": reported_sales,
                "forecast_max": forecast_max,
                "forecast_min": forecast_min,
                "range_width": compute_range_width(forecast_max, forecast_min),
                "avg_penetration_pct": avg_penetration_pct,
                "quarterly_sales": quarterly_sales,
                "reported_sales": reported_sales,
                "growth_rate_pct": growth_rate_pct,
                "sales_captured_in_db_pct": sales_captured_in_db_pct,
                "source_file": source_file,
            }
        )

    return rows


def extract_regression_rows(
    workbook: xw.Book,
    labels: Dict[str, str],
    source_file: str,
) -> List[Dict[str, Any]]:
    try:
        sheet = workbook.sheets[REGRESSION_SHEET]
    except Exception:
        print(f"Skipped regression extraction ({source_file}): '{REGRESSION_SHEET}' sheet not found")
        return []

    snapshot = build_sheet_snapshot(sheet)
    anchor = find_max_anchor(snapshot)
    if anchor is None:
        print(f"Skipped regression extraction ({source_file}): 'max' anchor not found")
        return []

    anchor_row, anchor_col = anchor
    y_col = anchor_col - 7
    x_col = anchor_col - 11
    if x_col < 1 or y_col < 1:
        print(f"Skipped regression extraction ({source_file}): invalid x/y anchor offsets")
        return []

    history_rows = contiguous_xy_rows(sheet=sheet, x_col=x_col, y_col=y_col, anchor_row=anchor_row)
    max_n = min(10, len(history_rows))
    if max_n < 2:
        print(
            f"Skipped regression extraction ({source_file}): not enough x/y history rows "
            "for INTERCEPT/SLOPE"
        )
        return []

    offsets = {
        "num_quarters_used": -5,
        "actual_value": -2,
        "forecast_value": -1,  # TOT FCST w/o SA
        "forecast_max": 0,
        "forecast_min": 1,
    }
    first_data_row = anchor_row + 1
    intercept_col = max(anchor_col + 2, snapshot["end_col"] + 2)
    slope_col = intercept_col + 1

    n_values = list(range(2, max_n + 1))
    for idx, n_quarters in enumerate(n_values):
        row_idx = first_data_row + idx
        start_row = history_rows[-n_quarters]
        end_row = history_rows[-1]

        intercept_formula = (
            f'=IFERROR(INTERCEPT(R{start_row}C{y_col}:R{end_row}C{y_col},'
            f'R{start_row}C{x_col}:R{end_row}C{x_col}), "")'
        )
        slope_formula = (
            f'=IFERROR(SLOPE(R{start_row}C{y_col}:R{end_row}C{y_col},'
            f'R{start_row}C{x_col}:R{end_row}C{x_col}), "")'
        )
        set_formula2(sheet.range((row_idx, intercept_col)), intercept_formula)
        set_formula2(sheet.range((row_idx, slope_col)), slope_formula)

    recalculate(workbook.app)

    rows: List[Dict[str, Any]] = []
    previous_signature: Optional[Tuple[Any, ...]] = None
    for idx, n_quarters in enumerate(n_values):
        row_idx = first_data_row + idx
        num_quarters_used = read_cell_by_offset(sheet, row_idx, anchor_col, offsets["num_quarters_used"])
        forecast_value = read_cell_by_offset(sheet, row_idx, anchor_col, offsets["forecast_value"])
        actual_value = read_cell_by_offset(sheet, row_idx, anchor_col, offsets["actual_value"])
        forecast_max = read_cell_by_offset(sheet, row_idx, anchor_col, offsets["forecast_max"])
        forecast_min = read_cell_by_offset(sheet, row_idx, anchor_col, offsets["forecast_min"])
        intercept = safe_value(sheet.range((row_idx, intercept_col)).value)
        slope = safe_value(sheet.range((row_idx, slope_col)).value)

        signature = rounded_signature(
            [
                as_int(num_quarters_used) or n_quarters,
                forecast_value,
                forecast_max,
                forecast_min,
                intercept,
                slope,
            ]
        )
        if previous_signature is not None and signature == previous_signature:
            continue
        previous_signature = signature

        rows.append(
            {
                "model": labels["model"],
                "ticker": labels["ticker"],
                "model_period": labels["model_period"],
                "model_date": labels["model_date"],
                "method": "regression",
                "parameter_name": "num_quarters_used",
                "parameter_value": as_int(num_quarters_used) or n_quarters,
                "num_quarters_used": as_int(num_quarters_used) or n_quarters,
                "forecast_value": forecast_value,
                "actual_value": actual_value,
                "forecast_max": forecast_max,
                "forecast_min": forecast_min,
                "range_width": compute_range_width(forecast_max, forecast_min),
                "intercept": intercept,
                "slope": slope,
                "source_file": source_file,
            }
        )

    if n_values:
        last_formula_row = first_data_row + len(n_values) - 1
        sheet.range((first_data_row, intercept_col), (last_formula_row, slope_col)).clear_contents()

    return rows


def choose_output_path(input_path: Path, output_path: Path) -> Path:
    stem = f"{input_path.name}_PARAM"
    candidate = output_path / f"{stem}.xlsx"
    if not candidate.exists():
        return candidate

    suffix = 1
    while True:
        candidate = output_path / f"{stem}.{suffix}.xlsx"
        if not candidate.exists():
            return candidate
        suffix += 1


def write_sheet_rows(
    worksheet,
    headers: Sequence[str],
    rows: Iterable[Dict[str, Any]],
) -> None:
    worksheet.append(list(headers))
    for row in rows:
        worksheet.append([row.get(header) for header in headers])

    for cell in worksheet[1]:
        cell.font = Font(bold=True)

    worksheet.freeze_panes = "A2"
    worksheet.auto_filter.ref = f"A1:{get_column_letter(worksheet.max_column)}{worksheet.max_row}"

    for col_idx, header in enumerate(headers, start=1):
        max_len = len(header)
        for row_idx in range(2, worksheet.max_row + 1):
            value = worksheet.cell(row=row_idx, column=col_idx).value
            if value is None:
                continue
            max_len = max(max_len, len(str(value)))
        worksheet.column_dimensions[get_column_letter(col_idx)].width = min(max(max_len + 2, 12), 48)


def write_output_workbook(
    destination: Path,
    empirical_rows: Sequence[Dict[str, Any]],
    regression_rows: Sequence[Dict[str, Any]],
) -> None:
    out_wb = Workbook()
    empirical_ws = out_wb.active
    empirical_ws.title = "empirical_candidates"
    regression_ws = out_wb.create_sheet("regression_candidates")

    write_sheet_rows(empirical_ws, EMPIRICAL_HEADERS, empirical_rows)
    write_sheet_rows(regression_ws, REGRESSION_HEADERS, regression_rows)
    out_wb.save(destination)


def create_excel_app() -> xw.App:
    app = xw.App(visible=False, add_book=False)
    for attr, value in (
        ("display_alerts", False),
        ("screen_updating", False),
        ("enable_events", False),
    ):
        try:
            setattr(app, attr, value)
        except Exception:
            pass
    try:
        app.calculation = "manual"
    except Exception:
        pass
    return app


def run() -> None:
    if not input_dir.exists():
        raise FileNotFoundError(f"input_dir does not exist: {input_dir}")
    if not input_dir.is_dir():
        raise NotADirectoryError(f"input_dir is not a directory: {input_dir}")

    output_dir.mkdir(parents=True, exist_ok=True)
    destination = choose_output_path(input_path=input_dir, output_path=output_dir)

    empirical_rows: List[Dict[str, Any]] = []
    regression_rows: List[Dict[str, Any]] = []
    processed_files = 0

    app = create_excel_app()
    try:
        for file_path in sorted(input_dir.iterdir(), key=lambda p: p.name.lower()):
            if not file_path.is_file():
                print(f"Skipped file: {file_path.name} (not a regular file)")
                continue
            if file_path.name.startswith("~"):
                print(f"Skipped file: {file_path.name} (temporary file)")
                continue
            if file_path.suffix.lower() != ".xlsx":
                print(f"Skipped file: {file_path.name} (not .xlsx)")
                continue

            workbook = None
            try:
                workbook = app.books.open(str(file_path), update_links=False)
                labels = parse_file_labels(file_path)
                empirical_rows.extend(
                    extract_empirical_rows(
                        workbook=workbook,
                        labels=labels,
                        source_file=file_path.name,
                    )
                )
                regression_rows.extend(
                    extract_regression_rows(
                        workbook=workbook,
                        labels=labels,
                        source_file=file_path.name,
                    )
                )
                processed_files += 1
                print(f"Processed file: {file_path.name}")
            except Exception as exc:
                print(f"Skipped file: {file_path.name} (processing error: {exc})")
            finally:
                if workbook is not None:
                    safe_close_workbook(workbook)
    finally:
        app.quit()

    write_output_workbook(
        destination=destination,
        empirical_rows=empirical_rows,
        regression_rows=regression_rows,
    )

    print(f"Output path: {destination}")
    print(f"Number of files processed: {processed_files}")
    print(f"Number of empirical rows: {len(empirical_rows)}")
    print(f"Number of regression rows: {len(regression_rows)}")


if __name__ == "__main__":
    run()
