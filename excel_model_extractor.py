#!/usr/bin/env python3
from __future__ import annotations

import math
import re
from datetime import date
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Sequence, Tuple

import xwings as xw
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

# -----------------------------------------------------------------------------
# Configure these two paths before running.
# -----------------------------------------------------------------------------
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

PERIOD_DAY_MAP = {"early": 5, "mid": 15, "late": 25}
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


def normalize_text(value: Any) -> str:
    if value is None:
        return ""
    text = str(value).strip().lower()
    text = text.replace("\n", " ")
    text = re.sub(r"[^a-z0-9%]+", " ", text)
    return re.sub(r"\s+", " ", text).strip()


def to_float(value: Any) -> Optional[float]:
    if value is None:
        return None
    if isinstance(value, (int, float)):
        if isinstance(value, float) and (math.isnan(value) or math.isinf(value)):
            return None
        return float(value)

    text = str(value).strip().replace(",", "")
    if not text or text.startswith("#"):
        return None
    try:
        return float(text)
    except ValueError:
        return None


def to_int(value: Any) -> Optional[int]:
    number = to_float(value)
    if number is None:
        return None
    return int(round(number))


def as_2d(values: Any) -> List[List[Any]]:
    if values is None:
        return []
    if not isinstance(values, (list, tuple)):
        return [[values]]
    if values and not isinstance(values[0], (list, tuple)):
        return [list(values)]
    return [list(row) for row in values]


def subtract_or_none(a: Optional[float], b: Optional[float]) -> Optional[float]:
    if a is None or b is None:
        return None
    return a - b


def safe_close_workbook(workbook: xw.Book) -> None:
    try:
        workbook.close(save=False)
        return
    except TypeError:
        pass
    except Exception:
        pass

    try:
        workbook.close(False)
        return
    except Exception:
        pass

    try:
        workbook.api.Close(SaveChanges=False)
        return
    except Exception:
        pass

    try:
        workbook.saved = True
        workbook.close()
    except Exception:
        pass


def set_formula2(cell: xw.Range, formula_r1c1: str) -> None:
    try:
        cell.formula2 = formula_r1c1
    except Exception:
        cell.formula = formula_r1c1


def parse_file_labels(file_path: Path) -> Dict[str, str]:
    stem = file_path.stem.strip()
    parts = [part.strip() for part in stem.split(" - ") if part.strip()]

    ticker = ""
    if len(parts) >= 2:
        ticker = parts[1].upper()
    else:
        ticker_match = re.search(r"\b[A-Z]{2,8}\b", stem)
        if ticker_match:
            ticker = ticker_match.group(0).upper()

    period_match = re.search(
        r"(Early|Mid|Late)([A-Za-z]{3})(\d{4})", stem, flags=re.IGNORECASE
    )
    model_period = ""
    model_date = ""
    if period_match:
        period_label = period_match.group(1).title()
        month_abbrev = period_match.group(2).title()
        year_text = period_match.group(3)

        day = PERIOD_DAY_MAP.get(period_label.lower())
        month = MONTH_MAP.get(month_abbrev.lower())
        if day and month:
            year = int(year_text)
            model_period = f"{period_label}{month_abbrev}_{year}"
            model_date = date(year, month, day).isoformat()

    model = ""
    if ticker and model_period:
        model = f"{ticker}_{model_period}"
    elif ticker:
        model = ticker
    else:
        model = stem

    return {
        "model": model,
        "ticker": ticker,
        "model_period": model_period,
        "model_date": model_date,
    }


def get_output_path(input_folder: Path, out_folder: Path) -> Path:
    out_folder.mkdir(parents=True, exist_ok=True)
    base_name = f"{input_folder.name}_PARAM.xlsx"
    base_path = out_folder / base_name
    if not base_path.exists():
        return base_path

    suffix = 1
    while True:
        candidate = out_folder / f"{input_folder.name}_PARAM.{suffix}.xlsx"
        if not candidate.exists():
            return candidate
        suffix += 1


def get_sheet_if_exists(workbook: xw.Book, sheet_name: str) -> Optional[xw.Sheet]:
    try:
        return workbook.sheets[sheet_name]
    except Exception:
        return None


def find_anchor(ws: xw.Sheet, anchor_text: str = "max") -> Optional[Tuple[int, int]]:
    anchor_norm = normalize_text(anchor_text)

    try:
        found = ws.api.Cells.Find(What=anchor_text, LookAt=1, MatchCase=False)
        if found is not None:
            return int(found.Row), int(found.Column)
    except Exception:
        pass

    used = ws.used_range
    used_values = as_2d(used.value)
    if not used_values:
        return None

    for r_idx, row_values in enumerate(used_values):
        for c_idx, value in enumerate(row_values):
            if normalize_text(value) == anchor_norm:
                row = int(used.row) + r_idx
                col = int(used.column) + c_idx
                return row, col
    return None


def get_header_entries(
    ws: xw.Sheet, header_row: int, anchor_col: int, window: int = 24
) -> List[Tuple[int, str]]:
    entries: List[Tuple[int, str]] = []
    start_col = max(1, anchor_col - window)
    end_col = anchor_col + window

    for col in range(start_col, end_col + 1):
        header = normalize_text(ws.cells(header_row, col).value)
        if header:
            entries.append((col, header))

    entries.sort(key=lambda item: abs(item[0] - anchor_col))
    return entries


def find_column(
    entries: Sequence[Tuple[int, str]],
    include_tokens: Sequence[str],
    exclude_tokens: Sequence[str] = (),
) -> Optional[int]:
    for col, header in entries:
        if all(token in header for token in include_tokens) and not any(
            token in header for token in exclude_tokens
        ):
            return col
    return None


def is_effectively_empty(values: Iterable[Any]) -> bool:
    for value in values:
        if value is None:
            continue
        if isinstance(value, str) and not value.strip():
            continue
        return False
    return True


def extract_empirical_rows(
    workbook: xw.Book,
    labels: Dict[str, str],
    source_file: str,
) -> List[Dict[str, Any]]:
    ws = get_sheet_if_exists(workbook, "Empirical Model")
    if ws is None:
        print(f"SKIPPED Empirical Model in {source_file}: sheet not found")
        return []

    anchor = find_anchor(ws, "max")
    if anchor is None:
        print(f"SKIPPED Empirical Model in {source_file}: 'max' anchor not found")
        return []
    anchor_row, anchor_col = anchor

    headers = get_header_entries(ws, anchor_row, anchor_col)
    forecast_max_col = anchor_col
    forecast_min_col = find_column(headers, ["min"]) or (anchor_col + 1)
    num_quarters_col = (
        find_column(headers, ["num", "quarter"])
        or find_column(headers, ["quarters", "used"])
        or find_column(headers, ["quarter", "used"])
        or find_column(headers, ["quarter"], ["last", "sales"])
        or (anchor_col - 5)
    )
    last_quarter_col = find_column(headers, ["last", "quarter"]) or (anchor_col - 4)
    forecast_col = (
        find_column(headers, ["estimated", "total", "sold"])
        or find_column(headers, ["forecast"], ["max", "min"])
        or (anchor_col - 2)
    )
    actual_col = find_column(headers, ["reported", "sales"]) or (anchor_col - 1)
    avg_pen_col = find_column(headers, ["avg", "penetration"]) or (anchor_col - 3)
    quarterly_sales_col = find_column(headers, ["quarterly", "sales"]) or (anchor_col - 11)
    reported_sales_col = find_column(headers, ["reported", "sales"]) or (anchor_col - 10)
    growth_rate_col = find_column(headers, ["growth"]) or (anchor_col + 2)
    sales_captured_col = (
        find_column(headers, ["sales", "captured"], ["reported"]) or (anchor_col + 3)
    )

    last_used = ws.used_range.last_cell
    helper_row = max(int(last_used.row) + 3, anchor_row + N_QUARTERS + 3)
    helper_col = max(int(last_used.column) + 2, anchor_col + 8)
    avg_formula_cell = ws.cells(helper_row, helper_col)
    hist_end_row = anchor_row - 1

    rows: List[Dict[str, Any]] = []
    for i in range(N_QUARTERS):
        table_row = anchor_row + 1 + i
        num_quarters_used = to_int(ws.cells(table_row, num_quarters_col).value) or (i + 1)
        hist_start_row = hist_end_row - num_quarters_used + 1

        avg_penetration = to_float(ws.cells(table_row, avg_pen_col).value)
        if (
            hist_start_row >= 1
            and quarterly_sales_col >= 1
            and reported_sales_col >= 1
            and hist_start_row <= hist_end_row
        ):
            avg_formula = (
                "=IFERROR("
                "AVERAGE(IFERROR("
                f"R{hist_start_row}C{quarterly_sales_col}:R{hist_end_row}C{quarterly_sales_col}/"
                f"R{hist_start_row}C{reported_sales_col}:R{hist_end_row}C{reported_sales_col},"
                '""'
                ")),"
                '""'
                ")"
            )
            set_formula2(avg_formula_cell, avg_formula)
            workbook.app.calculate()
            calculated_avg = to_float(avg_formula_cell.value)
            if calculated_avg is not None:
                avg_penetration = calculated_avg

        forecast_value = to_float(ws.cells(table_row, forecast_col).value)
        actual_value = to_float(ws.cells(table_row, actual_col).value)
        forecast_max = to_float(ws.cells(table_row, forecast_max_col).value)
        forecast_min = to_float(ws.cells(table_row, forecast_min_col).value)
        quarterly_sales = to_float(ws.cells(table_row, quarterly_sales_col).value)
        reported_sales = to_float(ws.cells(table_row, reported_sales_col).value)
        growth_rate_pct = to_float(ws.cells(table_row, growth_rate_col).value)
        sales_captured_pct = to_float(ws.cells(table_row, sales_captured_col).value)
        last_quarter_used = ws.cells(table_row, last_quarter_col).value

        if is_effectively_empty(
            [
                forecast_value,
                actual_value,
                forecast_max,
                forecast_min,
                avg_penetration,
                quarterly_sales,
                reported_sales,
            ]
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
                "parameter_value": avg_penetration,
                "num_quarters_used": num_quarters_used,
                "last_quarter_used": last_quarter_used,
                "forecast_value": forecast_value,
                "actual_value": actual_value,
                "forecast_max": forecast_max,
                "forecast_min": forecast_min,
                "range_width": subtract_or_none(forecast_max, forecast_min),
                "avg_penetration_pct": avg_penetration,
                "quarterly_sales": quarterly_sales,
                "reported_sales": reported_sales,
                "growth_rate_pct": growth_rate_pct,
                "sales_captured_in_db_pct": sales_captured_pct,
                "source_file": source_file,
            }
        )

    return rows


def extract_regression_rows(
    workbook: xw.Book,
    labels: Dict[str, str],
    source_file: str,
) -> List[Dict[str, Any]]:
    ws = get_sheet_if_exists(workbook, "Regression Model")
    if ws is None:
        print(f"SKIPPED Regression Model in {source_file}: sheet not found")
        return []

    anchor = find_anchor(ws, "max")
    if anchor is None:
        print(f"SKIPPED Regression Model in {source_file}: 'max' anchor not found")
        return []
    anchor_row, anchor_col = anchor

    y_col = anchor_col - 7
    x_col = anchor_col - 11

    headers = get_header_entries(ws, anchor_row, anchor_col)
    forecast_max_col = anchor_col
    forecast_min_col = find_column(headers, ["min"]) or (anchor_col + 1)
    num_quarters_col = (
        find_column(headers, ["num", "quarter"])
        or find_column(headers, ["quarters", "used"])
        or find_column(headers, ["quarter"], ["sales", "last"])
        or (anchor_col - 4)
    )
    forecast_total_col = (
        find_column(headers, ["tot", "fcst", "sa"])
        or find_column(headers, ["forecast"], ["max", "min"])
        or (anchor_col - 1)
    )
    actual_col = find_column(headers, ["actual", "sales"])

    last_used = ws.used_range.last_cell
    helper_row = max(int(last_used.row) + 3, anchor_row + N_QUARTERS + 3)
    helper_col = max(int(last_used.column) + 2, anchor_col + 8)
    intercept_cell = ws.cells(helper_row, helper_col)
    slope_cell = ws.cells(helper_row + 1, helper_col)

    hist_end_row = anchor_row - 1
    rows: List[Dict[str, Any]] = []
    previous_signature: Optional[Tuple[Any, ...]] = None

    for i in range(N_QUARTERS):
        table_row = anchor_row + 1 + i
        num_quarters_used = to_int(ws.cells(table_row, num_quarters_col).value) or (i + 1)
        hist_start_row = hist_end_row - num_quarters_used + 1
        if hist_start_row < 1 or x_col < 1 or y_col < 1:
            continue

        intercept_formula = (
            "=IFERROR("
            f"INTERCEPT(R{hist_start_row}C{y_col}:R{hist_end_row}C{y_col},"
            f"R{hist_start_row}C{x_col}:R{hist_end_row}C{x_col}),"
            '""'
            ")"
        )
        slope_formula = (
            "=IFERROR("
            f"SLOPE(R{hist_start_row}C{y_col}:R{hist_end_row}C{y_col},"
            f"R{hist_start_row}C{x_col}:R{hist_end_row}C{x_col}),"
            '""'
            ")"
        )
        set_formula2(intercept_cell, intercept_formula)
        set_formula2(slope_cell, slope_formula)
        workbook.app.calculate()

        intercept = to_float(intercept_cell.value)
        slope = to_float(slope_cell.value)
        forecast_value = to_float(ws.cells(table_row, forecast_total_col).value)
        forecast_max = to_float(ws.cells(table_row, forecast_max_col).value)
        forecast_min = to_float(ws.cells(table_row, forecast_min_col).value)
        actual_value = to_float(ws.cells(table_row, actual_col).value) if actual_col else None

        if is_effectively_empty(
            [num_quarters_used, intercept, slope, forecast_value, forecast_max, forecast_min]
        ):
            continue

        signature = (
            num_quarters_used,
            round(intercept, 12) if intercept is not None else None,
            round(slope, 12) if slope is not None else None,
            round(forecast_value, 12) if forecast_value is not None else None,
            round(forecast_max, 12) if forecast_max is not None else None,
            round(forecast_min, 12) if forecast_min is not None else None,
        )
        if previous_signature == signature:
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
                "parameter_value": num_quarters_used,
                "num_quarters_used": num_quarters_used,
                "forecast_value": forecast_value,
                "actual_value": actual_value,
                "forecast_max": forecast_max,
                "forecast_min": forecast_min,
                "range_width": subtract_or_none(forecast_max, forecast_min),
                "intercept": intercept,
                "slope": slope,
                "source_file": source_file,
            }
        )

    return rows


def write_sheet(
    wb: Workbook,
    sheet_name: str,
    columns: List[str],
    rows: List[Dict[str, Any]],
) -> None:
    ws = wb.create_sheet(sheet_name)
    ws.append(columns)
    for row in rows:
        ws.append([row.get(column, "") for column in columns])

    for cell in ws[1]:
        cell.font = Font(bold=True)

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions

    for idx, column_name in enumerate(columns, start=1):
        col_letter = get_column_letter(idx)
        max_len = len(column_name)
        for cell in ws[col_letter]:
            if cell.value is None:
                continue
            max_len = max(max_len, len(str(cell.value)))
        ws.column_dimensions[col_letter].width = min(max(max_len + 2, 12), 42)


def write_output_workbook(
    output_path: Path,
    empirical_rows: List[Dict[str, Any]],
    regression_rows: List[Dict[str, Any]],
) -> None:
    workbook = Workbook()
    default_sheet = workbook.active
    workbook.remove(default_sheet)

    write_sheet(workbook, "empirical_candidates", EMPIRICAL_COLUMNS, empirical_rows)
    write_sheet(workbook, "regression_candidates", REGRESSION_COLUMNS, regression_rows)
    workbook.save(output_path)


def process_workbooks() -> None:
    in_dir = Path(input_dir)
    out_dir = Path(output_dir)

    if not in_dir.exists() or not in_dir.is_dir():
        raise FileNotFoundError(f"Input directory does not exist: {in_dir}")

    output_path = get_output_path(in_dir, out_dir)

    empirical_rows: List[Dict[str, Any]] = []
    regression_rows: List[Dict[str, Any]] = []
    processed_file_count = 0

    app = xw.App(visible=False, add_book=False)
    app.display_alerts = False
    app.screen_updating = False

    try:
        for file_path in sorted(in_dir.iterdir(), key=lambda path: path.name.lower()):
            if not file_path.is_file():
                print(f"SKIPPED {file_path.name}: not a file")
                continue
            if file_path.name.startswith("~"):
                print(f"SKIPPED {file_path.name}: temporary Excel file")
                continue
            if file_path.suffix.lower() != ".xlsx":
                print(f"SKIPPED {file_path.name}: not an .xlsx file")
                continue

            print(f"PROCESSING {file_path.name}")
            labels = parse_file_labels(file_path)
            source_wb: Optional[xw.Book] = None

            try:
                source_wb = app.books.open(str(file_path), update_links=False)
                empirical_rows.extend(
                    extract_empirical_rows(source_wb, labels=labels, source_file=file_path.name)
                )
                regression_rows.extend(
                    extract_regression_rows(source_wb, labels=labels, source_file=file_path.name)
                )
                processed_file_count += 1
            except Exception as exc:
                print(f"SKIPPED {file_path.name}: processing error ({exc})")
            finally:
                if source_wb is not None:
                    safe_close_workbook(source_wb)
    finally:
        app.quit()

    write_output_workbook(output_path, empirical_rows, regression_rows)
    print(f"OUTPUT {output_path}")
    print(f"FILES_PROCESSED {processed_file_count}")
    print(f"EMPIRICAL_ROWS {len(empirical_rows)}")
    print(f"REGRESSION_ROWS {len(regression_rows)}")


if __name__ == "__main__":
    process_workbooks()
