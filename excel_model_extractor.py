from __future__ import annotations

import math
import re
from datetime import date
from pathlib import Path
from statistics import stdev
from typing import Any, Dict, List, Optional, Sequence, Tuple

import xlwings as xw
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

# --- User-configurable paths ---
input_dir = r"/path/to/input"
output_dir = r"/path/to/output"


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

MONTHS = {
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

DAY_BY_PHASE = {"early": 5, "mid": 15, "late": 25}


def to_2d(values: Any) -> List[List[Any]]:
    if values is None:
        return []
    if not isinstance(values, list):
        return [[values]]
    if values and not isinstance(values[0], list):
        return [values]
    return values


def normalize_text(value: Any) -> str:
    if value is None:
        return ""
    text = str(value).strip().lower()
    text = re.sub(r"[\s_\-\/]+", " ", text)
    return text


def as_float(value: Any) -> Optional[float]:
    if value is None or value == "":
        return None
    if isinstance(value, (int, float)):
        if isinstance(value, float) and math.isnan(value):
            return None
        return float(value)
    try:
        return float(str(value).replace(",", "").strip())
    except Exception:
        return None


class SheetSnapshot:
    def __init__(self, sheet: xw.Sheet):
        self.sheet = sheet
        used = sheet.used_range
        self.row0 = used.row
        self.col0 = used.column
        self.values = to_2d(used.options(ndim=2).value)
        self.nrows = len(self.values)
        self.ncols = max((len(row) for row in self.values), default=0)

    def get(self, row: int, col: int) -> Any:
        r = row - self.row0
        c = col - self.col0
        if r < 0 or c < 0 or r >= self.nrows:
            return None
        row_vals = self.values[r]
        if c >= len(row_vals):
            return None
        return row_vals[c]

    def iter_cells(self) -> Sequence[Tuple[int, int, Any]]:
        for r_off, row_vals in enumerate(self.values):
            for c_off, value in enumerate(row_vals):
                yield self.row0 + r_off, self.col0 + c_off, value

    def find_text(self, target: str) -> Optional[Tuple[int, int]]:
        target_norm = normalize_text(target)
        for row, col, value in self.iter_cells():
            if normalize_text(value) == target_norm:
                return row, col
        return None


def set_r1c1_formula2(cell: xw.Range, formula_r1c1: str) -> None:
    # Prefer formula2 as requested; use API fallback for hosts that require Formula2R1C1.
    try:
        cell.formula2 = formula_r1c1
        return
    except Exception:
        pass

    try:
        cell.api.Formula2R1C1 = formula_r1c1
        return
    except Exception:
        pass

    cell.api.FormulaR1C1 = formula_r1c1


def safe_close_workbook(wb: xw.Book) -> None:
    try:
        wb.close(save=False)
        return
    except TypeError:
        pass
    except Exception:
        pass

    try:
        wb.close(False)
        return
    except Exception:
        pass

    try:
        wb.api.Close(SaveChanges=False)
    except Exception:
        pass


def parse_file_labels(file_name: str) -> Dict[str, str]:
    stem = Path(file_name).stem
    parts = [part.strip() for part in stem.split(" - ")]

    ticker = ""
    if len(parts) >= 2:
        ticker = parts[1].strip().upper()
    else:
        ticker_match = re.search(r"\b([A-Z]{2,8})\b", stem)
        if ticker_match:
            ticker = ticker_match.group(1)

    period_match = re.search(
        r"\b(Early|Mid|Late)\s*([A-Za-z]{3,4})\s*(20\d{2})\b", stem, flags=re.IGNORECASE
    )

    model_period = ""
    model_date = ""
    if period_match:
        phase_raw = period_match.group(1)
        month_raw = period_match.group(2)
        year = int(period_match.group(3))

        month_key = month_raw[:3].lower()
        month_num = MONTHS.get(month_key)
        if month_num is not None:
            phase = phase_raw.title()
            month_abbrev = month_raw[:3].title()
            model_period = f"{phase}{month_abbrev}_{year}"

            day = DAY_BY_PHASE[phase.lower()]
            model_date = date(year, month_num, day).isoformat()

    if not model_period:
        model_period = "unknown_period"
    if not model_date:
        model_date = ""
    if not ticker:
        ticker = "UNKNOWN"

    model = f"{ticker}_{model_period}"
    return {
        "model": model,
        "ticker": ticker,
        "model_period": model_period,
        "model_date": model_date,
    }


def next_output_path(input_path: Path, output_path: Path) -> Path:
    folder_name = input_path.name
    base = output_path / f"{folder_name}_PARAM.xlsx"
    if not base.exists():
        return base

    idx = 1
    while True:
        candidate = output_path / f"{folder_name}_PARAM.{idx}.xlsx"
        if not candidate.exists():
            return candidate
        idx += 1


def is_generated_output_name(name: str) -> bool:
    return re.search(r"_PARAM(?:\.\d+)?\.xlsx$", name, flags=re.IGNORECASE) is not None


def collect_source_files(folder: Path) -> List[Path]:
    files: List[Path] = []
    for item in sorted(folder.iterdir()):
        if not item.is_file():
            print(f"Skipped {item.name}: not a file")
            continue
        if item.name.startswith("~"):
            print(f"Skipped {item.name}: temp file")
            continue
        if item.suffix.lower() != ".xlsx":
            print(f"Skipped {item.name}: not an .xlsx file")
            continue
        if is_generated_output_name(item.name):
            print(f"Skipped {item.name}: generated output file pattern")
            continue
        files.append(item)
    return files


def find_nearest_header_column(
    snapshot: SheetSnapshot,
    anchor_row: int,
    anchor_col: int,
    include_tokens: Sequence[str],
    exclude_tokens: Optional[Sequence[str]] = None,
) -> Optional[int]:
    best_col = None
    best_score = None
    exclude_tokens = exclude_tokens or []

    for row, col, value in snapshot.iter_cells():
        text = normalize_text(value)
        if not text:
            continue
        if any(token not in text for token in include_tokens):
            continue
        if any(token in text for token in exclude_tokens):
            continue

        score = abs(col - anchor_col) + (2 * abs(row - anchor_row))
        if best_score is None or score < best_score:
            best_score = score
            best_col = col

    return best_col


def extract_empirical_rows(
    wb: xw.Book,
    file_path: Path,
    labels: Dict[str, str],
) -> List[Dict[str, Any]]:
    if "Empirical Model" not in [s.name for s in wb.sheets]:
        print(f"Skipped empirical in {file_path.name}: missing 'Empirical Model' sheet")
        return []

    sheet = wb.sheets["Empirical Model"]
    snapshot = SheetSnapshot(sheet)
    anchor = snapshot.find_text("max")
    if anchor is None:
        print(f"Skipped empirical in {file_path.name}: no 'max' anchor found")
        return []

    anchor_row, anchor_col = anchor

    quarterly_col = find_nearest_header_column(
        snapshot, anchor_row, anchor_col, include_tokens=["quarterly", "sales"], exclude_tokens=["reported"]
    )
    reported_col = find_nearest_header_column(
        snapshot, anchor_row, anchor_col, include_tokens=["reported", "sales"]
    )
    penetration_col = find_nearest_header_column(
        snapshot, anchor_row, anchor_col, include_tokens=["penetration"]
    )
    growth_col = find_nearest_header_column(
        snapshot, anchor_row, anchor_col, include_tokens=["growth", "rate"]
    )
    captured_col = find_nearest_header_column(
        snapshot, anchor_row, anchor_col, include_tokens=["captured", "db"]
    )
    quarter_label_col = find_nearest_header_column(
        snapshot, anchor_row, anchor_col, include_tokens=["quarter"]
    )

    if quarterly_col is None or reported_col is None:
        print(
            f"Skipped empirical in {file_path.name}: could not find quarterly/reported sales columns near anchor"
        )
        return []

    historical_rows: List[int] = []
    for row in range(1, anchor_row):
        quarterly_value = as_float(snapshot.get(row, quarterly_col))
        reported_value = as_float(snapshot.get(row, reported_col))
        if quarterly_value is None or reported_value is None:
            continue
        if reported_value == 0:
            continue
        historical_rows.append(row)

    if len(historical_rows) < 2:
        print(f"Skipped empirical in {file_path.name}: not enough historical rows")
        return []

    # Use the latest row as the "actual quarter", and prior rows for the rolling averages.
    actual_row = historical_rows[-1]
    training_rows = historical_rows[:-1]
    if not training_rows:
        print(f"Skipped empirical in {file_path.name}: not enough training rows")
        return []

    current_quarterly_sales = as_float(snapshot.get(actual_row, quarterly_col))
    current_reported_sales = as_float(snapshot.get(actual_row, reported_col))
    if current_quarterly_sales is None or current_reported_sales is None or current_reported_sales == 0:
        print(f"Skipped empirical in {file_path.name}: invalid current quarter values")
        return []

    max_n = min(N_QUARTERS, len(training_rows))

    temp_col_base = max(anchor_col + 30, snapshot.col0 + snapshot.ncols + 5)
    penetration_work_col = penetration_col if penetration_col is not None else temp_col_base
    avg_work_col = temp_col_base + 1

    # If the sheet lacks a penetration column, create a temporary one from quarterly/reported.
    if penetration_col is None:
        for row in training_rows:
            formula = f"=IFERROR(RC{quarterly_col}/RC{reported_col},\"\")"
            set_r1c1_formula2(sheet.cells(row, penetration_work_col), formula)
        wb.app.calculate()

    # Write all rolling-average formulas first, then calculate once.
    for idx, n in enumerate(range(1, max_n + 1), start=1):
        out_row = anchor_row + idx
        start_row = training_rows[-n]
        end_row = training_rows[-1]
        avg_formula = f"=AVERAGE(R{start_row}C{penetration_work_col}:R{end_row}C{penetration_work_col})"
        set_r1c1_formula2(sheet.cells(out_row, avg_work_col), avg_formula)
    wb.app.calculate()

    rows: List[Dict[str, Any]] = []
    for idx, n in enumerate(range(1, max_n + 1), start=1):
        avg_penetration = as_float(sheet.cells(anchor_row + idx, avg_work_col).value)
        if avg_penetration is None or avg_penetration == 0:
            continue

        pen_values: List[float] = []
        for r in training_rows[-n:]:
            value = as_float(sheet.cells(r, penetration_work_col).value)
            if value is not None and value > 0:
                pen_values.append(value)
        if not pen_values:
            continue

        min_pen = min(pen_values)
        max_pen = max(pen_values)
        if min_pen <= 0 or max_pen <= 0:
            continue

        forecast_value = current_quarterly_sales / avg_penetration
        forecast_max = current_quarterly_sales / min_pen
        forecast_min = current_quarterly_sales / max_pen
        range_width = forecast_max - forecast_min

        last_quarter_used = ""
        if quarter_label_col is not None:
            last_quarter_used = snapshot.get(training_rows[-1], quarter_label_col) or ""

        growth_rate = as_float(snapshot.get(actual_row, growth_col)) if growth_col is not None else None
        sales_captured_in_db_pct = (
            as_float(snapshot.get(actual_row, captured_col))
            if captured_col is not None
            else (current_quarterly_sales / current_reported_sales)
        )

        rows.append(
            {
                "model": labels["model"],
                "ticker": labels["ticker"],
                "model_period": labels["model_period"],
                "model_date": labels["model_date"],
                "method": "empirical",
                "parameter_name": "avg_penetration_pct",
                "parameter_value": avg_penetration,
                "num_quarters_used": n,
                "last_quarter_used": last_quarter_used,
                "forecast_value": forecast_value,
                "actual_value": current_reported_sales,
                "forecast_max": forecast_max,
                "forecast_min": forecast_min,
                "range_width": range_width,
                "avg_penetration_pct": avg_penetration,
                "quarterly_sales": current_quarterly_sales,
                "reported_sales": current_reported_sales,
                "growth_rate_pct": growth_rate,
                "sales_captured_in_db_pct": sales_captured_in_db_pct,
                "source_file": file_path.name,
            }
        )

    return rows


def extract_regression_rows(
    wb: xw.Book,
    file_path: Path,
    labels: Dict[str, str],
) -> List[Dict[str, Any]]:
    if "Regression Model" not in [s.name for s in wb.sheets]:
        print(f"Skipped regression in {file_path.name}: missing 'Regression Model' sheet")
        return []

    sheet = wb.sheets["Regression Model"]
    snapshot = SheetSnapshot(sheet)
    anchor = snapshot.find_text("max")
    if anchor is None:
        print(f"Skipped regression in {file_path.name}: no 'max' anchor found")
        return []

    anchor_row, anchor_col = anchor
    y_col = anchor_col - 7
    x_col = anchor_col - 11

    series: List[Tuple[int, float, float]] = []
    for row in range(1, anchor_row):
        x_val = as_float(snapshot.get(row, x_col))
        y_val = as_float(snapshot.get(row, y_col))
        if x_val is None or y_val is None:
            continue
        series.append((row, x_val, y_val))

    if len(series) < 2:
        print(f"Skipped regression in {file_path.name}: not enough x/y points")
        return []

    max_n = min(N_QUARTERS, len(series))
    if max_n < 2:
        return []

    temp_col_base = max(anchor_col + 30, snapshot.col0 + snapshot.ncols + 5)

    # Write all INTERCEPT/SLOPE formulas in one pass, then calculate once.
    # Uses R1C1 + formula2 as requested.
    num_values: List[int] = []
    for idx, n in enumerate(range(2, max_n + 1), start=1):
        out_row = anchor_row + idx
        start_row = series[-n][0]
        end_row = series[-1][0]

        intercept_formula = (
            f"=INTERCEPT(R{start_row}C{y_col}:R{end_row}C{y_col},R{start_row}C{x_col}:R{end_row}C{x_col})"
        )
        slope_formula = (
            f"=SLOPE(R{start_row}C{y_col}:R{end_row}C{y_col},R{start_row}C{x_col}:R{end_row}C{x_col})"
        )
        set_r1c1_formula2(sheet.cells(out_row, temp_col_base), intercept_formula)
        set_r1c1_formula2(sheet.cells(out_row, temp_col_base + 1), slope_formula)
        num_values.append(n)
    wb.app.calculate()

    rows: List[Dict[str, Any]] = []
    prev_signature: Optional[Tuple[Any, ...]] = None

    next_actual_row = series[-1][0] + 1
    next_actual_value = as_float(snapshot.get(next_actual_row, y_col))

    for idx, n in enumerate(num_values, start=1):
        calc_row = anchor_row + idx
        intercept = as_float(sheet.cells(calc_row, temp_col_base).value)
        slope = as_float(sheet.cells(calc_row, temp_col_base + 1).value)
        if intercept is None or slope is None:
            continue

        window = series[-n:]
        x_vals = [item[1] for item in window]
        y_vals = [item[2] for item in window]
        x_next = max(x_vals) + 1

        forecast_total_without_sa = intercept + (slope * x_next)
        residuals = [y - (intercept + slope * x) for x, y in zip(x_vals, y_vals)]
        sigma = stdev(residuals) if len(residuals) > 1 else 0.0

        forecast_max = forecast_total_without_sa + sigma
        forecast_min = forecast_total_without_sa - sigma
        range_width = forecast_max - forecast_min

        signature = (
            round(intercept, 10),
            round(slope, 10),
            round(forecast_total_without_sa, 10),
            round(forecast_max, 10),
            round(forecast_min, 10),
        )
        if prev_signature is not None and signature == prev_signature:
            continue
        prev_signature = signature

        rows.append(
            {
                "model": labels["model"],
                "ticker": labels["ticker"],
                "model_period": labels["model_period"],
                "model_date": labels["model_date"],
                "method": "regression",
                "parameter_name": "num_quarters_used",
                "parameter_value": n,
                "num_quarters_used": n,
                "forecast_value": forecast_total_without_sa,
                "actual_value": next_actual_value,
                "forecast_max": forecast_max,
                "forecast_min": forecast_min,
                "range_width": range_width,
                "intercept": intercept,
                "slope": slope,
                "source_file": file_path.name,
            }
        )

    return rows


def write_sheet(ws, columns: List[str], rows: List[Dict[str, Any]]) -> None:
    ws.append(columns)
    for col_idx, _ in enumerate(columns, start=1):
        ws.cell(row=1, column=col_idx).font = Font(bold=True)

    for row_data in rows:
        ws.append([row_data.get(col) for col in columns])

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions

    for idx, col_name in enumerate(columns, start=1):
        max_len = len(col_name)
        for row_idx in range(2, ws.max_row + 1):
            value = ws.cell(row=row_idx, column=idx).value
            if value is None:
                continue
            max_len = max(max_len, len(str(value)))
        ws.column_dimensions[get_column_letter(idx)].width = min(max_len + 2, 50)


def write_output_workbook(
    output_path: Path,
    empirical_rows: List[Dict[str, Any]],
    regression_rows: List[Dict[str, Any]],
) -> None:
    wb = Workbook()
    ws_empirical = wb.active
    ws_empirical.title = "empirical_candidates"
    ws_regression = wb.create_sheet("regression_candidates")

    write_sheet(ws_empirical, EMPIRICAL_COLUMNS, empirical_rows)
    write_sheet(ws_regression, REGRESSION_COLUMNS, regression_rows)
    wb.save(output_path)


def main() -> None:
    in_path = Path(input_dir).expanduser().resolve()
    out_path = Path(output_dir).expanduser().resolve()

    if not in_path.exists() or not in_path.is_dir():
        raise FileNotFoundError(f"input_dir does not exist or is not a folder: {in_path}")

    out_path.mkdir(parents=True, exist_ok=True)
    output_file = next_output_path(in_path, out_path)

    source_files = collect_source_files(in_path)

    empirical_rows: List[Dict[str, Any]] = []
    regression_rows: List[Dict[str, Any]] = []
    processed_count = 0

    app = xw.App(visible=False, add_book=False)
    app.display_alerts = False
    app.screen_updating = False

    try:
        for file_path in source_files:
            print(f"Processing {file_path.name}")
            wb: Optional[xw.Book] = None
            try:
                wb = app.books.open(str(file_path), update_links=False)
                labels = parse_file_labels(file_path.name)

                empirical = extract_empirical_rows(wb, file_path, labels)
                regression = extract_regression_rows(wb, file_path, labels)

                empirical_rows.extend(empirical)
                regression_rows.extend(regression)
                processed_count += 1
            except Exception as exc:
                print(f"Skipped {file_path.name}: processing error -> {exc}")
            finally:
                if wb is not None:
                    safe_close_workbook(wb)
    finally:
        app.quit()

    write_output_workbook(output_file, empirical_rows, regression_rows)

    print(f"Output path: {output_file}")
    print(f"Number of files processed: {processed_count}")
    print(f"Number of empirical rows: {len(empirical_rows)}")
    print(f"Number of regression rows: {len(regression_rows)}")


if __name__ == "__main__":
    main()
