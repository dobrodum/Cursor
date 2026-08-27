from __future__ import annotations

import re
from dataclasses import dataclass
from datetime import date
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Tuple

import xlwings as xw
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

# ==============================
# User-configurable input/output
# ==============================
input_dir = Path("/path/to/input")
output_dir = Path("/path/to/output")


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

PHASE_DAY = {"early": 5, "mid": 15, "late": 25}
MONTH_NUM = {
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
N_QUARTERS = 10


@dataclass
class FileLabels:
    model: str
    ticker: str
    model_period: str
    model_date: str


def parse_file_labels(file_path: Path) -> FileLabels:
    stem = file_path.stem

    ticker = "UNKNOWN"
    parts = [p.strip() for p in stem.split("-")]
    if len(parts) >= 2 and parts[1]:
        ticker = re.sub(r"\W+", "", parts[1]).upper() or "UNKNOWN"

    period_match = re.search(
        r"(Early|Mid|Late)\s*([A-Za-z]{3})\s*(\d{4})",
        stem,
        flags=re.IGNORECASE,
    )

    if period_match:
        phase_raw, month_raw, year_raw = period_match.groups()
        phase = phase_raw.title()
        month_abb = month_raw[:3].title()
        month_key = month_abb.lower()
        year = int(year_raw)

        day = PHASE_DAY.get(phase.lower(), 15)
        month_num = MONTH_NUM.get(month_key, 1)

        model_period = f"{phase}{month_abb}_{year}"
        model_date = date(year, month_num, day).isoformat()
    else:
        model_period = "UnknownPeriod"
        model_date = ""

    model = f"{ticker}_{model_period}" if ticker else model_period
    return FileLabels(model=model, ticker=ticker, model_period=model_period, model_date=model_date)


def is_temp_file(file_path: Path) -> bool:
    return file_path.name.startswith("~")


def unique_output_path(input_folder: Path, out_folder: Path) -> Path:
    base = f"{input_folder.name}_PARAM.xlsx"
    candidate = out_folder / base
    if not candidate.exists():
        return candidate

    idx = 1
    while True:
        candidate = out_folder / f"{input_folder.name}_PARAM.{idx}.xlsx"
        if not candidate.exists():
            return candidate
        idx += 1


def normalize_text(value: Any) -> str:
    if value is None:
        return ""
    text = str(value).strip().lower()
    text = re.sub(r"\s+", " ", text)
    return text


def to_float(value: Any) -> Optional[float]:
    if value is None or value == "":
        return None
    if isinstance(value, bool):
        return float(value)
    if isinstance(value, (int, float)):
        return float(value)
    text = str(value).strip().replace(",", "")
    if text.endswith("%"):
        try:
            return float(text.rstrip("%")) / 100.0
        except ValueError:
            return None
    try:
        return float(text)
    except ValueError:
        return None


def safe_close_workbook(wb: Any) -> None:
    if wb is None:
        return

    try:
        wb.close(save=False)
        return
    except TypeError:
        pass
    except Exception:
        pass

    # Fallbacks for different Excel engines / API wrappers.
    close_attempts = [
        lambda: wb.api.Close(SaveChanges=False),
        lambda: wb.api.Close(False),
        lambda: wb.close(),
    ]
    for close_fn in close_attempts:
        try:
            close_fn()
            return
        except Exception:
            continue


def set_formula2(cell: Any, formula_r1c1: str) -> None:
    try:
        cell.formula2 = formula_r1c1
    except Exception:
        # Keep a fallback for older Excel wrappers.
        cell.formula = formula_r1c1


def find_anchor(sheet: Any, anchor_text: str = "max") -> Tuple[int, int]:
    used = sheet.used_range
    values = used.value

    if values is None:
        raise ValueError(f"Sheet '{sheet.name}' is empty.")
    if not isinstance(values, list):
        values = [[values]]
    elif values and not isinstance(values[0], list):
        values = [values]

    for r_idx, row in enumerate(values):
        for c_idx, val in enumerate(row):
            if normalize_text(val) == anchor_text:
                return used.row + r_idx, used.column + c_idx

    raise ValueError(f"Could not find '{anchor_text}' anchor on sheet '{sheet.name}'.")


def build_header_offsets(sheet: Any, anchor_row: int, anchor_col: int, window: int = 18) -> Dict[str, int]:
    start_col = max(1, anchor_col - window)
    end_col = anchor_col + window
    row_values = sheet.range((anchor_row, start_col), (anchor_row, end_col)).value

    if not isinstance(row_values, list):
        row_values = [row_values]

    offsets: Dict[str, int] = {}
    for idx, raw in enumerate(row_values):
        header = normalize_text(raw)
        if not header:
            continue
        offsets[header] = (start_col + idx) - anchor_col
    return offsets


def find_offset_by_keywords(
    header_offsets: Dict[str, int],
    keywords: Iterable[str],
    default: Optional[int] = None,
) -> Optional[int]:
    for header, offset in header_offsets.items():
        if all(k in header for k in keywords):
            return offset
    return default


def read_cell(sheet: Any, row: int, col: int) -> Any:
    return sheet.cells(row, col).value


def process_empirical_sheet(wb: Any, labels: FileLabels, source_file: str) -> List[Dict[str, Any]]:
    try:
        sheet = wb.sheets["Empirical Model"]
    except Exception:
        return []

    anchor_row, anchor_col = find_anchor(sheet, "max")
    header_offsets = build_header_offsets(sheet, anchor_row, anchor_col)
    data_start = anchor_row + 1
    data_end_hist = anchor_row - 1

    offset_num_q = find_offset_by_keywords(header_offsets, ("quarter", "used"), default=-8)
    offset_last_q = find_offset_by_keywords(header_offsets, ("last", "quarter"), default=-7)
    offset_avg_pen = find_offset_by_keywords(header_offsets, ("avg", "penetration"), default=-6)
    offset_fcst = find_offset_by_keywords(header_offsets, ("estimated", "sold"), default=-3)
    if offset_fcst is None:
        offset_fcst = find_offset_by_keywords(header_offsets, ("forecast",), default=-3)
    offset_actual = find_offset_by_keywords(header_offsets, ("reported", "sales"), default=-2)
    offset_max = 0
    offset_min = find_offset_by_keywords(header_offsets, ("min",), default=1)
    offset_q_sales = find_offset_by_keywords(header_offsets, ("quarterly", "sales"), default=-5)
    offset_rep_sales = find_offset_by_keywords(header_offsets, ("reported", "sales"), default=-2)
    offset_growth = find_offset_by_keywords(header_offsets, ("growth",), default=-4)
    offset_capture = find_offset_by_keywords(header_offsets, ("captured", "db"), default=-1)
    pen_source_offset = find_offset_by_keywords(header_offsets, ("penetration",), default=offset_avg_pen)

    formula_rows: List[int] = []
    if offset_avg_pen is not None and pen_source_offset is not None and data_end_hist >= 1:
        pen_col = anchor_col + pen_source_offset
        avg_col = anchor_col + offset_avg_pen

        for n in range(1, N_QUARTERS + 1):
            out_row = data_start + (n - 1)
            start_row = max(1, data_end_hist - n + 1)
            formula = f"=AVERAGE(R{start_row}C{pen_col}:R{data_end_hist}C{pen_col})"
            set_formula2(sheet.cells(out_row, avg_col), formula)
            formula_rows.append(out_row)

        if formula_rows:
            wb.app.calculate()

    rows: List[Dict[str, Any]] = []
    for n in range(1, N_QUARTERS + 1):
        row_idx = data_start + (n - 1)

        num_quarters = read_cell(sheet, row_idx, anchor_col + offset_num_q) if offset_num_q is not None else n
        if num_quarters in (None, ""):
            num_quarters = n

        last_quarter = read_cell(sheet, row_idx, anchor_col + offset_last_q) if offset_last_q is not None else None
        avg_pen = read_cell(sheet, row_idx, anchor_col + offset_avg_pen) if offset_avg_pen is not None else None
        forecast_value = read_cell(sheet, row_idx, anchor_col + offset_fcst) if offset_fcst is not None else None
        actual_value = read_cell(sheet, row_idx, anchor_col + offset_actual) if offset_actual is not None else None
        forecast_max = read_cell(sheet, row_idx, anchor_col + offset_max)
        forecast_min = read_cell(sheet, row_idx, anchor_col + offset_min) if offset_min is not None else None
        quarterly_sales = read_cell(sheet, row_idx, anchor_col + offset_q_sales) if offset_q_sales is not None else None
        reported_sales = read_cell(sheet, row_idx, anchor_col + offset_rep_sales) if offset_rep_sales is not None else actual_value
        growth_rate = read_cell(sheet, row_idx, anchor_col + offset_growth) if offset_growth is not None else None
        captured_pct = read_cell(sheet, row_idx, anchor_col + offset_capture) if offset_capture is not None else None

        has_payload = any(
            x not in (None, "")
            for x in (avg_pen, forecast_value, forecast_max, forecast_min, actual_value)
        )
        if not has_payload:
            continue

        max_f = to_float(forecast_max)
        min_f = to_float(forecast_min)
        range_width = max_f - min_f if max_f is not None and min_f is not None else None

        rows.append(
            {
                "model": labels.model,
                "ticker": labels.ticker,
                "model_period": labels.model_period,
                "model_date": labels.model_date,
                "method": "empirical",
                "parameter_name": "avg_penetration_pct",
                "parameter_value": avg_pen,
                "num_quarters_used": num_quarters,
                "last_quarter_used": last_quarter,
                "forecast_value": forecast_value,
                "actual_value": actual_value,
                "forecast_max": forecast_max,
                "forecast_min": forecast_min,
                "range_width": range_width,
                "avg_penetration_pct": avg_pen,
                "quarterly_sales": quarterly_sales,
                "reported_sales": reported_sales,
                "growth_rate_pct": growth_rate,
                "sales_captured_in_db_pct": captured_pct,
                "source_file": source_file,
            }
        )

    if formula_rows and offset_avg_pen is not None:
        avg_col = anchor_col + offset_avg_pen
        for row_idx in formula_rows:
            try:
                sheet.cells(row_idx, avg_col).value = None
            except Exception:
                pass

    return rows


def process_regression_sheet(wb: Any, labels: FileLabels, source_file: str) -> List[Dict[str, Any]]:
    try:
        sheet = wb.sheets["Regression Model"]
    except Exception:
        return []

    anchor_row, anchor_col = find_anchor(sheet, "max")
    header_offsets = build_header_offsets(sheet, anchor_row, anchor_col)
    data_start = anchor_row + 1

    y_col = anchor_col - 7
    x_col = anchor_col - 11

    offset_num_q = find_offset_by_keywords(header_offsets, ("quarter", "used"), default=-12)
    offset_fcst = find_offset_by_keywords(header_offsets, ("tot", "fcst"), default=-3)
    if offset_fcst is None:
        offset_fcst = find_offset_by_keywords(header_offsets, ("forecast",), default=-3)
    offset_actual = find_offset_by_keywords(header_offsets, ("actual",), default=None)
    offset_max = 0
    offset_min = find_offset_by_keywords(header_offsets, ("min",), default=1)

    # Scratch columns (temporary formulas) so we don't overwrite any business cells.
    scratch_intercept_col = anchor_col + 20
    scratch_slope_col = anchor_col + 21
    hist_end = anchor_row - 1

    if hist_end < 1:
        return []

    for n in range(1, N_QUARTERS + 1):
        row_idx = data_start + (n - 1)
        start_row = max(1, hist_end - n + 1)

        intercept_formula = (
            f"=INTERCEPT(R{start_row}C{y_col}:R{hist_end}C{y_col},"
            f"R{start_row}C{x_col}:R{hist_end}C{x_col})"
        )
        slope_formula = (
            f"=SLOPE(R{start_row}C{y_col}:R{hist_end}C{y_col},"
            f"R{start_row}C{x_col}:R{hist_end}C{x_col})"
        )
        set_formula2(sheet.cells(row_idx, scratch_intercept_col), intercept_formula)
        set_formula2(sheet.cells(row_idx, scratch_slope_col), slope_formula)

    wb.app.calculate()

    rows: List[Dict[str, Any]] = []
    last_signature: Optional[Tuple[Any, ...]] = None

    for n in range(1, N_QUARTERS + 1):
        row_idx = data_start + (n - 1)

        num_quarters = read_cell(sheet, row_idx, anchor_col + offset_num_q) if offset_num_q is not None else n
        if num_quarters in (None, ""):
            num_quarters = n

        forecast_value = read_cell(sheet, row_idx, anchor_col + offset_fcst) if offset_fcst is not None else None
        actual_value = read_cell(sheet, row_idx, anchor_col + offset_actual) if offset_actual is not None else None
        forecast_max = read_cell(sheet, row_idx, anchor_col + offset_max)
        forecast_min = read_cell(sheet, row_idx, anchor_col + offset_min) if offset_min is not None else None
        intercept_val = read_cell(sheet, row_idx, scratch_intercept_col)
        slope_val = read_cell(sheet, row_idx, scratch_slope_col)

        has_payload = any(
            x not in (None, "")
            for x in (forecast_value, forecast_max, forecast_min, intercept_val, slope_val)
        )
        if not has_payload:
            continue

        max_f = to_float(forecast_max)
        min_f = to_float(forecast_min)
        range_width = max_f - min_f if max_f is not None and min_f is not None else None

        signature = (
            str(num_quarters),
            round(to_float(forecast_value) or 0.0, 9),
            round(to_float(forecast_max) or 0.0, 9),
            round(to_float(forecast_min) or 0.0, 9),
            round(to_float(intercept_val) or 0.0, 9),
            round(to_float(slope_val) or 0.0, 9),
        )
        if signature == last_signature:
            continue
        last_signature = signature

        rows.append(
            {
                "model": labels.model,
                "ticker": labels.ticker,
                "model_period": labels.model_period,
                "model_date": labels.model_date,
                "method": "regression",
                "parameter_name": "num_quarters_used",
                "parameter_value": num_quarters,
                "num_quarters_used": num_quarters,
                "forecast_value": forecast_value,
                "actual_value": actual_value if actual_value not in (None, "") else "",
                "forecast_max": forecast_max,
                "forecast_min": forecast_min,
                "range_width": range_width,
                "intercept": intercept_val,
                "slope": slope_val,
                "source_file": source_file,
            }
        )

    # Cleanup scratch formulas.
    for n in range(1, N_QUARTERS + 1):
        row_idx = data_start + (n - 1)
        try:
            sheet.cells(row_idx, scratch_intercept_col).value = None
            sheet.cells(row_idx, scratch_slope_col).value = None
        except Exception:
            pass

    return rows


def write_sheet(ws: Any, columns: List[str], rows: List[Dict[str, Any]]) -> None:
    ws.append(columns)
    for row in rows:
        ws.append([row.get(col, "") for col in columns])

    for cell in ws[1]:
        cell.font = Font(bold=True)

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions

    for col_idx, col_name in enumerate(columns, start=1):
        max_len = len(col_name)
        for r_idx in range(2, ws.max_row + 1):
            value = ws.cell(r_idx, col_idx).value
            if value is None:
                continue
            max_len = max(max_len, len(str(value)))
        width = min(max(12, max_len + 2), 48)
        ws.column_dimensions[get_column_letter(col_idx)].width = width


def create_output_workbook(
    empirical_rows: List[Dict[str, Any]],
    regression_rows: List[Dict[str, Any]],
    output_path: Path,
) -> None:
    wb = Workbook()
    default_sheet = wb.active
    wb.remove(default_sheet)

    ws_emp = wb.create_sheet("empirical_candidates")
    write_sheet(ws_emp, EMPIRICAL_COLUMNS, empirical_rows)

    ws_reg = wb.create_sheet("regression_candidates")
    write_sheet(ws_reg, REGRESSION_COLUMNS, regression_rows)

    wb.save(output_path)


def process_all_files(input_folder: Path, out_folder: Path) -> Tuple[List[Dict[str, Any]], List[Dict[str, Any]], int]:
    out_folder.mkdir(parents=True, exist_ok=True)
    entries = sorted(input_folder.iterdir())

    empirical_rows: List[Dict[str, Any]] = []
    regression_rows: List[Dict[str, Any]] = []
    processed_files = 0

    app = xw.App(visible=False, add_book=False)
    app.display_alerts = False
    app.screen_updating = False

    original_calc = None
    try:
        try:
            original_calc = app.calculation
            app.calculation = "manual"
        except Exception:
            pass

        for file_path in entries:
            if not file_path.is_file():
                print(f"Skipped file: {file_path.name} (not a file)")
                continue

            if file_path.suffix.lower() != ".xlsx":
                print(f"Skipped file: {file_path.name} (not an .xlsx file)")
                continue

            if is_temp_file(file_path):
                print(f"Skipped file: {file_path.name} (temporary file)")
                continue

            wb = None
            try:
                print(f"Processing file: {file_path.name}")
                labels = parse_file_labels(file_path)
                wb = app.books.open(str(file_path), update_links=False)

                emp_rows = process_empirical_sheet(wb, labels, file_path.name)
                reg_rows = process_regression_sheet(wb, labels, file_path.name)

                empirical_rows.extend(emp_rows)
                regression_rows.extend(reg_rows)
                processed_files += 1
            except Exception as exc:
                print(f"Skipped file: {file_path.name} (error: {exc})")
            finally:
                safe_close_workbook(wb)
    finally:
        if original_calc is not None:
            try:
                app.calculation = original_calc
            except Exception:
                pass
        app.quit()

    return empirical_rows, regression_rows, processed_files


def main() -> None:
    if not input_dir.exists():
        raise FileNotFoundError(f"Input directory not found: {input_dir}")

    output_path = unique_output_path(input_dir, output_dir)
    empirical_rows, regression_rows, processed_files = process_all_files(input_dir, output_dir)
    create_output_workbook(empirical_rows, regression_rows, output_path)

    print(f"Output path: {output_path}")
    print(f"Files processed: {processed_files}")
    print(f"Empirical rows: {len(empirical_rows)}")
    print(f"Regression rows: {len(regression_rows)}")


if __name__ == "__main__":
    main()
