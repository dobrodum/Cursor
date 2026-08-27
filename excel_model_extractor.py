from __future__ import annotations

import math
import re
from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path
from statistics import median
from typing import Any, Dict, Iterable, List, Optional, Sequence, Tuple

import xlwings as xw
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter


# =========================
# User-configurable inputs
# =========================
input_dir = r"./input"
output_dir = r"./output"


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


N_QUARTERS = 10
NUMERIC_TOL = 1e-9


@dataclass(frozen=True)
class FileMeta:
    model: str
    ticker: str
    model_period: str
    model_date: str


def normalize_text(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, str):
        return re.sub(r"\s+", " ", value.strip().lower())
    return str(value).strip().lower()


def is_number(value: Any) -> bool:
    if value is None or isinstance(value, bool):
        return False
    if isinstance(value, (int, float)):
        if isinstance(value, float) and (math.isnan(value) or math.isinf(value)):
            return False
        return True
    return False


def clean_excel_value(value: Any) -> Any:
    if isinstance(value, str) and value.startswith("#"):
        return None
    return value


def as_float(value: Any) -> Optional[float]:
    value = clean_excel_value(value)
    if is_number(value):
        return float(value)
    return None


def rows_2d(values: Any) -> List[List[Any]]:
    if values is None:
        return []
    if not isinstance(values, (list, tuple)):
        return [[values]]
    if not values:
        return []
    first = values[0]
    if not isinstance(first, (list, tuple)):
        return [list(values)]
    return [list(row) for row in values]


def make_unique_output_path(input_path: Path, out_dir: Path) -> Path:
    base_name = f"{input_path.name}_PARAM.xlsx"
    candidate = out_dir / base_name
    if not candidate.exists():
        return candidate

    idx = 1
    while True:
        candidate = out_dir / f"{input_path.name}_PARAM.{idx}.xlsx"
        if not candidate.exists():
            return candidate
        idx += 1


def parse_file_meta(file_path: Path) -> FileMeta:
    stem = file_path.stem
    pattern = re.compile(
        r"-\s*(?P<ticker>[A-Za-z0-9]+)\s*-\s*"
        r"(?P<phase>Early|Mid|Late)(?P<month>[A-Za-z]{3})(?P<year>\d{4})",
        re.IGNORECASE,
    )
    match = pattern.search(stem)
    if not match:
        raise ValueError("filename does not match expected model naming convention")

    ticker = match.group("ticker").upper()
    phase_raw = match.group("phase").lower()
    month_raw = match.group("month").title()
    year_raw = match.group("year")

    month_num = datetime.strptime(month_raw, "%b").month
    day_map = {"early": 5, "mid": 15, "late": 25}
    day = day_map[phase_raw]

    phase = phase_raw.capitalize()
    model_period = f"{phase}{month_raw}_{year_raw}"
    model_date = date(int(year_raw), month_num, day).isoformat()
    model = f"{ticker}_{model_period}"

    return FileMeta(model=model, ticker=ticker, model_period=model_period, model_date=model_date)


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
        # Last-resort attempt with positional argument.
        try:
            wb.api.Close(False)
        except Exception:
            pass


def set_r1c1_formula2(cell: xw.Range, formula_r1c1: str) -> None:
    # Prefer .formula2 first; fall back to COM Formula2R1C1 if needed.
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
    cell.formula = formula_r1c1


def scan_sheet(sheet: xw.Sheet) -> Dict[str, Any]:
    used = sheet.used_range
    values = rows_2d(used.value)
    return {
        "values": values,
        "base_row": used.row,
        "base_col": used.column,
        "last_row": used.last_cell.row,
        "last_col": used.last_cell.column,
    }


def scan_value(scan: Dict[str, Any], row: int, col: int) -> Any:
    row_idx = row - scan["base_row"]
    col_idx = col - scan["base_col"]
    values = scan["values"]
    if row_idx < 0 or col_idx < 0:
        return None
    if row_idx >= len(values):
        return None
    row_values = values[row_idx]
    if col_idx >= len(row_values):
        return None
    return row_values[col_idx]


def find_anchor_max(scan: Dict[str, Any]) -> Optional[Tuple[int, int]]:
    candidates: List[Tuple[int, int, int]] = []
    values = scan["values"]
    for r_idx, row in enumerate(values):
        for c_idx, raw in enumerate(row):
            if normalize_text(raw) == "max":
                row_abs = scan["base_row"] + r_idx
                col_abs = scan["base_col"] + c_idx
                score = 0
                if normalize_text(scan_value(scan, row_abs + 1, col_abs)) == "min":
                    score += 3
                for c2 in range(col_abs + 1, col_abs + 5):
                    if is_number(scan_value(scan, row_abs, c2)):
                        score += 1
                        break
                candidates.append((score, row_abs, col_abs))

    if not candidates:
        return None
    candidates.sort(key=lambda item: (-item[0], item[1], item[2]))
    _, row_abs, col_abs = candidates[0]
    return row_abs, col_abs


def nearby_max_min_factors(scan: Dict[str, Any], anchor_row: int, anchor_col: int) -> Tuple[float, float]:
    max_factor = None
    min_factor = None

    for c in range(anchor_col + 1, anchor_col + 6):
        val = as_float(scan_value(scan, anchor_row, c))
        if val is not None:
            max_factor = val
            break

    for r in range(anchor_row - 3, anchor_row + 4):
        for c in range(anchor_col - 3, anchor_col + 4):
            if normalize_text(scan_value(scan, r, c)) == "min":
                for c2 in range(c + 1, c + 6):
                    val = as_float(scan_value(scan, r, c2))
                    if val is not None:
                        min_factor = val
                        break
            if min_factor is not None:
                break
        if min_factor is not None:
            break

    if max_factor is None:
        max_factor = 1.0
    if min_factor is None:
        min_factor = 1.0
    return float(max_factor), float(min_factor)


def numeric_rows_for_columns(
    scan: Dict[str, Any], row_cap: int, col_a: int, col_b: int
) -> List[int]:
    rows: List[int] = []
    start = scan["base_row"]
    end = min(scan["last_row"], row_cap)
    for r in range(start, end + 1):
        a = scan_value(scan, r, col_a)
        b = scan_value(scan, r, col_b)
        if is_number(a) and is_number(b):
            rows.append(r)
    return rows


def values_equal(a: Any, b: Any, tol: float = NUMERIC_TOL) -> bool:
    af = as_float(a)
    bf = as_float(b)
    if af is None and bf is None:
        return clean_excel_value(a) == clean_excel_value(b)
    if af is None or bf is None:
        return False
    return abs(af - bf) <= tol


def infer_empirical_columns(
    scan: Dict[str, Any], anchor_col: int, anchor_row: int
) -> Tuple[int, int, List[int]]:
    capture_candidates = [anchor_col - 11, anchor_col - 10, anchor_col - 9, anchor_col - 8]
    report_candidates = [anchor_col - 7, anchor_col - 6, anchor_col - 5, anchor_col - 4]
    best_score = -1.0
    best_capture = anchor_col - 11
    best_report = anchor_col - 7
    best_rows: List[int] = []

    for capture_col in capture_candidates:
        for report_col in report_candidates:
            if capture_col == report_col:
                continue
            rows = numeric_rows_for_columns(scan, anchor_row, capture_col, report_col)
            if len(rows) < 3:
                continue

            ratios: List[float] = []
            for row_idx in rows[-8:]:
                c_val = as_float(scan_value(scan, row_idx, capture_col))
                r_val = as_float(scan_value(scan, row_idx, report_col))
                if c_val is None or r_val is None or abs(r_val) < NUMERIC_TOL:
                    continue
                ratios.append(abs(c_val / r_val))

            if not ratios:
                continue

            med = median(ratios)
            score = len(rows)
            # Favor plausible penetration ratios.
            if 0.0 < med < 2.0:
                score += 5.0
            elif 2.0 <= med < 5.0:
                score += 1.0

            if score > best_score:
                best_score = score
                best_capture = capture_col
                best_report = report_col
                best_rows = rows

    if not best_rows:
        best_rows = numeric_rows_for_columns(scan, anchor_row, best_capture, best_report)
    return best_capture, best_report, best_rows


def infer_quarter_label_col(scan: Dict[str, Any], capture_col: int, sample_rows: Sequence[int]) -> int:
    candidates = [capture_col - 1, capture_col - 2, capture_col - 3]
    for col in candidates:
        non_empty = 0
        non_numeric = 0
        for row in sample_rows[-5:]:
            value = scan_value(scan, row, col)
            if value not in (None, ""):
                non_empty += 1
                if not is_number(value):
                    non_numeric += 1
        if non_empty > 0 and non_numeric >= max(1, non_empty // 2):
            return col
    return capture_col - 1


def extract_empirical_rows(wb: xw.Book, meta: FileMeta, source_file: str) -> List[Dict[str, Any]]:
    if "Empirical Model" not in [sheet.name for sheet in wb.sheets]:
        print(f"SKIPPED sheet: {source_file} (missing sheet 'Empirical Model')")
        return []

    sheet = wb.sheets["Empirical Model"]
    scan = scan_sheet(sheet)
    anchor = find_anchor_max(scan)
    if anchor is None:
        print(f"SKIPPED sheet: {source_file} (no 'max' anchor in Empirical Model)")
        return []
    anchor_row, anchor_col = anchor

    max_factor, min_factor = nearby_max_min_factors(scan, anchor_row, anchor_col)
    capture_col, report_col, data_rows = infer_empirical_columns(scan, anchor_col, anchor_row)
    if len(data_rows) < 3:
        print(f"SKIPPED sheet: {source_file} (insufficient empirical numeric history)")
        return []

    data_start = data_rows[0]
    data_end = data_rows[-1]
    forecast_row = data_end + 1

    forecast_capture = as_float(sheet.cells(forecast_row, capture_col).value)
    forecast_report = as_float(sheet.cells(forecast_row, report_col).value)
    if forecast_capture is None:
        forecast_row = data_end
        forecast_capture = as_float(sheet.cells(forecast_row, capture_col).value)
        forecast_report = as_float(sheet.cells(forecast_row, report_col).value)

    if forecast_capture is None:
        print(f"SKIPPED sheet: {source_file} (unable to locate empirical forecast sales input)")
        return []

    quarter_col = infer_quarter_label_col(scan, capture_col, data_rows)
    last_quarter_used = scan_value(scan, data_end, quarter_col)
    prev_capture = as_float(sheet.cells(max(data_start, forecast_row - 1), capture_col).value)

    temp_start_row = max(scan["last_row"] + 5, anchor_row + 5)
    temp_col = max(scan["last_col"] + 2, anchor_col + 6)

    calc_specs: List[Tuple[int, int, int]] = []
    for n in range(1, N_QUARTERS + 1):
        start_row = data_end - n + 1
        if start_row < data_start:
            continue
        calc_specs.append((n, start_row, data_end))

    if not calc_specs:
        return []

    for idx, (n, start_row, end_row) in enumerate(calc_specs):
        row = temp_start_row + idx
        avg_cell = sheet.cells(row, temp_col)
        quarterly_cell = sheet.cells(row, temp_col + 1)
        reported_cell = sheet.cells(row, temp_col + 2)
        forecast_cell = sheet.cells(row, temp_col + 3)
        max_cell = sheet.cells(row, temp_col + 4)
        min_cell = sheet.cells(row, temp_col + 5)
        growth_cell = sheet.cells(row, temp_col + 6)
        captured_cell = sheet.cells(row, temp_col + 7)

        set_r1c1_formula2(
            avg_cell,
            f"=AVERAGE(R{start_row}C{capture_col}:R{end_row}C{capture_col}/R{start_row}C{report_col}:R{end_row}C{report_col})",
        )
        set_r1c1_formula2(quarterly_cell, f"=R{forecast_row}C{capture_col}")
        set_r1c1_formula2(reported_cell, f"=R{forecast_row}C{report_col}")
        set_r1c1_formula2(
            forecast_cell,
            f"=IF(R{row}C{temp_col}>0,R{row}C{temp_col+1}/R{row}C{temp_col},NA())",
        )
        set_r1c1_formula2(max_cell, f"=R{row}C{temp_col+3}*{max_factor}")
        set_r1c1_formula2(min_cell, f"=R{row}C{temp_col+3}*{min_factor}")
        if prev_capture is not None and abs(prev_capture) > NUMERIC_TOL:
            set_r1c1_formula2(
                growth_cell,
                f"=IF(R{forecast_row-1}C{capture_col}=0,NA(),R{forecast_row}C{capture_col}/R{forecast_row-1}C{capture_col}-1)",
            )
        else:
            growth_cell.value = None
        set_r1c1_formula2(
            captured_cell,
            f"=IF(R{row}C{temp_col+2}=0,NA(),R{row}C{temp_col+1}/R{row}C{temp_col+2})",
        )

    wb.app.calculate()

    last_row = temp_start_row + len(calc_specs) - 1
    calc_values = rows_2d(
        sheet.range((temp_start_row, temp_col), (last_row, temp_col + 7)).value
    )

    output: List[Dict[str, Any]] = []
    for idx, (n, _, _) in enumerate(calc_specs):
        vals = [clean_excel_value(v) for v in calc_values[idx]]
        avg_pen, quarterly_sales, reported_sales, forecast_value, forecast_max, forecast_min, growth_rate, captured_pct = vals

        range_width = None
        if as_float(forecast_max) is not None and as_float(forecast_min) is not None:
            range_width = as_float(forecast_max) - as_float(forecast_min)

        output.append(
            {
                "model": meta.model,
                "ticker": meta.ticker,
                "model_period": meta.model_period,
                "model_date": meta.model_date,
                "method": "empirical",
                "parameter_name": "avg_penetration_pct",
                "parameter_value": avg_pen,
                "num_quarters_used": n,
                "last_quarter_used": last_quarter_used,
                "forecast_value": forecast_value,
                "actual_value": reported_sales,
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

    return output


def extract_regression_rows(wb: xw.Book, meta: FileMeta, source_file: str) -> List[Dict[str, Any]]:
    if "Regression Model" not in [sheet.name for sheet in wb.sheets]:
        print(f"SKIPPED sheet: {source_file} (missing sheet 'Regression Model')")
        return []

    sheet = wb.sheets["Regression Model"]
    scan = scan_sheet(sheet)
    anchor = find_anchor_max(scan)
    if anchor is None:
        print(f"SKIPPED sheet: {source_file} (no 'max' anchor in Regression Model)")
        return []
    anchor_row, anchor_col = anchor

    y_col = anchor_col - 7
    x_col = anchor_col - 11
    data_rows = numeric_rows_for_columns(scan, anchor_row, x_col, y_col)
    if len(data_rows) < 3:
        print(f"SKIPPED sheet: {source_file} (insufficient regression numeric history)")
        return []

    data_start = data_rows[0]
    data_end = data_rows[-1]
    forecast_x_row = data_end + 1
    next_x_value = as_float(sheet.cells(forecast_x_row, x_col).value)
    if next_x_value is None:
        current_x = as_float(sheet.cells(data_end, x_col).value)
        if current_x is None:
            print(f"SKIPPED sheet: {source_file} (unable to infer next X value)")
            return []
        sheet.cells(forecast_x_row, x_col).value = current_x + 1.0

    max_factor, min_factor = nearby_max_min_factors(scan, anchor_row, anchor_col)

    temp_start_row = max(scan["last_row"] + 5, anchor_row + 5)
    temp_col = max(scan["last_col"] + 2, anchor_col + 6)

    calc_specs: List[Tuple[int, int, int]] = []
    for n in range(1, N_QUARTERS + 1):
        start_row = data_end - n + 1
        if start_row < data_start:
            continue
        calc_specs.append((n, start_row, data_end))

    if not calc_specs:
        return []

    for idx, (_, start_row, end_row) in enumerate(calc_specs):
        row = temp_start_row + idx
        intercept_cell = sheet.cells(row, temp_col)
        slope_cell = sheet.cells(row, temp_col + 1)
        forecast_cell = sheet.cells(row, temp_col + 2)
        max_cell = sheet.cells(row, temp_col + 3)
        min_cell = sheet.cells(row, temp_col + 4)

        set_r1c1_formula2(
            intercept_cell,
            f"=INTERCEPT(R{start_row}C{y_col}:R{end_row}C{y_col},R{start_row}C{x_col}:R{end_row}C{x_col})",
        )
        set_r1c1_formula2(
            slope_cell,
            f"=SLOPE(R{start_row}C{y_col}:R{end_row}C{y_col},R{start_row}C{x_col}:R{end_row}C{x_col})",
        )
        set_r1c1_formula2(
            forecast_cell,
            f"=R{row}C{temp_col}+R{row}C{temp_col+1}*R{forecast_x_row}C{x_col}",
        )
        set_r1c1_formula2(max_cell, f"=R{row}C{temp_col+2}*{max_factor}")
        set_r1c1_formula2(min_cell, f"=R{row}C{temp_col+2}*{min_factor}")

    wb.app.calculate()

    last_row = temp_start_row + len(calc_specs) - 1
    calc_values = rows_2d(
        sheet.range((temp_start_row, temp_col), (last_row, temp_col + 4)).value
    )

    output: List[Dict[str, Any]] = []
    prev_calc: Optional[Tuple[Any, ...]] = None
    for idx, (n, _, _) in enumerate(calc_specs):
        intercept, slope, forecast_value, forecast_max, forecast_min = [
            clean_excel_value(v) for v in calc_values[idx]
        ]
        current_calc = (intercept, slope, forecast_value, forecast_max, forecast_min)

        # Prevent duplicate terminal rows when formulas converge to same result.
        if prev_calc is not None and all(values_equal(a, b) for a, b in zip(prev_calc, current_calc)):
            continue

        range_width = None
        if as_float(forecast_max) is not None and as_float(forecast_min) is not None:
            range_width = as_float(forecast_max) - as_float(forecast_min)

        output.append(
            {
                "model": meta.model,
                "ticker": meta.ticker,
                "model_period": meta.model_period,
                "model_date": meta.model_date,
                "method": "regression",
                "parameter_name": "num_quarters_used",
                "parameter_value": n,
                "num_quarters_used": n,
                "forecast_value": forecast_value,
                "actual_value": None,
                "forecast_max": forecast_max,
                "forecast_min": forecast_min,
                "range_width": range_width,
                "intercept": intercept,
                "slope": slope,
                "source_file": source_file,
            }
        )
        prev_calc = current_calc

    return output


def write_sheet(ws: Any, columns: Sequence[str], rows: Iterable[Dict[str, Any]]) -> None:
    ws.append(list(columns))
    for row in rows:
        ws.append([row.get(col) for col in columns])

    for cell in ws[1]:
        cell.font = Font(bold=True)

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions

    for col_idx, _ in enumerate(columns, start=1):
        letter = get_column_letter(col_idx)
        max_len = len(str(ws.cell(row=1, column=col_idx).value or ""))
        for row_idx in range(2, ws.max_row + 1):
            value = ws.cell(row=row_idx, column=col_idx).value
            text_len = len(str(value)) if value is not None else 0
            if text_len > max_len:
                max_len = text_len
        ws.column_dimensions[letter].width = min(max(12, max_len + 2), 40)


def write_output_workbook(
    output_path: Path, empirical_rows: List[Dict[str, Any]], regression_rows: List[Dict[str, Any]]
) -> None:
    wb = Workbook()
    ws_emp = wb.active
    ws_emp.title = "empirical_candidates"
    ws_reg = wb.create_sheet("regression_candidates")

    write_sheet(ws_emp, EMPIRICAL_COLUMNS, empirical_rows)
    write_sheet(ws_reg, REGRESSION_COLUMNS, regression_rows)

    wb.save(output_path)


def iter_source_files(in_dir: Path) -> Iterable[Path]:
    for file_path in sorted(in_dir.iterdir()):
        if not file_path.is_file():
            continue
        if file_path.name.startswith("~"):
            print(f"SKIPPED file: {file_path.name} (temporary file)")
            continue
        if file_path.suffix.lower() != ".xlsx":
            print(f"SKIPPED file: {file_path.name} (not .xlsx)")
            continue
        yield file_path


def main() -> None:
    in_dir = Path(input_dir).expanduser().resolve()
    out_dir = Path(output_dir).expanduser().resolve()

    if not in_dir.exists() or not in_dir.is_dir():
        raise FileNotFoundError(f"Input directory does not exist: {in_dir}")
    out_dir.mkdir(parents=True, exist_ok=True)

    output_path = make_unique_output_path(in_dir, out_dir)
    empirical_rows: List[Dict[str, Any]] = []
    regression_rows: List[Dict[str, Any]] = []
    processed_files = 0

    app = xw.App(visible=False, add_book=False)
    app.display_alerts = False
    app.screen_updating = False
    try:
        try:
            app.api.EnableEvents = False
        except Exception:
            pass
        try:
            app.calculation = "manual"
        except Exception:
            pass

        for file_path in iter_source_files(in_dir):
            print(f"PROCESSING file: {file_path.name}")
            try:
                meta = parse_file_meta(file_path)
            except Exception as exc:
                print(f"SKIPPED file: {file_path.name} ({exc})")
                continue

            wb = None
            try:
                wb = app.books.open(str(file_path), update_links=False)
                empirical_rows.extend(extract_empirical_rows(wb, meta, file_path.name))
                regression_rows.extend(extract_regression_rows(wb, meta, file_path.name))
                processed_files += 1
            except Exception as exc:
                print(f"SKIPPED file: {file_path.name} (processing error: {exc})")
            finally:
                if wb is not None:
                    safe_close_workbook(wb)
    finally:
        app.quit()

    write_output_workbook(output_path, empirical_rows, regression_rows)

    print(f"OUTPUT path: {output_path}")
    print(f"FILES processed: {processed_files}")
    print(f"EMPIRICAL rows: {len(empirical_rows)}")
    print(f"REGRESSION rows: {len(regression_rows)}")


if __name__ == "__main__":
    main()
