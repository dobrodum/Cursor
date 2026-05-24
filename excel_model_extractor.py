#!/usr/bin/env python3
from __future__ import annotations

import re
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Sequence, Tuple

import openpyxl
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter
import xwings as xw


# ---------------------------------------------------------------------------
# User-configurable paths
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

EMPIRICAL_OFFSETS = {
    "num_quarters_used": -10,
    "last_quarter_used": -9,
    "quarterly_sales": -8,
    "reported_sales": -7,
    "growth_rate_pct": -6,
    "sales_captured_in_db_pct": -5,
    "forecast_value": -2,
    "forecast_max": 0,
    "forecast_min": 1,
    "avg_penetration_pct": -4,
}

REGRESSION_OFFSETS = {
    "num_quarters_used": -10,
    "tot_fcst_wo_sa": -2,
    "forecast_max": 0,
    "forecast_min": 1,
    "actual_value": -1,
}

PERIOD_DAY_MAP = {"Early": 5, "Mid": 15, "Late": 25}
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
    return re.sub(r"\s+", " ", str(value).strip().lower())


def to_matrix(values: Any) -> List[List[Any]]:
    if values is None:
        return []
    if not isinstance(values, (list, tuple)):
        return [[values]]

    rows: List[List[Any]] = []
    if values and not isinstance(values[0], (list, tuple)):
        rows.append(list(values))
        return rows

    for row in values:
        if isinstance(row, (list, tuple)):
            rows.append(list(row))
        else:
            rows.append([row])
    return rows


def parse_file_label(file_name: str) -> Dict[str, str]:
    stem = Path(file_name).stem
    parts = [part.strip() for part in stem.split(" - ")]

    ticker = ""
    if len(parts) >= 2:
        ticker = parts[1].strip().split()[0]

    period_source = parts[2].strip() if len(parts) >= 3 else stem
    period_source = period_source.split("_")[0]

    match = re.search(
        r"(Early|Mid|Late)\s*([A-Za-z]{3,9})\s*(\d{4})",
        period_source,
        flags=re.IGNORECASE,
    )
    if not match:
        match = re.search(
            r"(Early|Mid|Late)\s*([A-Za-z]{3,9})\s*(\d{4})",
            stem,
            flags=re.IGNORECASE,
        )

    model_period = period_source
    model_date = ""
    if match:
        timing = match.group(1).title()
        month_token = match.group(2)[:3].lower()
        year = int(match.group(3))
        month_num = MONTH_MAP.get(month_token)
        if month_num:
            model_period = f"{timing}{month_token.title()}_{year}"
            day = PERIOD_DAY_MAP[timing]
            model_date = f"{year:04d}-{month_num:02d}-{day:02d}"

    model = f"{ticker}_{model_period}" if ticker else model_period

    return {
        "ticker": ticker,
        "model_period": model_period,
        "model_date": model_date,
        "model": model,
    }


def make_output_path(input_path: Path, output_path: Path) -> Path:
    base_name = f"{input_path.name}_PARAM"
    candidate = output_path / f"{base_name}.xlsx"
    if not candidate.exists():
        return candidate

    idx = 1
    while True:
        candidate = output_path / f"{base_name}.{idx}.xlsx"
        if not candidate.exists():
            return candidate
        idx += 1


def is_blank(value: Any) -> bool:
    return value is None or (isinstance(value, str) and value.strip() == "")


def numeric_or_none(value: Any) -> Optional[float]:
    if is_blank(value):
        return None
    try:
        return float(value)
    except (TypeError, ValueError):
        return None


def diff_or_blank(left: Any, right: Any) -> Any:
    left_num = numeric_or_none(left)
    right_num = numeric_or_none(right)
    if left_num is None or right_num is None:
        return ""
    return left_num - right_num


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
        return
    except Exception:
        pass

    try:
        wb.api.Close(False)
    except Exception:
        pass


def get_cell_value(sheet: xw.Sheet, row: int, col: int) -> Any:
    if row < 1 or col < 1:
        return None
    try:
        return sheet.cells(row, col).value
    except Exception:
        return None


def set_formula2_r1c1(cell: xw.Range, formula_r1c1: str) -> None:
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


def find_anchor_cell(
    sheet: xw.Sheet,
    anchor_token: str = "max",
) -> Optional[Tuple[int, int, List[List[Any]], int, int]]:
    used = sheet.used_range
    values = to_matrix(used.value)
    if not values:
        return None

    start_row = used.row
    start_col = used.column
    expected = normalize_text(anchor_token)

    for row_offset, row in enumerate(values):
        for col_offset, value in enumerate(row):
            text = normalize_text(value)
            if text == expected or text.startswith(f"{expected} "):
                return (
                    start_row + row_offset,
                    start_col + col_offset,
                    values,
                    start_row,
                    start_col,
                )
    return None


def locate_header_columns(
    matrix: Sequence[Sequence[Any]],
    start_row: int,
    start_col: int,
    header_row: int,
    anchor_col: int,
) -> Dict[str, int]:
    header_idx = header_row - start_row
    if header_idx < 0 or header_idx >= len(matrix):
        return {}

    header_values = matrix[header_idx]
    located: Dict[str, int] = {}

    def pick(key: str, col_idx: int) -> None:
        prev = located.get(key)
        if prev is None or abs(col_idx - anchor_col) < abs(prev - anchor_col):
            located[key] = col_idx

    for rel_col_idx, raw_value in enumerate(header_values):
        col_idx = start_col + rel_col_idx
        text = normalize_text(raw_value)
        if not text:
            continue

        if text == "max" or text.startswith("max "):
            pick("forecast_max", col_idx)
        if text == "min" or text.startswith("min "):
            pick("forecast_min", col_idx)
        if "estimated total sold" in text or "est total sold" in text:
            pick("forecast_value", col_idx)
        if "tot fcst w/o sa" in text or "tot fcst wo sa" in text:
            pick("tot_fcst_wo_sa", col_idx)
        if "forecast without sa" in text or "forecast w/o sa" in text:
            pick("tot_fcst_wo_sa", col_idx)
        if "reported sales" in text:
            pick("reported_sales", col_idx)
            pick("actual_value", col_idx)
        if text in {"actual", "actual value"}:
            pick("actual_value", col_idx)
        if "quarterly sales" in text or "quarter sales" in text:
            pick("quarterly_sales", col_idx)
        if "growth rate" in text:
            pick("growth_rate_pct", col_idx)
        if "sales captured in db" in text or "captured in db" in text:
            pick("sales_captured_in_db_pct", col_idx)
        if "avg penetration" in text:
            pick("avg_penetration_pct", col_idx)
        if "last quarter" in text:
            pick("last_quarter_used", col_idx)
        if "num quarters" in text or "quarters used" in text:
            pick("num_quarters_used", col_idx)

    return located


def resolve_column(
    detected_columns: Dict[str, int],
    field: str,
    anchor_col: int,
    offset_map: Dict[str, int],
) -> int:
    if field in detected_columns:
        return detected_columns[field]
    return anchor_col + offset_map[field]


def build_empirical_rows(
    wb: xw.Book,
    metadata: Dict[str, str],
    source_file: str,
) -> List[Dict[str, Any]]:
    try:
        sheet = wb.sheets["Empirical Model"]
    except Exception:
        print(f"Skipped empirical extraction for {source_file}: sheet 'Empirical Model' missing")
        return []

    anchor_info = find_anchor_cell(sheet, "max")
    if anchor_info is None:
        print(f"Skipped empirical extraction for {source_file}: anchor 'max' not found")
        return []

    anchor_row, anchor_col, matrix, start_row, start_col = anchor_info
    detected = locate_header_columns(matrix, start_row, start_col, anchor_row, anchor_col)

    n_quarters = 10
    data_rows = [anchor_row + i for i in range(1, n_quarters + 1)]

    num_quarters_col = resolve_column(detected, "num_quarters_used", anchor_col, EMPIRICAL_OFFSETS)
    last_quarter_col = resolve_column(detected, "last_quarter_used", anchor_col, EMPIRICAL_OFFSETS)
    quarterly_sales_col = resolve_column(detected, "quarterly_sales", anchor_col, EMPIRICAL_OFFSETS)
    reported_sales_col = resolve_column(detected, "reported_sales", anchor_col, EMPIRICAL_OFFSETS)
    forecast_col = resolve_column(detected, "forecast_value", anchor_col, EMPIRICAL_OFFSETS)
    forecast_max_col = resolve_column(detected, "forecast_max", anchor_col, EMPIRICAL_OFFSETS)
    forecast_min_col = resolve_column(detected, "forecast_min", anchor_col, EMPIRICAL_OFFSETS)
    growth_rate_col = resolve_column(detected, "growth_rate_pct", anchor_col, EMPIRICAL_OFFSETS)
    sales_captured_col = resolve_column(
        detected,
        "sales_captured_in_db_pct",
        anchor_col,
        EMPIRICAL_OFFSETS,
    )
    avg_penetration_raw_col = resolve_column(
        detected,
        "avg_penetration_pct",
        anchor_col,
        EMPIRICAL_OFFSETS,
    )

    helper_col = sheet.used_range.last_cell.column + 2
    helper_values: List[Any] = [None] * n_quarters
    wrote_formulas = False

    for idx, row in enumerate(data_rows, start=1):
        formula = (
            f'=IFERROR(AVERAGE(R[-{idx - 1}]C{quarterly_sales_col}:RC{quarterly_sales_col})/'
            f'AVERAGE(R[-{idx - 1}]C{reported_sales_col}:RC{reported_sales_col}),"")'
        )
        try:
            set_formula2_r1c1(sheet.cells(row, helper_col), formula)
            wrote_formulas = True
        except Exception:
            continue

    if wrote_formulas:
        wb.app.calculate()
        for idx, row in enumerate(data_rows):
            helper_values[idx] = get_cell_value(sheet, row, helper_col)
        for row in data_rows:
            try:
                sheet.cells(row, helper_col).value = None
            except Exception:
                pass

    rows: List[Dict[str, Any]] = []
    for idx, row in enumerate(data_rows, start=1):
        num_quarters = get_cell_value(sheet, row, num_quarters_col)
        if is_blank(num_quarters):
            num_quarters = idx

        forecast_value = get_cell_value(sheet, row, forecast_col)
        actual_value = get_cell_value(sheet, row, reported_sales_col)
        forecast_max = get_cell_value(sheet, row, forecast_max_col)
        forecast_min = get_cell_value(sheet, row, forecast_min_col)
        last_quarter_used = get_cell_value(sheet, row, last_quarter_col)
        quarterly_sales = get_cell_value(sheet, row, quarterly_sales_col)
        reported_sales = get_cell_value(sheet, row, reported_sales_col)
        growth_rate_pct = get_cell_value(sheet, row, growth_rate_col)
        sales_captured_pct = get_cell_value(sheet, row, sales_captured_col)

        avg_penetration_pct = helper_values[idx - 1]
        if is_blank(avg_penetration_pct):
            avg_penetration_pct = get_cell_value(sheet, row, avg_penetration_raw_col)

        if all(
            is_blank(v)
            for v in (
                forecast_value,
                actual_value,
                forecast_max,
                forecast_min,
                quarterly_sales,
                reported_sales,
            )
        ):
            continue

        rows.append(
            {
                "model": metadata["model"],
                "ticker": metadata["ticker"],
                "model_period": metadata["model_period"],
                "model_date": metadata["model_date"],
                "method": "empirical",
                "parameter_name": "avg_penetration_pct",
                "parameter_value": avg_penetration_pct,
                "num_quarters_used": num_quarters,
                "last_quarter_used": last_quarter_used,
                "forecast_value": forecast_value,
                "actual_value": actual_value,
                "forecast_max": forecast_max,
                "forecast_min": forecast_min,
                "range_width": diff_or_blank(forecast_max, forecast_min),
                "avg_penetration_pct": avg_penetration_pct,
                "quarterly_sales": quarterly_sales,
                "reported_sales": reported_sales,
                "growth_rate_pct": growth_rate_pct,
                "sales_captured_in_db_pct": sales_captured_pct,
                "source_file": source_file,
            }
        )

    return rows


def compare_token(value: Any) -> Any:
    num = numeric_or_none(value)
    if num is not None:
        return round(num, 10)
    return normalize_text(value)


def build_regression_rows(
    wb: xw.Book,
    metadata: Dict[str, str],
    source_file: str,
) -> List[Dict[str, Any]]:
    try:
        sheet = wb.sheets["Regression Model"]
    except Exception:
        print(f"Skipped regression extraction for {source_file}: sheet 'Regression Model' missing")
        return []

    anchor_info = find_anchor_cell(sheet, "max")
    if anchor_info is None:
        print(f"Skipped regression extraction for {source_file}: anchor 'max' not found")
        return []

    anchor_row, anchor_col, matrix, start_row, start_col = anchor_info
    detected = locate_header_columns(matrix, start_row, start_col, anchor_row, anchor_col)

    n_quarters = 10
    data_rows = [anchor_row + i for i in range(1, n_quarters + 1)]

    y_col = anchor_col - 7
    x_col = anchor_col - 11

    num_quarters_col = resolve_column(detected, "num_quarters_used", anchor_col, REGRESSION_OFFSETS)
    forecast_col = resolve_column(detected, "tot_fcst_wo_sa", anchor_col, REGRESSION_OFFSETS)
    forecast_max_col = resolve_column(detected, "forecast_max", anchor_col, REGRESSION_OFFSETS)
    forecast_min_col = resolve_column(detected, "forecast_min", anchor_col, REGRESSION_OFFSETS)
    actual_value_col = resolve_column(detected, "actual_value", anchor_col, REGRESSION_OFFSETS)

    helper_start_col = sheet.used_range.last_cell.column + 2
    intercept_col = helper_start_col
    slope_col = helper_start_col + 1

    wrote_formulas = False
    for idx, row in enumerate(data_rows, start=1):
        intercept_formula = (
            f'=IFERROR(INTERCEPT(R[-{idx - 1}]C{y_col}:RC{y_col},'
            f'R[-{idx - 1}]C{x_col}:RC{x_col}),"")'
        )
        slope_formula = (
            f'=IFERROR(SLOPE(R[-{idx - 1}]C{y_col}:RC{y_col},'
            f'R[-{idx - 1}]C{x_col}:RC{x_col}),"")'
        )
        try:
            set_formula2_r1c1(sheet.cells(row, intercept_col), intercept_formula)
            wrote_formulas = True
        except Exception:
            pass
        try:
            set_formula2_r1c1(sheet.cells(row, slope_col), slope_formula)
            wrote_formulas = True
        except Exception:
            pass

    if wrote_formulas:
        wb.app.calculate()

    intercept_values = [get_cell_value(sheet, row, intercept_col) for row in data_rows]
    slope_values = [get_cell_value(sheet, row, slope_col) for row in data_rows]

    for row in data_rows:
        try:
            sheet.cells(row, intercept_col).value = None
            sheet.cells(row, slope_col).value = None
        except Exception:
            pass

    rows: List[Dict[str, Any]] = []
    for idx, row in enumerate(data_rows, start=1):
        num_quarters = get_cell_value(sheet, row, num_quarters_col)
        if is_blank(num_quarters):
            num_quarters = idx

        forecast_value = get_cell_value(sheet, row, forecast_col)
        actual_value = get_cell_value(sheet, row, actual_value_col)
        if is_blank(actual_value):
            actual_value = ""
        forecast_max = get_cell_value(sheet, row, forecast_max_col)
        forecast_min = get_cell_value(sheet, row, forecast_min_col)
        intercept = intercept_values[idx - 1]
        slope = slope_values[idx - 1]

        if all(
            is_blank(v)
            for v in (forecast_value, forecast_max, forecast_min, intercept, slope)
        ):
            continue

        rows.append(
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
                "actual_value": actual_value,
                "forecast_max": forecast_max,
                "forecast_min": forecast_min,
                "range_width": diff_or_blank(forecast_max, forecast_min),
                "intercept": intercept,
                "slope": slope,
                "source_file": source_file,
            }
        )

    if len(rows) >= 2:
        prev = rows[-2]
        curr = rows[-1]
        prev_sig = (
            compare_token(prev["forecast_value"]),
            compare_token(prev["forecast_max"]),
            compare_token(prev["forecast_min"]),
            compare_token(prev["intercept"]),
            compare_token(prev["slope"]),
        )
        curr_sig = (
            compare_token(curr["forecast_value"]),
            compare_token(curr["forecast_max"]),
            compare_token(curr["forecast_min"]),
            compare_token(curr["intercept"]),
            compare_token(curr["slope"]),
        )
        if curr_sig == prev_sig:
            rows.pop()

    return rows


def apply_output_formatting(ws: openpyxl.worksheet.worksheet.Worksheet) -> None:
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions

    for cell in ws[1]:
        cell.font = Font(bold=True)

    for col_idx in range(1, ws.max_column + 1):
        max_len = 0
        for row_idx in range(1, ws.max_row + 1):
            value = ws.cell(row=row_idx, column=col_idx).value
            if value is None:
                continue
            max_len = max(max_len, len(str(value)))
        ws.column_dimensions[get_column_letter(col_idx)].width = min(max(12, max_len + 2), 48)


def write_output_workbook(
    output_path: Path,
    empirical_rows: Sequence[Dict[str, Any]],
    regression_rows: Sequence[Dict[str, Any]],
) -> None:
    out_wb = openpyxl.Workbook()
    default_ws = out_wb.active
    out_wb.remove(default_ws)

    empirical_ws = out_wb.create_sheet("empirical_candidates")
    empirical_ws.append(EMPIRICAL_COLUMNS)
    for row in empirical_rows:
        empirical_ws.append([row.get(col, "") for col in EMPIRICAL_COLUMNS])
    apply_output_formatting(empirical_ws)

    regression_ws = out_wb.create_sheet("regression_candidates")
    regression_ws.append(REGRESSION_COLUMNS)
    for row in regression_rows:
        regression_ws.append([row.get(col, "") for col in REGRESSION_COLUMNS])
    apply_output_formatting(regression_ws)

    out_wb.save(output_path)


def run() -> None:
    source_dir = Path(input_dir).expanduser().resolve()
    destination_dir = Path(output_dir).expanduser().resolve()

    if not source_dir.exists():
        raise SystemExit(f"Input folder does not exist: {source_dir}")

    destination_dir.mkdir(parents=True, exist_ok=True)
    output_path = make_output_path(source_dir, destination_dir)

    empirical_rows: List[Dict[str, Any]] = []
    regression_rows: List[Dict[str, Any]] = []
    files_processed = 0

    app: Optional[xw.App] = None
    try:
        app = xw.App(visible=False, add_book=False)
        app.display_alerts = False
        app.screen_updating = False
        try:
            app.api.EnableEvents = False
        except Exception:
            pass
        try:
            app.calculation = "manual"
        except Exception:
            pass

        for file_path in sorted(source_dir.iterdir(), key=lambda p: p.name.lower()):
            if file_path.is_dir():
                continue

            if file_path.name.startswith("~"):
                print(f"Skipped {file_path.name}: temporary Excel file")
                continue

            if file_path.suffix.lower() != ".xlsx":
                print(f"Skipped {file_path.name}: not an .xlsx file")
                continue

            print(f"Processing {file_path.name}")
            wb: Optional[xw.Book] = None
            try:
                wb = app.books.open(str(file_path), update_links=False)
                metadata = parse_file_label(file_path.name)

                empirical_rows.extend(build_empirical_rows(wb, metadata, file_path.name))
                regression_rows.extend(build_regression_rows(wb, metadata, file_path.name))
                files_processed += 1
            except Exception as exc:
                print(f"Skipped {file_path.name}: {exc}")
            finally:
                if wb is not None:
                    safe_close_workbook(wb)
    finally:
        if app is not None:
            try:
                app.api.EnableEvents = True
            except Exception:
                pass
            try:
                app.calculation = "automatic"
            except Exception:
                pass
            app.quit()

    write_output_workbook(output_path, empirical_rows, regression_rows)

    print(f"Output file: {output_path}")
    print(f"Files processed: {files_processed}")
    print(f"Empirical rows: {len(empirical_rows)}")
    print(f"Regression rows: {len(regression_rows)}")


if __name__ == "__main__":
    run()
