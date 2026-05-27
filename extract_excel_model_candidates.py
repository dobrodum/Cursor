#!/usr/bin/env python3
"""
Extract empirical and regression model candidates from .xlsx workbooks.

The script opens each source workbook only once, processes both model sheets
while the workbook is open, and writes one combined output workbook with:
  - empirical_candidates
  - regression_candidates
"""

from __future__ import annotations

import calendar
import re
import sys
from datetime import date
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Sequence, Tuple

import xlwings as xw
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter


# -----------------------------------------------------------------------------
# User-configurable paths
# -----------------------------------------------------------------------------
input_dir = Path("./input").resolve()
output_dir = Path("./output").resolve()


# -----------------------------------------------------------------------------
# Constants
# -----------------------------------------------------------------------------
EMPIRICAL_SHEET_NAME = "Empirical Model"
REGRESSION_SHEET_NAME = "Regression Model"
EMPIRICAL_N_QUARTERS = 10
REGRESSION_N_QUARTERS = 10

FILENAME_PATTERN = re.compile(
    r"^(?P<prefix>.+?)\s*-\s*(?P<ticker>[A-Za-z0-9]+)\s*-\s*(?P<period>(Early|Mid|Late)[A-Za-z]{3}\d{4})",
    re.IGNORECASE,
)
PERIOD_PATTERN = re.compile(r"^(Early|Mid|Late)([A-Za-z]{3})(\d{4})$", re.IGNORECASE)
DAY_MAP = {"early": 5, "mid": 15, "late": 25}

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

EMPIRICAL_LABEL_ALIASES = {
    "avg_penetration_pct": ["avg penetration", "average penetration", "avg pen"],
    "forecast_value": ["estimated total sold", "est total sold", "estimated sold", "total sold"],
    "reported_sales": ["reported sales", "actual sales"],
    "quarterly_sales": ["quarterly sales", "qtr sales"],
    "growth_rate_pct": ["growth rate", "growth %"],
    "sales_captured_in_db_pct": ["sales captured in db", "captured in db", "sales captured"],
    "last_quarter_used": ["last quarter used", "latest quarter"],
    "num_quarters_used_input": ["num quarters used", "number of quarters", "num quarters"],
}

REGRESSION_LABEL_ALIASES = {
    "forecast_value": [
        "tot fcst w/o sa",
        "tot fcst wo sa",
        "total fcst w/o sa",
        "tot fcst without sa",
    ],
    "actual_value": ["actual value", "reported sales", "actual sales"],
    "num_quarters_used_input": ["num quarters used", "number of quarters", "num quarters"],
}


def log(message: str) -> None:
    print(message)


def normalize_text(value: Any) -> str:
    if value is None:
        return ""
    return str(value).strip().lower()


def is_number(value: Any) -> bool:
    return isinstance(value, (int, float)) and not isinstance(value, bool)


def to_float(value: Any) -> Optional[float]:
    if is_number(value):
        return float(value)
    return None


def normalize_matrix(raw_value: Any) -> List[List[Any]]:
    if raw_value is None:
        return []
    if isinstance(raw_value, tuple):
        raw_value = list(raw_value)
    if not isinstance(raw_value, list):
        return [[raw_value]]
    if not raw_value:
        return []
    if isinstance(raw_value[0], tuple):
        return [list(row) for row in raw_value]
    if not isinstance(raw_value[0], list):
        return [list(raw_value)]
    return raw_value


def parse_filename_metadata(file_path: Path) -> Dict[str, Any]:
    stem = file_path.stem
    match = FILENAME_PATTERN.search(stem)

    ticker = "UNKNOWN"
    model_period = "Unknown_0000"
    model_date = ""

    if match:
        ticker = match.group("ticker").upper()
        period_token = match.group("period")
        period_match = PERIOD_PATTERN.match(period_token)
        if period_match:
            period_part = period_match.group(1).title()
            month_part = period_match.group(2).title()
            year_part = period_match.group(3)
            month_token = month_part[:3]
            try:
                month_num = list(calendar.month_abbr).index(month_token)
                day_num = DAY_MAP[period_part.lower()]
                parsed_date = date(int(year_part), month_num, day_num)
                model_period = f"{period_part}{month_part}_{year_part}"
                model_date = parsed_date.isoformat()
            except (ValueError, KeyError):
                model_period = f"{period_part}{month_part}_{year_part}"

    model = f"{ticker}_{model_period}"
    return {
        "model": model,
        "ticker": ticker,
        "model_period": model_period,
        "model_date": model_date,
        "source_file": file_path.name,
    }


def next_output_path(in_dir: Path, out_dir: Path) -> Path:
    base_name = f"{in_dir.name}_PARAM.xlsx"
    candidate = out_dir / base_name
    if not candidate.exists():
        return candidate

    idx = 1
    while True:
        candidate = out_dir / f"{in_dir.name}_PARAM.{idx}.xlsx"
        if not candidate.exists():
            return candidate
        idx += 1


def output_artifact_pattern(in_dir: Path) -> re.Pattern[str]:
    return re.compile(
        rf"^{re.escape(in_dir.name)}_PARAM(?:\.\d+)?\.xlsx$",
        re.IGNORECASE,
    )


def close_source_workbook(wb: xw.Book) -> None:
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
        wb.close()
    except Exception:
        pass


def set_formula2(cell: xw.Range, formula_r1c1: str) -> None:
    try:
        cell.formula2 = formula_r1c1
    except Exception:
        cell.formula = formula_r1c1


def snapshot_sheet(sheet: xw.Sheet) -> Tuple[int, int, List[List[Any]], Dict[Tuple[int, int], Any]]:
    used = sheet.used_range
    start_row = used.row
    start_col = used.column
    matrix = normalize_matrix(used.value)
    cell_map: Dict[Tuple[int, int], Any] = {}
    for r_idx, row_values in enumerate(matrix):
        for c_idx, value in enumerate(row_values):
            cell_map[(start_row + r_idx, start_col + c_idx)] = value
    return start_row, start_col, matrix, cell_map


def find_anchor_max(start_row: int, start_col: int, matrix: Sequence[Sequence[Any]]) -> Optional[Tuple[int, int]]:
    max_candidates: List[Tuple[int, int]] = []
    min_candidates: List[Tuple[int, int]] = []

    for r_idx, row_values in enumerate(matrix):
        for c_idx, value in enumerate(row_values):
            text = normalize_text(value)
            abs_row = start_row + r_idx
            abs_col = start_col + c_idx
            if text == "max":
                max_candidates.append((abs_row, abs_col))
            elif text == "min":
                min_candidates.append((abs_row, abs_col))

    if not max_candidates:
        return None
    if not min_candidates:
        return max_candidates[0]

    best = None
    best_dist = float("inf")
    for candidate in max_candidates:
        cand_row, cand_col = candidate
        dist = min(abs(cand_row - min_row) + abs(cand_col - min_col) for min_row, min_col in min_candidates)
        if dist < best_dist:
            best_dist = dist
            best = candidate
    return best


def find_label_cell(
    start_row: int,
    start_col: int,
    matrix: Sequence[Sequence[Any]],
    anchor_row: int,
    anchor_col: int,
    aliases: Iterable[str],
    row_window: int = 120,
    col_window: int = 40,
) -> Optional[Tuple[int, int]]:
    alias_list = [a.lower() for a in aliases]
    best_cell: Optional[Tuple[int, int]] = None
    best_dist = float("inf")

    for r_idx, row_values in enumerate(matrix):
        abs_row = start_row + r_idx
        if abs(abs_row - anchor_row) > row_window:
            continue
        for c_idx, value in enumerate(row_values):
            abs_col = start_col + c_idx
            if abs(abs_col - anchor_col) > col_window:
                continue
            text = normalize_text(value)
            if not text:
                continue
            if any(alias in text for alias in alias_list):
                dist = abs(abs_row - anchor_row) + abs(abs_col - anchor_col)
                if dist < best_dist:
                    best_dist = dist
                    best_cell = (abs_row, abs_col)
    return best_cell


def resolve_value_cell(label_row: int, label_col: int, cell_map: Dict[Tuple[int, int], Any]) -> Tuple[int, int]:
    for dc in (1, 2, -1):
        candidate = (label_row, label_col + dc)
        value = cell_map.get(candidate)
        if value is None:
            continue
        if isinstance(value, str) and normalize_text(value) in {"max", "min"}:
            continue
        if value != "":
            return candidate

    for dr in (1, -1, 2):
        candidate = (label_row + dr, label_col)
        value = cell_map.get(candidate)
        if value is None:
            continue
        if isinstance(value, str) and normalize_text(value) in {"max", "min"}:
            continue
        if value != "":
            return candidate

    return (label_row, label_col + 1)


def resolve_offset(
    start_row: int,
    start_col: int,
    matrix: Sequence[Sequence[Any]],
    cell_map: Dict[Tuple[int, int], Any],
    anchor_row: int,
    anchor_col: int,
    aliases: Iterable[str],
    default: Optional[Tuple[int, int]] = None,
) -> Optional[Tuple[int, int]]:
    label_cell = find_label_cell(start_row, start_col, matrix, anchor_row, anchor_col, aliases)
    if not label_cell:
        return default
    value_row, value_col = resolve_value_cell(label_cell[0], label_cell[1], cell_map)
    return value_row - anchor_row, value_col - anchor_col


def read_offset(sheet: xw.Sheet, anchor_row: int, anchor_col: int, offset: Optional[Tuple[int, int]]) -> Any:
    if offset is None:
        return None
    row = anchor_row + offset[0]
    col = anchor_col + offset[1]
    if row < 1 or col < 1:
        return None
    return sheet.range((row, col)).value


def find_penetration_series(
    start_row: int,
    start_col: int,
    matrix: Sequence[Sequence[Any]],
    anchor_row: int,
    anchor_col: int,
) -> Optional[Tuple[int, List[int]]]:
    best: Optional[Tuple[int, List[int], int]] = None

    for r_idx, row_values in enumerate(matrix):
        abs_row = start_row + r_idx
        if abs(abs_row - anchor_row) > 120:
            continue
        for c_idx, value in enumerate(row_values):
            abs_col = start_col + c_idx
            if abs(abs_col - anchor_col) > 40:
                continue
            text = normalize_text(value)
            if "penetration" not in text:
                continue
            numeric_cols: List[int] = []
            for n_idx in range(c_idx + 1, len(row_values)):
                maybe_number = row_values[n_idx]
                if is_number(maybe_number):
                    numeric_cols.append(start_col + n_idx)
            if len(numeric_cols) >= EMPIRICAL_N_QUARTERS:
                dist = abs(abs_row - anchor_row) + abs(abs_col - anchor_col)
                candidate = (abs_row, numeric_cols, dist)
                if best is None or candidate[2] < best[2]:
                    best = candidate

    if best is None:
        return None
    return best[0], best[1]


def find_regression_data_rows(sheet: xw.Sheet, anchor_row: int, x_col: int, y_col: int) -> List[int]:
    start_scan = max(1, anchor_row - 120)
    end_scan = max(1, anchor_row - 1)

    rows: List[int] = []
    for row in range(start_scan, end_scan + 1):
        x_value = sheet.range((row, x_col)).value
        y_value = sheet.range((row, y_col)).value
        if is_number(x_value) and is_number(y_value):
            rows.append(row)
    return rows


def extract_empirical_rows(wb: xw.Book, meta: Dict[str, Any]) -> List[Dict[str, Any]]:
    try:
        sheet = wb.sheets[EMPIRICAL_SHEET_NAME]
    except Exception:
        return []

    start_row, start_col, matrix, cell_map = snapshot_sheet(sheet)
    anchor = find_anchor_max(start_row, start_col, matrix)
    if not anchor:
        return []
    anchor_row, anchor_col = anchor

    offsets: Dict[str, Optional[Tuple[int, int]]] = {}
    max_value_row, max_value_col = resolve_value_cell(anchor_row, anchor_col, cell_map)
    offsets["forecast_max"] = (max_value_row - anchor_row, max_value_col - anchor_col)
    offsets["forecast_min"] = resolve_offset(
        start_row,
        start_col,
        matrix,
        cell_map,
        anchor_row,
        anchor_col,
        aliases=["min"],
        default=(1, 1),
    )
    for key, aliases in EMPIRICAL_LABEL_ALIASES.items():
        offsets[key] = resolve_offset(
            start_row,
            start_col,
            matrix,
            cell_map,
            anchor_row,
            anchor_col,
            aliases=aliases,
        )

    penetration_series = find_penetration_series(start_row, start_col, matrix, anchor_row, anchor_col)

    rows: List[Dict[str, Any]] = []
    for n_quarters in range(1, EMPIRICAL_N_QUARTERS + 1):
        formula_updated = False

        num_quarters_input_offset = offsets.get("num_quarters_used_input")
        if num_quarters_input_offset is not None:
            input_row = anchor_row + num_quarters_input_offset[0]
            input_col = anchor_col + num_quarters_input_offset[1]
            sheet.range((input_row, input_col)).value = n_quarters
            formula_updated = True

        avg_pen_offset = offsets.get("avg_penetration_pct")
        if avg_pen_offset is not None and penetration_series is not None:
            series_row, series_cols = penetration_series
            if len(series_cols) >= n_quarters:
                start_col_idx = series_cols[-n_quarters]
                end_col_idx = series_cols[-1]
                avg_cell = sheet.range((anchor_row + avg_pen_offset[0], anchor_col + avg_pen_offset[1]))
                set_formula2(
                    avg_cell,
                    f"=AVERAGE(R{series_row}C{start_col_idx}:R{series_row}C{end_col_idx})",
                )
                formula_updated = True

        if formula_updated:
            wb.app.calculate()

        num_quarters_used = read_offset(sheet, anchor_row, anchor_col, offsets.get("num_quarters_used_input"))
        if num_quarters_used is None:
            num_quarters_used = n_quarters

        avg_penetration_pct = read_offset(sheet, anchor_row, anchor_col, offsets.get("avg_penetration_pct"))
        forecast_value = read_offset(sheet, anchor_row, anchor_col, offsets.get("forecast_value"))
        reported_sales = read_offset(sheet, anchor_row, anchor_col, offsets.get("reported_sales"))
        forecast_max = read_offset(sheet, anchor_row, anchor_col, offsets.get("forecast_max"))
        forecast_min = read_offset(sheet, anchor_row, anchor_col, offsets.get("forecast_min"))
        quarterly_sales = read_offset(sheet, anchor_row, anchor_col, offsets.get("quarterly_sales"))
        growth_rate_pct = read_offset(sheet, anchor_row, anchor_col, offsets.get("growth_rate_pct"))
        sales_captured_in_db_pct = read_offset(
            sheet, anchor_row, anchor_col, offsets.get("sales_captured_in_db_pct")
        )
        last_quarter_used = read_offset(sheet, anchor_row, anchor_col, offsets.get("last_quarter_used"))

        if (
            avg_penetration_pct is None
            and forecast_value is None
            and reported_sales is None
            and forecast_max is None
            and forecast_min is None
        ):
            continue

        max_num = to_float(forecast_max)
        min_num = to_float(forecast_min)
        range_width = (max_num - min_num) if max_num is not None and min_num is not None else None

        rows.append(
            {
                "model": meta["model"],
                "ticker": meta["ticker"],
                "model_period": meta["model_period"],
                "model_date": meta["model_date"],
                "method": "empirical",
                "parameter_name": "avg_penetration_pct",
                "parameter_value": avg_penetration_pct,
                "num_quarters_used": num_quarters_used,
                "last_quarter_used": last_quarter_used,
                "forecast_value": forecast_value,
                "actual_value": reported_sales,
                "forecast_max": forecast_max,
                "forecast_min": forecast_min,
                "range_width": range_width,
                "avg_penetration_pct": avg_penetration_pct,
                "quarterly_sales": quarterly_sales,
                "reported_sales": reported_sales,
                "growth_rate_pct": growth_rate_pct,
                "sales_captured_in_db_pct": sales_captured_in_db_pct,
                "source_file": meta["source_file"],
            }
        )

    return rows


def extract_regression_rows(wb: xw.Book, meta: Dict[str, Any]) -> List[Dict[str, Any]]:
    try:
        sheet = wb.sheets[REGRESSION_SHEET_NAME]
    except Exception:
        return []

    start_row, start_col, matrix, cell_map = snapshot_sheet(sheet)
    anchor = find_anchor_max(start_row, start_col, matrix)
    if not anchor:
        return []
    anchor_row, anchor_col = anchor

    y_col = anchor_col - 7
    x_col = anchor_col - 11

    offsets: Dict[str, Optional[Tuple[int, int]]] = {}
    max_value_row, max_value_col = resolve_value_cell(anchor_row, anchor_col, cell_map)
    offsets["forecast_max"] = (max_value_row - anchor_row, max_value_col - anchor_col)
    offsets["forecast_min"] = resolve_offset(
        start_row,
        start_col,
        matrix,
        cell_map,
        anchor_row,
        anchor_col,
        aliases=["min"],
        default=(1, 1),
    )
    for key, aliases in REGRESSION_LABEL_ALIASES.items():
        offsets[key] = resolve_offset(
            start_row,
            start_col,
            matrix,
            cell_map,
            anchor_row,
            anchor_col,
            aliases=aliases,
        )

    used_last_col = start_col + (max((len(r) for r in matrix), default=1) - 1)
    scratch_col = max(used_last_col + 2, anchor_col + 2)
    intercept_cell = sheet.range((anchor_row, scratch_col))
    slope_cell = sheet.range((anchor_row + 1, scratch_col))

    data_rows = find_regression_data_rows(sheet, anchor_row=anchor_row, x_col=x_col, y_col=y_col)
    if len(data_rows) < 2:
        return []

    rows: List[Dict[str, Any]] = []
    prev_signature: Optional[Tuple[Any, ...]] = None

    max_window = min(REGRESSION_N_QUARTERS, len(data_rows))
    for n_quarters in range(2, max_window + 1):
        window_rows = data_rows[-n_quarters:]
        start_data_row = window_rows[0]
        end_data_row = window_rows[-1]

        formula_updated = False

        num_quarters_input_offset = offsets.get("num_quarters_used_input")
        if num_quarters_input_offset is not None:
            input_row = anchor_row + num_quarters_input_offset[0]
            input_col = anchor_col + num_quarters_input_offset[1]
            sheet.range((input_row, input_col)).value = n_quarters
            formula_updated = True

        set_formula2(
            intercept_cell,
            (
                f"=INTERCEPT(R{start_data_row}C{y_col}:R{end_data_row}C{y_col},"
                f"R{start_data_row}C{x_col}:R{end_data_row}C{x_col})"
            ),
        )
        set_formula2(
            slope_cell,
            (
                f"=SLOPE(R{start_data_row}C{y_col}:R{end_data_row}C{y_col},"
                f"R{start_data_row}C{x_col}:R{end_data_row}C{x_col})"
            ),
        )
        formula_updated = True

        if formula_updated:
            wb.app.calculate()

        intercept = intercept_cell.value
        slope = slope_cell.value
        forecast_value = read_offset(sheet, anchor_row, anchor_col, offsets.get("forecast_value"))
        actual_value = read_offset(sheet, anchor_row, anchor_col, offsets.get("actual_value"))
        forecast_max = read_offset(sheet, anchor_row, anchor_col, offsets.get("forecast_max"))
        forecast_min = read_offset(sheet, anchor_row, anchor_col, offsets.get("forecast_min"))

        max_num = to_float(forecast_max)
        min_num = to_float(forecast_min)
        range_width = (max_num - min_num) if max_num is not None and min_num is not None else None

        intercept_num = to_float(intercept)
        slope_num = to_float(slope)
        forecast_num = to_float(forecast_value)
        signature = (
            round(intercept_num, 10) if intercept_num is not None else None,
            round(slope_num, 10) if slope_num is not None else None,
            round(forecast_num, 10) if forecast_num is not None else None,
            round(max_num, 10) if max_num is not None else None,
            round(min_num, 10) if min_num is not None else None,
        )
        if prev_signature == signature:
            continue
        prev_signature = signature

        rows.append(
            {
                "model": meta["model"],
                "ticker": meta["ticker"],
                "model_period": meta["model_period"],
                "model_date": meta["model_date"],
                "method": "regression",
                "parameter_name": "num_quarters_used",
                "parameter_value": n_quarters,
                "num_quarters_used": n_quarters,
                "forecast_value": forecast_value,
                "actual_value": actual_value,
                "forecast_max": forecast_max,
                "forecast_min": forecast_min,
                "range_width": range_width,
                "intercept": intercept,
                "slope": slope,
                "source_file": meta["source_file"],
            }
        )

    intercept_cell.value = None
    slope_cell.value = None
    return rows


def write_sheet(ws, headers: Sequence[str], rows: Sequence[Dict[str, Any]]) -> None:
    ws.append(list(headers))
    for cell in ws[1]:
        cell.font = Font(bold=True)

    for row in rows:
        ws.append([row.get(header) for header in headers])

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions

    for col_idx, header in enumerate(headers, start=1):
        max_len = len(header)
        for row_idx in range(2, ws.max_row + 1):
            value = ws.cell(row=row_idx, column=col_idx).value
            if value is None:
                continue
            max_len = max(max_len, len(str(value)))
        ws.column_dimensions[get_column_letter(col_idx)].width = min(max(12, max_len + 2), 60)


def write_output_workbook(
    output_path: Path,
    empirical_rows: Sequence[Dict[str, Any]],
    regression_rows: Sequence[Dict[str, Any]],
) -> None:
    workbook = Workbook()
    default_sheet = workbook.active
    workbook.remove(default_sheet)

    empirical_ws = workbook.create_sheet("empirical_candidates")
    write_sheet(empirical_ws, EMPIRICAL_HEADERS, empirical_rows)

    regression_ws = workbook.create_sheet("regression_candidates")
    write_sheet(regression_ws, REGRESSION_HEADERS, regression_rows)

    workbook.save(output_path)


def should_process_file(file_path: Path, out_artifact_re: re.Pattern[str]) -> Tuple[bool, str]:
    if not file_path.is_file():
        return False, "not a regular file"
    if file_path.suffix.lower() != ".xlsx":
        return False, "not an .xlsx file"
    if file_path.name.startswith("~"):
        return False, "temporary file"
    if out_artifact_re.match(file_path.name):
        return False, "output artifact"
    return True, ""


def configure_app(app: xw.App) -> None:
    app.visible = False
    app.display_alerts = False
    app.screen_updating = False
    try:
        app.calculation = "manual"
    except Exception:
        pass
    try:
        app.api.EnableEvents = False
    except Exception:
        pass


def main() -> int:
    if not input_dir.exists() or not input_dir.is_dir():
        log(f"Input directory does not exist or is not a directory: {input_dir}")
        return 1

    output_dir.mkdir(parents=True, exist_ok=True)
    output_path = next_output_path(input_dir, output_dir)
    artifact_re = output_artifact_pattern(input_dir)

    empirical_rows: List[Dict[str, Any]] = []
    regression_rows: List[Dict[str, Any]] = []
    files_processed = 0

    app = xw.App(visible=False, add_book=False)
    configure_app(app)

    try:
        for file_path in sorted(input_dir.iterdir()):
            should_process, reason = should_process_file(file_path, artifact_re)
            if not should_process:
                log(f"Skipped: {file_path.name} ({reason})")
                continue

            log(f"Processing: {file_path.name}")
            wb = None
            try:
                wb = app.books.open(str(file_path), update_links=False)
                meta = parse_filename_metadata(file_path)

                file_empirical_rows = extract_empirical_rows(wb, meta)
                file_regression_rows = extract_regression_rows(wb, meta)

                empirical_rows.extend(file_empirical_rows)
                regression_rows.extend(file_regression_rows)
                files_processed += 1
            except Exception as exc:
                log(f"Skipped: {file_path.name} (processing error: {exc})")
            finally:
                if wb is not None:
                    close_source_workbook(wb)
    finally:
        try:
            app.quit()
        except Exception:
            pass

    write_output_workbook(output_path, empirical_rows, regression_rows)

    log(f"Output path: {output_path}")
    log(f"Number of files processed: {files_processed}")
    log(f"Number of empirical rows: {len(empirical_rows)}")
    log(f"Number of regression rows: {len(regression_rows)}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
