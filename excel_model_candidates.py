#!/usr/bin/env python3
"""Build empirical/regression candidate tables from Excel model workbooks."""

from __future__ import annotations

import math
import re
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Sequence, Tuple

import xlwings as xw
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

# Configure these two folders before running.
input_dir = Path("./input")
output_dir = Path("./output")

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

EMPIRICAL_FALLBACK_OFFSETS = {
    "num_quarters_used": -7,
    "last_quarter_used": -6,
    "forecast_value": -5,
    "actual_value": -4,
    "avg_penetration_pct": -3,
    "quarterly_sales": -2,
    "reported_sales": -1,
    "forecast_max": 0,
    "forecast_min": 1,
    "growth_rate_pct": 2,
    "sales_captured_in_db_pct": 3,
}

REGRESSION_FALLBACK_OFFSETS = {
    "num_quarters_used": -6,
    "forecast_value": -5,
    "actual_value": -4,
    "forecast_max": 0,
    "forecast_min": 1,
}

HEADER_PATTERNS = {
    "num_quarters_used": (
        "num quarters",
        "num_q",
        "n quarters",
        "n qtr",
        "quarters used",
    ),
    "last_quarter_used": ("last quarter", "quarter used", "period"),
    "forecast_value": (
        "estimated total sold",
        "tot fcst w/o sa",
        "tot fcst",
        "forecast total",
        "forecast value",
    ),
    "actual_value": ("actual", "reported sales", "actual sales"),
    "forecast_max": ("max", "forecast max"),
    "forecast_min": ("min", "forecast min"),
    "avg_penetration_pct": ("avg penetration", "average penetration", "avg_penetration"),
    "quarterly_sales": ("quarterly sales", "qtr sales"),
    "reported_sales": ("reported sales",),
    "growth_rate_pct": ("growth rate", "growth %", "growth_rate"),
    "sales_captured_in_db_pct": (
        "sales captured in db",
        "captured in db",
        "penetration",
    ),
    "intercept": ("intercept",),
    "slope": ("slope",),
}

MONTH_LOOKUP = {
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
PERIOD_DAY = {"early": 5, "mid": 15, "late": 25}


@dataclass
class SheetSnapshot:
    start_row: int
    start_col: int
    values: List[List[Any]]

    @property
    def end_row(self) -> int:
        return self.start_row + len(self.values) - 1

    @property
    def end_col(self) -> int:
        width = len(self.values[0]) if self.values else 0
        return self.start_col + width - 1


def to_2d(values: Any) -> List[List[Any]]:
    if isinstance(values, list):
        if not values:
            return [[None]]
        if isinstance(values[0], list):
            return values
        return [values]
    return [[values]]


def to_1d(values: Any, expected_len: int) -> List[Any]:
    if isinstance(values, list):
        if values and isinstance(values[0], list):
            return [row[0] if row else None for row in values]
        return values
    return [values] * expected_len


def normalize_text(value: Any) -> str:
    if value is None:
        return ""
    return str(value).strip().lower()


def is_number(value: Any) -> bool:
    number = to_float(value)
    return number is not None


def to_float(value: Any) -> Optional[float]:
    if value is None:
        return None
    if isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        if isinstance(value, float) and math.isnan(value):
            return None
        return float(value)
    text = str(value).strip().replace(",", "")
    if not text:
        return None
    is_percent = text.endswith("%")
    if is_percent:
        text = text[:-1]
    try:
        number = float(text)
    except ValueError:
        return None
    if is_percent:
        number /= 100.0
    return number


def to_int(value: Any, fallback: int) -> int:
    number = to_float(value)
    if number is None:
        return fallback
    return int(round(number))


def pct_to_ratio(value: Any) -> Optional[float]:
    number = to_float(value)
    if number is None:
        return None
    if number > 1:
        return number / 100.0
    return number


def diff_or_none(a: Any, b: Any) -> Optional[float]:
    a_num = to_float(a)
    b_num = to_float(b)
    if a_num is None or b_num is None:
        return None
    return a_num - b_num


def first_not_none(*values: Any) -> Any:
    for value in values:
        if value is None:
            continue
        if isinstance(value, str) and not value.strip():
            continue
        return value
    return None


def parse_model_fields(file_name: str) -> Dict[str, str]:
    stem = Path(file_name).stem
    parts = [part.strip() for part in stem.split(" - ")]
    ticker = parts[1] if len(parts) > 1 else "UNKNOWN"
    period_chunk = parts[2] if len(parts) > 2 else ""
    period_token = period_chunk.split("_")[0].strip()

    model_period = period_token
    model_date = ""

    match = re.match(r"^(Early|Mid|Late)([A-Za-z]{3,9})(\d{4})$", period_token, flags=re.IGNORECASE)
    if match:
        phase, month_name, year = match.groups()
        phase_clean = phase[0].upper() + phase[1:].lower()
        month_key = month_name[:3].lower()
        month_num = MONTH_LOOKUP.get(month_key)
        if month_num:
            model_period = f"{phase_clean}{month_name[:3].title()}_{year}"
            day = PERIOD_DAY[phase.lower()]
            model_date = f"{int(year):04d}-{month_num:02d}-{day:02d}"

    return {
        "ticker": ticker,
        "model_period": model_period,
        "model_date": model_date,
        "model": f"{ticker}_{model_period}",
    }


def load_sheet_snapshot(sheet: xw.Sheet) -> SheetSnapshot:
    used = sheet.used_range
    values = to_2d(used.value)
    return SheetSnapshot(start_row=used.row, start_col=used.column, values=values)


def snapshot_value(snapshot: SheetSnapshot, row: int, col: int) -> Any:
    row_idx = row - snapshot.start_row
    col_idx = col - snapshot.start_col
    if row_idx < 0 or col_idx < 0:
        return None
    if row_idx >= len(snapshot.values):
        return None
    current_row = snapshot.values[row_idx]
    if col_idx >= len(current_row):
        return None
    return current_row[col_idx]


def find_max_anchor(snapshot: SheetSnapshot) -> Optional[Tuple[int, int]]:
    candidates: List[Tuple[int, int, int]] = []
    for row_idx, row in enumerate(snapshot.values):
        for col_idx, value in enumerate(row):
            if normalize_text(value) != "max":
                continue
            abs_row = snapshot.start_row + row_idx
            abs_col = snapshot.start_col + col_idx
            right_cell = snapshot_value(snapshot, abs_row, abs_col + 1)
            bonus = 0 if normalize_text(right_cell) == "min" else 1
            candidates.append((bonus, abs_row, abs_col))
    if not candidates:
        return None
    candidates.sort(key=lambda item: (item[0], item[1], item[2]))
    _, row, col = candidates[0]
    return row, col


def locate_header_columns(
    snapshot: SheetSnapshot,
    anchor_row: int,
    anchor_col: int,
    row_window: int = 3,
    col_window: int = 30,
) -> Dict[str, int]:
    result: Dict[str, Tuple[int, int]] = {}
    row_start = max(snapshot.start_row, anchor_row - row_window)
    row_end = min(snapshot.end_row, anchor_row + row_window)
    col_start = max(snapshot.start_col, anchor_col - col_window)
    col_end = min(snapshot.end_col, anchor_col + col_window)

    for row in range(row_start, row_end + 1):
        for col in range(col_start, col_end + 1):
            text = normalize_text(snapshot_value(snapshot, row, col))
            if not text:
                continue
            for key, patterns in HEADER_PATTERNS.items():
                if not any(pattern in text for pattern in patterns):
                    continue
                distance = abs(row - anchor_row) + abs(col - anchor_col)
                prev = result.get(key)
                if prev is None or distance < prev[0]:
                    result[key] = (distance, col)

    resolved: Dict[str, int] = {key: value[1] for key, value in result.items()}
    if "forecast_max" not in resolved:
        resolved["forecast_max"] = anchor_col
    if "forecast_min" not in resolved:
        right_cell = normalize_text(snapshot_value(snapshot, anchor_row, anchor_col + 1))
        resolved["forecast_min"] = anchor_col + 1 if right_cell == "min" else anchor_col - 1
    return resolved


def with_fallback_cols(
    header_cols: Dict[str, int], anchor_col: int, fallback_offsets: Dict[str, int]
) -> Dict[str, int]:
    resolved = dict(header_cols)
    for key, offset in fallback_offsets.items():
        resolved.setdefault(key, anchor_col + offset)
    return resolved


def collect_candidate_rows(
    snapshot: SheetSnapshot,
    anchor_row: int,
    num_col: int,
    anchor_col: int,
    limit: int = N_QUARTERS,
) -> List[int]:
    rows: List[int] = []
    blank_run = 0
    search_end = min(snapshot.end_row, anchor_row + 60)

    for row in range(anchor_row + 1, search_end + 1):
        value = snapshot_value(snapshot, row, num_col)
        if is_number(value):
            rows.append(row)
            blank_run = 0
        else:
            if rows:
                blank_run += 1
                if blank_run >= 2:
                    break
        if len(rows) >= limit:
            return rows[:limit]

    if rows:
        return rows[:limit]

    for row in range(anchor_row + 1, min(anchor_row + limit, snapshot.end_row) + 1):
        values = [
            snapshot_value(snapshot, row, col)
            for col in range(max(snapshot.start_col, anchor_col - 2), min(snapshot.end_col, anchor_col + 2) + 1)
        ]
        if any(value is not None and str(value).strip() != "" for value in values):
            rows.append(row)

    return rows[:limit]


def set_formula2(cell: xw.Range, formula: str) -> None:
    try:
        cell.formula2 = formula
    except Exception:
        cell.formula = formula


def close_workbook_no_save(wb: xw.Book) -> None:
    try:
        wb.close(save=False)
        return
    except TypeError:
        pass
    except Exception:
        pass

    fallback_closers = (
        lambda: wb.close(False),
        lambda: wb.api.Close(SaveChanges=False),
        lambda: wb.api.Close(False),
    )
    for closer in fallback_closers:
        try:
            closer()
            return
        except Exception:
            continue


def extract_empirical_rows(wb: xw.Book, model_fields: Dict[str, str], source_file: str) -> List[Dict[str, Any]]:
    if "Empirical Model" not in [sheet.name for sheet in wb.sheets]:
        return []

    sheet = wb.sheets["Empirical Model"]
    snapshot = load_sheet_snapshot(sheet)
    anchor = find_max_anchor(snapshot)
    if anchor is None:
        return []
    anchor_row, anchor_col = anchor

    header_cols = locate_header_columns(snapshot, anchor_row, anchor_col)
    cols = with_fallback_cols(header_cols, anchor_col, EMPIRICAL_FALLBACK_OFFSETS)
    candidate_rows = collect_candidate_rows(snapshot, anchor_row, cols["num_quarters_used"], anchor_col)
    if not candidate_rows:
        return []

    helper_col = sheet.used_range.last_cell.column + 2
    formula_rows: List[int] = []
    for idx, row in enumerate(candidate_rows):
        num_quarters = to_int(snapshot_value(snapshot, row, cols["num_quarters_used"]), idx + 1)
        start_row = max(snapshot.start_row, row - num_quarters + 1)
        formula = (
            f"=AVERAGE(R{start_row}C{cols['sales_captured_in_db_pct']}:"
            f"R{row}C{cols['sales_captured_in_db_pct']})"
        )
        set_formula2(sheet.range((row, helper_col)), formula)
        formula_rows.append(row)

    avg_penetration: Dict[int, Any] = {}
    if formula_rows:
        wb.app.calculate()
        avg_values = sheet.range((formula_rows[0], helper_col), (formula_rows[-1], helper_col)).value
        avg_values_list = to_1d(avg_values, len(formula_rows))
        for row, value in zip(formula_rows, avg_values_list):
            avg_penetration[row] = value
        sheet.range((formula_rows[0], helper_col), (formula_rows[-1], helper_col)).clear_contents()

    rows: List[Dict[str, Any]] = []
    for idx, row in enumerate(candidate_rows):
        num_quarters = to_int(snapshot_value(snapshot, row, cols["num_quarters_used"]), idx + 1)
        avg_pen = first_not_none(avg_penetration.get(row), snapshot_value(snapshot, row, cols["avg_penetration_pct"]))
        forecast_value = snapshot_value(snapshot, row, cols["forecast_value"])
        actual_value = first_not_none(
            snapshot_value(snapshot, row, cols["actual_value"]),
            snapshot_value(snapshot, row, cols["reported_sales"]),
        )
        reported_sales = snapshot_value(snapshot, row, cols["reported_sales"])

        if forecast_value is None:
            reported_num = to_float(reported_sales)
            avg_ratio = pct_to_ratio(avg_pen)
            if reported_num is not None and avg_ratio not in (None, 0):
                forecast_value = reported_num / avg_ratio

        forecast_max = snapshot_value(snapshot, row, cols["forecast_max"])
        forecast_min = snapshot_value(snapshot, row, cols["forecast_min"])

        row_dict = {
            "model": model_fields["model"],
            "ticker": model_fields["ticker"],
            "model_period": model_fields["model_period"],
            "model_date": model_fields["model_date"],
            "method": "empirical",
            "parameter_name": "avg_penetration_pct",
            "parameter_value": avg_pen,
            "num_quarters_used": num_quarters,
            "last_quarter_used": snapshot_value(snapshot, row, cols["last_quarter_used"]),
            "forecast_value": forecast_value,
            "actual_value": actual_value,
            "forecast_max": forecast_max,
            "forecast_min": forecast_min,
            "range_width": diff_or_none(forecast_max, forecast_min),
            "avg_penetration_pct": avg_pen,
            "quarterly_sales": snapshot_value(snapshot, row, cols["quarterly_sales"]),
            "reported_sales": reported_sales,
            "growth_rate_pct": snapshot_value(snapshot, row, cols["growth_rate_pct"]),
            "sales_captured_in_db_pct": snapshot_value(snapshot, row, cols["sales_captured_in_db_pct"]),
            "source_file": source_file,
        }
        rows.append(row_dict)

    return rows


def find_xy_data_block(
    snapshot: SheetSnapshot, anchor_row: int, x_col: int, y_col: int
) -> Optional[Tuple[int, int]]:
    row = anchor_row - 1
    while row >= snapshot.start_row:
        x_val = snapshot_value(snapshot, row, x_col)
        y_val = snapshot_value(snapshot, row, y_col)
        if is_number(x_val) and is_number(y_val):
            break
        row -= 1

    if row < snapshot.start_row:
        return None

    end_row = row
    start_row = row
    while start_row - 1 >= snapshot.start_row:
        x_val = snapshot_value(snapshot, start_row - 1, x_col)
        y_val = snapshot_value(snapshot, start_row - 1, y_col)
        if not (is_number(x_val) and is_number(y_val)):
            break
        start_row -= 1

    return start_row, end_row


def extract_regression_rows(wb: xw.Book, model_fields: Dict[str, str], source_file: str) -> List[Dict[str, Any]]:
    if "Regression Model" not in [sheet.name for sheet in wb.sheets]:
        return []

    sheet = wb.sheets["Regression Model"]
    snapshot = load_sheet_snapshot(sheet)
    anchor = find_max_anchor(snapshot)
    if anchor is None:
        return []
    anchor_row, anchor_col = anchor

    header_cols = locate_header_columns(snapshot, anchor_row, anchor_col)
    cols = with_fallback_cols(header_cols, anchor_col, REGRESSION_FALLBACK_OFFSETS)

    y_col = anchor_col - 7
    x_col = anchor_col - 11
    xy_block = find_xy_data_block(snapshot, anchor_row, x_col, y_col)
    if xy_block is None:
        return []
    data_start, data_end = xy_block

    available_points = data_end - data_start + 1
    if available_points < 2:
        return []

    max_rows = min(N_QUARTERS, available_points)
    helper_col = sheet.used_range.last_cell.column + 4
    intercept_col = helper_col
    slope_col = helper_col + 1

    for n in range(1, max_rows + 1):
        calc_row = anchor_row + n
        start_row = data_end - n + 1
        intercept_formula = (
            f"=INTERCEPT(R{start_row}C{y_col}:R{data_end}C{y_col},"
            f"R{start_row}C{x_col}:R{data_end}C{x_col})"
        )
        slope_formula = (
            f"=SLOPE(R{start_row}C{y_col}:R{data_end}C{y_col},"
            f"R{start_row}C{x_col}:R{data_end}C{x_col})"
        )
        set_formula2(sheet.range((calc_row, intercept_col)), intercept_formula)
        set_formula2(sheet.range((calc_row, slope_col)), slope_formula)

    wb.app.calculate()

    intercept_values = sheet.range((anchor_row + 1, intercept_col), (anchor_row + max_rows, intercept_col)).value
    slope_values = sheet.range((anchor_row + 1, slope_col), (anchor_row + max_rows, slope_col)).value
    intercept_list = to_1d(intercept_values, max_rows)
    slope_list = to_1d(slope_values, max_rows)
    sheet.range((anchor_row + 1, intercept_col), (anchor_row + max_rows, slope_col)).clear_contents()

    forecast_x = to_float(snapshot_value(snapshot, data_end + 1, x_col))
    if forecast_x is None:
        last_x = to_float(snapshot_value(snapshot, data_end, x_col))
        forecast_x = None if last_x is None else (last_x + 1.0)

    rows: List[Dict[str, Any]] = []
    previous_signature: Optional[Tuple[Any, ...]] = None
    for idx in range(max_rows):
        n = idx + 1
        table_row = anchor_row + n
        num_quarters = to_int(sheet.range((table_row, cols["num_quarters_used"])).value, n)
        intercept = intercept_list[idx]
        slope = slope_list[idx]
        forecast_value = sheet.range((table_row, cols["forecast_value"])).value
        if forecast_value is None:
            intercept_num = to_float(intercept)
            slope_num = to_float(slope)
            if intercept_num is not None and slope_num is not None and forecast_x is not None:
                forecast_value = intercept_num + (slope_num * forecast_x)

        forecast_max = sheet.range((table_row, cols["forecast_max"])).value
        forecast_min = sheet.range((table_row, cols["forecast_min"])).value
        actual_value = sheet.range((table_row, cols["actual_value"])).value if "actual_value" in cols else None
        range_width = diff_or_none(forecast_max, forecast_min)

        signature = (
            round(to_float(forecast_value) or 0.0, 6),
            round(to_float(forecast_max) or 0.0, 6),
            round(to_float(forecast_min) or 0.0, 6),
            round(to_float(intercept) or 0.0, 6),
            round(to_float(slope) or 0.0, 6),
        )
        if signature == previous_signature:
            continue
        previous_signature = signature

        row_dict = {
            "model": model_fields["model"],
            "ticker": model_fields["ticker"],
            "model_period": model_fields["model_period"],
            "model_date": model_fields["model_date"],
            "method": "regression",
            "parameter_name": "num_quarters_used",
            "parameter_value": num_quarters,
            "num_quarters_used": num_quarters,
            "forecast_value": forecast_value,
            "actual_value": actual_value,
            "forecast_max": forecast_max,
            "forecast_min": forecast_min,
            "range_width": range_width,
            "intercept": intercept,
            "slope": slope,
            "source_file": source_file,
        }
        rows.append(row_dict)

    return rows


def resolve_output_path(input_path: Path, output_path: Path) -> Path:
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


def write_table(
    worksheet,
    columns: Sequence[str],
    rows: Iterable[Dict[str, Any]],
) -> None:
    worksheet.append(list(columns))
    for row in rows:
        worksheet.append([row.get(column) for column in columns])

    for cell in worksheet[1]:
        cell.font = Font(bold=True)

    worksheet.freeze_panes = "A2"
    worksheet.auto_filter.ref = worksheet.dimensions

    for idx, column_name in enumerate(columns, start=1):
        max_len = len(column_name)
        for value, in worksheet.iter_rows(
            min_row=2,
            max_row=worksheet.max_row,
            min_col=idx,
            max_col=idx,
            values_only=True,
        ):
            if value is None:
                continue
            max_len = max(max_len, len(str(value)))
        worksheet.column_dimensions[get_column_letter(idx)].width = min(max(12, max_len + 2), 48)


def write_output_workbook(
    target_path: Path,
    empirical_rows: List[Dict[str, Any]],
    regression_rows: List[Dict[str, Any]],
) -> None:
    wb = Workbook()
    empirical_sheet = wb.active
    empirical_sheet.title = "empirical_candidates"
    write_table(empirical_sheet, EMPIRICAL_COLUMNS, empirical_rows)

    regression_sheet = wb.create_sheet("regression_candidates")
    write_table(regression_sheet, REGRESSION_COLUMNS, regression_rows)
    wb.save(target_path)


def process_file(app: xw.App, file_path: Path) -> Tuple[List[Dict[str, Any]], List[Dict[str, Any]], bool]:
    wb: Optional[xw.Book] = None
    try:
        wb = app.books.open(str(file_path), update_links=False)
        model_fields = parse_model_fields(file_path.name)
        empirical_rows = extract_empirical_rows(wb, model_fields, file_path.name)
        regression_rows = extract_regression_rows(wb, model_fields, file_path.name)
        print(f"processed: {file_path.name}")
        return empirical_rows, regression_rows, True
    except Exception as exc:
        print(f"skipped: {file_path.name} (process error: {exc})")
        return [], [], False
    finally:
        if wb is not None:
            close_workbook_no_save(wb)


def main() -> None:
    source_dir = input_dir.expanduser().resolve()
    target_dir = output_dir.expanduser().resolve()

    if not source_dir.exists():
        raise SystemExit(f"input_dir does not exist: {source_dir}")
    target_dir.mkdir(parents=True, exist_ok=True)

    output_path = resolve_output_path(source_dir, target_dir)

    empirical_rows: List[Dict[str, Any]] = []
    regression_rows: List[Dict[str, Any]] = []
    processed_files = 0

    app = xw.App(visible=False, add_book=False)
    app.display_alerts = False
    app.screen_updating = False

    try:
        for file_path in sorted(source_dir.iterdir(), key=lambda item: item.name.lower()):
            if not file_path.is_file():
                continue
            if file_path.name.startswith("~"):
                print(f"skipped: {file_path.name} (temporary file)")
                continue
            if file_path.suffix.lower() != ".xlsx":
                print(f"skipped: {file_path.name} (not .xlsx)")
                continue

            file_empirical, file_regression, ok = process_file(app, file_path)
            if not ok:
                continue
            processed_files += 1
            empirical_rows.extend(file_empirical)
            regression_rows.extend(file_regression)
    finally:
        app.quit()

    write_output_workbook(output_path, empirical_rows, regression_rows)
    print(f"output_path: {output_path}")
    print(f"files_processed: {processed_files}")
    print(f"empirical_rows: {len(empirical_rows)}")
    print(f"regression_rows: {len(regression_rows)}")


if __name__ == "__main__":
    main()
