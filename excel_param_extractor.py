from __future__ import annotations

import math
import re
from dataclasses import dataclass
from datetime import date
from pathlib import Path
from statistics import pstdev
from typing import Any, Dict, List, Optional, Sequence, Tuple

import xlwings as xw
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

# Update these two paths before running.
input_dir = Path("input")
output_dir = Path("output")

N_QUARTERS = 10
EMPIRICAL_SHEET_NAME = "Empirical Model"
REGRESSION_SHEET_NAME = "Regression Model"

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

# Anchor-based defaults; if a label exists, label lookup is used as fallback.
EMPIRICAL_OFFSETS = {
    "avg_penetration_input": (-6, 1),
    "estimated_total_sold": (-4, 1),
    "reported_sales": (-3, 1),
    "forecast_max": (0, 1),
    "forecast_min": (1, 1),
    "quarterly_sales": (-2, 1),
    "growth_rate_pct": (-1, 1),
    "sales_captured_in_db_pct": (-5, 1),
}

REGRESSION_OFFSETS = {
    "num_quarters_input": (-6, 1),
    "forecast_total_without_sa": (-4, 1),
    "actual_value": (-3, 1),
    "forecast_max": (0, 1),
    "forecast_min": (1, 1),
}

DAY_BY_PHASE = {"early": 5, "mid": 15, "late": 25}
MONTH_INDEX = {
    "jan": 1,
    "january": 1,
    "feb": 2,
    "february": 2,
    "mar": 3,
    "march": 3,
    "apr": 4,
    "april": 4,
    "may": 5,
    "jun": 6,
    "june": 6,
    "jul": 7,
    "july": 7,
    "aug": 8,
    "august": 8,
    "sep": 9,
    "sept": 9,
    "september": 9,
    "oct": 10,
    "october": 10,
    "nov": 11,
    "november": 11,
    "dec": 12,
    "december": 12,
}
PERIOD_RE = re.compile(r"(Early|Mid|Late)\s*([A-Za-z]{3,9})\s*(20\d{2})", re.IGNORECASE)
LABEL_STRIP_RE = re.compile(r"[^a-z0-9]+")


@dataclass
class SheetContext:
    sheet: xw.main.Sheet
    anchor_row: int
    anchor_col: int
    labels: Dict[str, Tuple[int, int]]


def normalize_label(value: Any) -> str:
    return LABEL_STRIP_RE.sub("", str(value).strip().lower()) if value is not None else ""


def is_number(value: Any) -> bool:
    if isinstance(value, bool):
        return False
    if isinstance(value, (int, float)):
        try:
            return math.isfinite(float(value))
        except (TypeError, ValueError):
            return False
    return False


def float_or_none(value: Any) -> Optional[float]:
    return float(value) if is_number(value) else None


def non_empty(value: Any) -> bool:
    return value is not None and value != ""


def parse_filename_metadata(file_path: Path) -> Dict[str, str]:
    stem = file_path.stem
    parts = [part.strip() for part in stem.split(" - ") if part.strip()]

    ticker = ""
    if len(parts) >= 2 and re.fullmatch(r"[A-Za-z0-9._-]+", parts[1]):
        ticker = parts[1].upper()
    else:
        ticker_match = re.search(r"\b[A-Z]{2,8}\b", stem)
        ticker = ticker_match.group(0) if ticker_match else "UNKNOWN"

    period_match = PERIOD_RE.search(stem)
    if period_match:
        phase = period_match.group(1).title()
        month_token = period_match.group(2).lower()
        year = int(period_match.group(3))
        month_num = MONTH_INDEX.get(month_token, MONTH_INDEX.get(month_token[:3], 1))
    else:
        phase = "Mid"
        year_match = re.search(r"(20\d{2})", stem)
        year = int(year_match.group(1)) if year_match else 1900
        month_num = 1

    month_abbrev = date(year, month_num, 1).strftime("%b")
    model_period = f"{phase}{month_abbrev}_{year}"
    model_date = date(year, month_num, DAY_BY_PHASE[phase.lower()]).isoformat()

    return {
        "ticker": ticker,
        "model_period": model_period,
        "model_date": model_date,
        "model": f"{ticker}_{model_period}",
    }


def unique_output_path(in_dir: Path, out_dir: Path) -> Path:
    out_dir.mkdir(parents=True, exist_ok=True)
    base_name = f"{in_dir.name}_PARAM"
    candidate = out_dir / f"{base_name}.xlsx"
    counter = 1
    while candidate.exists():
        candidate = out_dir / f"{base_name}.{counter}.xlsx"
        counter += 1
    return candidate


def ensure_2d(values: Any) -> List[List[Any]]:
    if values is None:
        return []
    if isinstance(values, tuple):
        values = list(values)
    if not isinstance(values, list):
        return [[values]]
    if values and isinstance(values[0], tuple):
        return [list(row) for row in values]
    if values and not isinstance(values[0], list):
        return [values]
    return values


def build_label_index(start_row: int, start_col: int, grid: List[List[Any]]) -> Dict[str, Tuple[int, int]]:
    labels: Dict[str, Tuple[int, int]] = {}
    for r_idx, row_values in enumerate(grid):
        for c_idx, value in enumerate(row_values):
            if isinstance(value, str):
                normalized = normalize_label(value)
                if normalized and normalized not in labels:
                    labels[normalized] = (start_row + r_idx, start_col + c_idx)
    return labels


def find_max_anchor(start_row: int, start_col: int, grid: List[List[Any]]) -> Tuple[int, int]:
    candidates: List[Tuple[int, int, int]] = []
    for r_idx, row_values in enumerate(grid):
        for c_idx, value in enumerate(row_values):
            if normalize_label(value) != "max":
                continue
            score = 0
            if r_idx + 1 < len(grid) and c_idx < len(grid[r_idx + 1]):
                if normalize_label(grid[r_idx + 1][c_idx]) == "min":
                    score += 10
            candidates.append((score, start_row + r_idx, start_col + c_idx))

    if not candidates:
        raise ValueError("Could not find 'max' anchor.")

    candidates.sort(reverse=True)
    return candidates[0][1], candidates[0][2]


def build_sheet_context(sheet: xw.main.Sheet) -> SheetContext:
    used = sheet.used_range
    start_row, start_col = used.row, used.column
    grid = ensure_2d(used.value)
    labels = build_label_index(start_row, start_col, grid)
    anchor_row, anchor_col = find_max_anchor(start_row, start_col, grid)
    return SheetContext(sheet=sheet, anchor_row=anchor_row, anchor_col=anchor_col, labels=labels)


def find_label_position(
    labels: Dict[str, Tuple[int, int]],
    hints: Sequence[str],
) -> Optional[Tuple[int, int]]:
    normalized_hints = [normalize_label(item) for item in hints]
    for hint in normalized_hints:
        if not hint:
            continue
        for label_key, position in labels.items():
            if hint in label_key:
                return position
    return None


def resolve_cell(
    ctx: SheetContext,
    offset: Tuple[int, int],
    label_hints: Sequence[str],
) -> Tuple[int, int]:
    row = ctx.anchor_row + offset[0]
    col = ctx.anchor_col + offset[1]
    current_value = ctx.sheet.range((row, col)).value
    if non_empty(current_value):
        return row, col

    label_pos = find_label_position(ctx.labels, label_hints)
    if label_pos:
        return label_pos[0], label_pos[1] + 1
    return row, col


def collect_numeric_column_points(
    sheet: xw.main.Sheet,
    column: int,
    start_row: int,
    max_scan_rows: int = 400,
) -> List[Tuple[int, float]]:
    points: List[Tuple[int, float]] = []
    row = start_row
    seen_numeric = False
    misses_after_data = 0
    scanned = 0

    while row >= 1 and scanned < max_scan_rows:
        scanned += 1
        value = sheet.range((row, column)).value
        if is_number(value):
            points.append((row, float(value)))
            seen_numeric = True
            misses_after_data = 0
        elif seen_numeric:
            misses_after_data += 1
            if misses_after_data >= 3:
                break
        row -= 1

    points.reverse()
    return points


def collect_xy_points(
    sheet: xw.main.Sheet,
    x_col: int,
    y_col: int,
    start_row: int,
    max_scan_rows: int = 400,
) -> List[Tuple[int, float, float]]:
    points: List[Tuple[int, float, float]] = []
    row = start_row
    seen_numeric = False
    misses_after_data = 0
    scanned = 0

    while row >= 1 and scanned < max_scan_rows:
        scanned += 1
        x_val = sheet.range((row, x_col)).value
        y_val = sheet.range((row, y_col)).value

        if is_number(x_val) and is_number(y_val):
            points.append((row, float(x_val), float(y_val)))
            seen_numeric = True
            misses_after_data = 0
        elif seen_numeric:
            misses_after_data += 1
            if misses_after_data >= 3:
                break
        row -= 1

    points.reverse()
    return points


def r1c1_abs(row: int, col: int) -> str:
    return f"R{row}C{col}"


def numeric_width(high_value: Any, low_value: Any) -> Optional[float]:
    hi = float_or_none(high_value)
    lo = float_or_none(low_value)
    if hi is None or lo is None:
        return None
    return hi - lo


def safe_close_workbook(wb: xw.main.Book) -> None:
    try:
        wb.close(save=False)
        return
    except TypeError:
        pass
    except Exception:
        pass

    try:
        wb.api.Close(SaveChanges=False)
        return
    except Exception:
        pass

    try:
        wb.close(False)
    except Exception:
        wb.close()


def extract_empirical_rows(
    wb: xw.main.Book,
    metadata: Dict[str, str],
    source_file: str,
) -> List[Dict[str, Any]]:
    try:
        sheet = wb.sheets[EMPIRICAL_SHEET_NAME]
    except Exception:
        print(f"skipped empirical extraction: {source_file} (missing sheet '{EMPIRICAL_SHEET_NAME}')")
        return []

    try:
        ctx = build_sheet_context(sheet)
    except Exception as exc:
        print(f"skipped empirical extraction: {source_file} ({exc})")
        return []

    avg_pen_cell = resolve_cell(
        ctx,
        EMPIRICAL_OFFSETS["avg_penetration_input"],
        ["avg penetration", "avg_penetration_pct", "average penetration"],
    )
    est_total_cell = resolve_cell(
        ctx,
        EMPIRICAL_OFFSETS["estimated_total_sold"],
        ["estimated total sold", "est total sold", "forecast total"],
    )
    reported_sales_cell = resolve_cell(
        ctx,
        EMPIRICAL_OFFSETS["reported_sales"],
        ["reported sales", "actual sales"],
    )
    forecast_max_cell = resolve_cell(ctx, EMPIRICAL_OFFSETS["forecast_max"], ["max"])
    forecast_min_cell = resolve_cell(ctx, EMPIRICAL_OFFSETS["forecast_min"], ["min"])
    quarterly_sales_cell = resolve_cell(
        ctx,
        EMPIRICAL_OFFSETS["quarterly_sales"],
        ["quarterly sales", "quarter sales"],
    )
    growth_cell = resolve_cell(ctx, EMPIRICAL_OFFSETS["growth_rate_pct"], ["growth rate"])
    captured_cell = resolve_cell(
        ctx,
        EMPIRICAL_OFFSETS["sales_captured_in_db_pct"],
        ["sales captured in db", "captured in db", "capture pct"],
    )

    penetration_col = ctx.anchor_col - 11
    sales_col = ctx.anchor_col - 7
    quarter_label_col = penetration_col - 1

    penetration_points = collect_numeric_column_points(sheet, penetration_col, ctx.anchor_row - 1)
    if not penetration_points:
        print(f"skipped empirical extraction: {source_file} (no penetration series near anchor)")
        return []

    sales_points = collect_numeric_column_points(sheet, sales_col, ctx.anchor_row - 1)
    sales_by_row = {row: val for row, val in sales_points}

    output_rows: List[Dict[str, Any]] = []
    max_steps = min(N_QUARTERS, len(penetration_points))
    for num_quarters in range(1, max_steps + 1):
        start_row = penetration_points[-num_quarters][0]
        end_row = penetration_points[-1][0]

        # Temporary formula write is intentionally used to trigger workbook-native calculations.
        avg_formula = f"=AVERAGE({r1c1_abs(start_row, penetration_col)}:{r1c1_abs(end_row, penetration_col)})"
        sheet.range(avg_pen_cell).formula2 = avg_formula
        wb.app.calculate()

        avg_pen = sheet.range(avg_pen_cell).value
        forecast_value = sheet.range(est_total_cell).value
        actual_value = sheet.range(reported_sales_cell).value
        forecast_max = sheet.range(forecast_max_cell).value
        forecast_min = sheet.range(forecast_min_cell).value
        quarterly_sales = sheet.range(quarterly_sales_cell).value
        growth_rate = sheet.range(growth_cell).value
        sales_captured = sheet.range(captured_cell).value

        if not is_number(quarterly_sales):
            quarterly_sales = sales_by_row.get(end_row)

        if not is_number(growth_rate):
            selected_sales = [sales_by_row[row] for row, _ in penetration_points[-num_quarters:] if row in sales_by_row]
            if len(selected_sales) >= 2 and selected_sales[-2] != 0:
                growth_rate = ((selected_sales[-1] / selected_sales[-2]) - 1.0) * 100.0

        if not is_number(sales_captured):
            q_sales_num = float_or_none(quarterly_sales)
            reported_num = float_or_none(actual_value)
            if q_sales_num is not None and reported_num not in (None, 0.0):
                sales_captured = (q_sales_num / reported_num) * 100.0

        last_quarter_used = sheet.range((start_row, quarter_label_col)).value
        if not non_empty(last_quarter_used):
            last_quarter_used = start_row

        output_rows.append(
            {
                "model": metadata["model"],
                "ticker": metadata["ticker"],
                "model_period": metadata["model_period"],
                "model_date": metadata["model_date"],
                "method": "empirical",
                "parameter_name": "avg_penetration_pct",
                "parameter_value": avg_pen,
                "num_quarters_used": num_quarters,
                "last_quarter_used": last_quarter_used,
                "forecast_value": forecast_value,
                "actual_value": actual_value,
                "forecast_max": forecast_max,
                "forecast_min": forecast_min,
                "range_width": numeric_width(forecast_max, forecast_min),
                "avg_penetration_pct": avg_pen,
                "quarterly_sales": quarterly_sales,
                "reported_sales": actual_value,
                "growth_rate_pct": growth_rate,
                "sales_captured_in_db_pct": sales_captured,
                "source_file": source_file,
            }
        )

    return output_rows


def extract_regression_rows(
    wb: xw.main.Book,
    metadata: Dict[str, str],
    source_file: str,
) -> List[Dict[str, Any]]:
    try:
        sheet = wb.sheets[REGRESSION_SHEET_NAME]
    except Exception:
        print(f"skipped regression extraction: {source_file} (missing sheet '{REGRESSION_SHEET_NAME}')")
        return []

    try:
        ctx = build_sheet_context(sheet)
    except Exception as exc:
        print(f"skipped regression extraction: {source_file} ({exc})")
        return []

    y_col = ctx.anchor_col - 7
    x_col = ctx.anchor_col - 11
    xy_points = collect_xy_points(sheet, x_col, y_col, ctx.anchor_row - 1)
    if len(xy_points) < 2:
        print(f"skipped regression extraction: {source_file} (insufficient x/y points near anchor)")
        return []

    num_quarters_input_cell = resolve_cell(
        ctx,
        REGRESSION_OFFSETS["num_quarters_input"],
        ["num quarters used", "quarters used"],
    )
    forecast_value_cell = resolve_cell(
        ctx,
        REGRESSION_OFFSETS["forecast_total_without_sa"],
        ["tot fcst w/o sa", "tot fcst wo sa", "forecast without sa"],
    )
    actual_value_cell = resolve_cell(ctx, REGRESSION_OFFSETS["actual_value"], ["reported sales", "actual sales"])
    forecast_max_cell = resolve_cell(ctx, REGRESSION_OFFSETS["forecast_max"], ["max"])
    forecast_min_cell = resolve_cell(ctx, REGRESSION_OFFSETS["forecast_min"], ["min"])

    helper_col = ctx.anchor_col + 3
    helper_row = ctx.anchor_row + 2
    intercept_cell = (helper_row, helper_col)
    slope_cell = (helper_row + 1, helper_col)
    helper_forecast_cell = (helper_row + 2, helper_col)

    output_rows: List[Dict[str, Any]] = []
    previous_key: Optional[Tuple[Optional[float], Optional[float], Optional[float], Optional[float], Optional[float]]] = None

    max_steps = min(N_QUARTERS, len(xy_points))
    for num_quarters in range(2, max_steps + 1):
        selected_points = xy_points[-num_quarters:]
        start_row = selected_points[0][0]
        end_row = selected_points[-1][0]

        sheet.range(num_quarters_input_cell).value = num_quarters

        intercept_formula = (
            f"=INTERCEPT({r1c1_abs(start_row, y_col)}:{r1c1_abs(end_row, y_col)},"
            f"{r1c1_abs(start_row, x_col)}:{r1c1_abs(end_row, x_col)})"
        )
        slope_formula = (
            f"=SLOPE({r1c1_abs(start_row, y_col)}:{r1c1_abs(end_row, y_col)},"
            f"{r1c1_abs(start_row, x_col)}:{r1c1_abs(end_row, x_col)})"
        )

        forecast_x_row = ctx.anchor_row if is_number(sheet.range((ctx.anchor_row, x_col)).value) else end_row
        forecast_formula = (
            f"={r1c1_abs(intercept_cell[0], intercept_cell[1])}"
            f"+{r1c1_abs(slope_cell[0], slope_cell[1])}*{r1c1_abs(forecast_x_row, x_col)}"
        )

        sheet.range(intercept_cell).formula2 = intercept_formula
        sheet.range(slope_cell).formula2 = slope_formula
        sheet.range(helper_forecast_cell).formula2 = forecast_formula
        wb.app.calculate()

        intercept = sheet.range(intercept_cell).value
        slope = sheet.range(slope_cell).value

        forecast_value = sheet.range(forecast_value_cell).value
        if not is_number(forecast_value):
            forecast_value = sheet.range(helper_forecast_cell).value

        forecast_max = sheet.range(forecast_max_cell).value
        forecast_min = sheet.range(forecast_min_cell).value
        actual_value = sheet.range(actual_value_cell).value

        if (not is_number(forecast_max) or not is_number(forecast_min)) and is_number(forecast_value):
            intercept_num = float_or_none(intercept)
            slope_num = float_or_none(slope)
            if intercept_num is not None and slope_num is not None:
                residuals = [y_val - (intercept_num + slope_num * x_val) for _, x_val, y_val in selected_points]
                spread = (2.0 * pstdev(residuals)) if len(residuals) > 1 else 0.0
                forecast_max = float(forecast_value) + spread
                forecast_min = float(forecast_value) - spread

        dedupe_key = (
            None if not is_number(forecast_value) else round(float(forecast_value), 8),
            None if not is_number(forecast_max) else round(float(forecast_max), 8),
            None if not is_number(forecast_min) else round(float(forecast_min), 8),
            None if not is_number(intercept) else round(float(intercept), 8),
            None if not is_number(slope) else round(float(slope), 8),
        )
        if dedupe_key == previous_key:
            continue
        previous_key = dedupe_key

        output_rows.append(
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
                "actual_value": actual_value if non_empty(actual_value) else "",
                "forecast_max": forecast_max,
                "forecast_min": forecast_min,
                "range_width": numeric_width(forecast_max, forecast_min),
                "intercept": intercept,
                "slope": slope,
                "source_file": source_file,
            }
        )

    return output_rows


def write_sheet(ws: Any, columns: Sequence[str], rows: List[Dict[str, Any]]) -> None:
    ws.append(list(columns))
    for row_data in rows:
        ws.append([row_data.get(column, "") if row_data.get(column, "") is not None else "" for column in columns])

    for cell in ws[1]:
        cell.font = Font(bold=True)

    ws.freeze_panes = "A2"
    last_col_letter = get_column_letter(len(columns))
    ws.auto_filter.ref = f"A1:{last_col_letter}{max(2, ws.max_row)}"

    for idx, column in enumerate(columns, start=1):
        max_len = len(column)
        for row_data in rows:
            value = row_data.get(column, "")
            value_len = len(str(value)) if value is not None else 0
            if value_len > max_len:
                max_len = value_len
        ws.column_dimensions[get_column_letter(idx)].width = min(max(12, max_len + 2), 44)


def write_output_workbook(
    output_path: Path,
    empirical_rows: List[Dict[str, Any]],
    regression_rows: List[Dict[str, Any]],
) -> None:
    out_wb = Workbook()
    default_sheet = out_wb.active
    out_wb.remove(default_sheet)

    emp_sheet = out_wb.create_sheet("empirical_candidates")
    reg_sheet = out_wb.create_sheet("regression_candidates")

    write_sheet(emp_sheet, EMPIRICAL_COLUMNS, empirical_rows)
    write_sheet(reg_sheet, REGRESSION_COLUMNS, regression_rows)

    out_wb.save(output_path)


def main() -> None:
    if not input_dir.exists():
        raise FileNotFoundError(f"input_dir does not exist: {input_dir}")

    output_path = unique_output_path(input_dir, output_dir)

    processed_files = 0
    empirical_rows: List[Dict[str, Any]] = []
    regression_rows: List[Dict[str, Any]] = []

    app = xw.App(visible=False, add_book=False)
    app.display_alerts = False
    app.screen_updating = False
    try:
        app.calculation = "manual"
    except Exception:
        pass

    try:
        for file_path in sorted(input_dir.iterdir()):
            if not file_path.is_file():
                continue

            file_name = file_path.name
            lower_name = file_name.lower()

            if file_name.startswith("~"):
                print(f"skipped: {file_name} (temporary file)")
                continue
            if not lower_name.endswith(".xlsx"):
                print(f"skipped: {file_name} (not an .xlsx file)")
                continue

            try:
                wb = app.books.open(str(file_path), update_links=False)
            except Exception as exc:
                print(f"skipped: {file_name} (open failed: {exc})")
                continue

            try:
                metadata = parse_filename_metadata(file_path)
                empirical_rows.extend(extract_empirical_rows(wb, metadata, file_name))
                regression_rows.extend(extract_regression_rows(wb, metadata, file_name))
                processed_files += 1
                print(f"processed: {file_name}")
            except Exception as exc:
                print(f"skipped: {file_name} (processing failed: {exc})")
            finally:
                safe_close_workbook(wb)
    finally:
        app.quit()

    write_output_workbook(output_path, empirical_rows, regression_rows)

    print(f"output_path: {output_path}")
    print(f"files_processed: {processed_files}")
    print(f"empirical_rows: {len(empirical_rows)}")
    print(f"regression_rows: {len(regression_rows)}")


if __name__ == "__main__":
    main()
