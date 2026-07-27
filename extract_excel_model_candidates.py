#!/usr/bin/env python3
"""Extract empirical and regression parameter candidates from Excel models.

This script scans all .xlsx files in input_dir and writes one consolidated
output workbook containing two sheets:
  - empirical_candidates
  - regression_candidates

Performance notes:
  - Uses one hidden Excel app for the full run.
  - Opens each source workbook exactly once, processes both target sheets
    while open, then closes the workbook without saving.
  - Uses R1C1 .formula2 formulas for AVG/INTERCEPT/SLOPE calculations.
"""

from __future__ import annotations

from dataclasses import dataclass
from datetime import date, datetime
import re
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Sequence, Tuple

try:
    import xlwings as xw
except ImportError as exc:
    raise SystemExit(
        "Missing dependency: xlwings. Install with `pip install xlwings`."
    ) from exc

try:
    from openpyxl import Workbook
    from openpyxl.styles import Font
    from openpyxl.utils import get_column_letter
except ImportError as exc:
    raise SystemExit(
        "Missing dependency: openpyxl. Install with `pip install openpyxl`."
    ) from exc

# -----------------------------
# User-configurable directories
# -----------------------------
input_dir = Path("./input")
output_dir = Path("./output")


EMPIRICAL_SHEET_NAME = "Empirical Model"
REGRESSION_SHEET_NAME = "Regression Model"
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

PERIOD_DAY = {"early": 5, "mid": 15, "late": 25}
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
    "sept": 9,
    "oct": 10,
    "nov": 11,
    "dec": 12,
}


@dataclass
class SheetSnapshot:
    start_row: int
    start_col: int
    end_row: int
    end_col: int
    values: List[List[Any]]
    text_cells: Dict[str, List[Tuple[int, int]]]


def log(message: str) -> None:
    print(message, flush=True)


def normalize_text(value: Any) -> str:
    if not isinstance(value, str):
        return ""
    return " ".join(value.strip().lower().split())


def to_2d(values: Any) -> List[List[Any]]:
    if values is None:
        return []
    if isinstance(values, (list, tuple)):
        if not values:
            return []
        first = values[0]
        if isinstance(first, (list, tuple)):
            return [list(row) for row in values]
        return [list(values)]
    return [[values]]


def to_float(value: Any) -> Optional[float]:
    if value is None or value == "":
        return None
    if isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        return float(value)
    if isinstance(value, str):
        cleaned = value.replace(",", "").replace("%", "").strip()
        if not cleaned:
            return None
        try:
            return float(cleaned)
        except ValueError:
            return None
    return None


def to_math_pct(value: Optional[float]) -> Optional[float]:
    if value is None:
        return None
    if abs(value) > 1.5:
        return value / 100.0
    return value


def make_sheet_snapshot(sheet: xw.main.Sheet) -> SheetSnapshot:
    used = sheet.used_range
    start_row = used.row
    start_col = used.column
    rows_count = used.rows.count
    cols_count = used.columns.count
    end_row = start_row + rows_count - 1
    end_col = start_col + cols_count - 1

    values = to_2d(used.value)
    text_cells: Dict[str, List[Tuple[int, int]]] = {}
    for r_idx, row in enumerate(values):
        for c_idx, cell_value in enumerate(row):
            key = normalize_text(cell_value)
            if key:
                text_cells.setdefault(key, []).append(
                    (start_row + r_idx, start_col + c_idx)
                )

    return SheetSnapshot(
        start_row=start_row,
        start_col=start_col,
        end_row=end_row,
        end_col=end_col,
        values=values,
        text_cells=text_cells,
    )


def snapshot_value(snapshot: SheetSnapshot, row: int, col: int) -> Any:
    if row < snapshot.start_row or row > snapshot.end_row:
        return None
    if col < snapshot.start_col or col > snapshot.end_col:
        return None
    r_idx = row - snapshot.start_row
    c_idx = col - snapshot.start_col
    if r_idx < 0 or c_idx < 0:
        return None
    if r_idx >= len(snapshot.values):
        return None
    row_values = snapshot.values[r_idx]
    if c_idx >= len(row_values):
        return None
    return row_values[c_idx]


def get_sheet_value(sheet: xw.main.Sheet, coord: Tuple[int, int]) -> Any:
    return sheet.range(coord).value


def get_sheet_float(sheet: xw.main.Sheet, coord: Optional[Tuple[int, int]]) -> Optional[float]:
    if coord is None:
        return None
    return to_float(get_sheet_value(sheet, coord))


def find_anchor_max(snapshot: SheetSnapshot) -> Optional[Tuple[int, int]]:
    for key, cells in snapshot.text_cells.items():
        if key == "max" or key.startswith("max "):
            return cells[0]
    # Fallback: fuzzy contains for uncommon formatting ("max:")
    for key, cells in snapshot.text_cells.items():
        if "max" == key.rstrip(":"):
            return cells[0]
    return None


def nearest_label(
    snapshot: SheetSnapshot,
    anchor_row: int,
    anchor_col: int,
    keywords: Sequence[str],
    max_distance: int = 100,
) -> Optional[Tuple[int, int, str]]:
    best: Optional[Tuple[int, int, int, str]] = None
    for text, cells in snapshot.text_cells.items():
        if not any(word in text for word in keywords):
            continue
        for row, col in cells:
            distance = abs(anchor_row - row) + abs(anchor_col - col)
            if distance > max_distance:
                continue
            if best is None or distance < best[0]:
                best = (distance, row, col, text)
    if best is None:
        return None
    return best[1], best[2], best[3]


def find_min_value_col(snapshot: SheetSnapshot, anchor_row: int, anchor_col: int) -> int:
    # First preference: a nearby "min" label.
    min_label = nearest_label(snapshot, anchor_row, anchor_col, ["min"], max_distance=16)
    if min_label and abs(min_label[0] - anchor_row) <= 2:
        return min_label[1]
    # Fallback to an adjacent column if label not found.
    return anchor_col + 1


def resolve_metric_cell(
    snapshot: SheetSnapshot,
    sheet: xw.main.Sheet,
    anchor_row: int,
    anchor_col: int,
    keywords: Sequence[str],
    fallback_offset: Optional[Tuple[int, int]] = None,
    allow_non_numeric: bool = False,
) -> Optional[Tuple[int, int]]:
    label = nearest_label(snapshot, anchor_row, anchor_col, keywords)
    if label:
        label_row, label_col, _ = label
        probes = [(0, 1), (1, 0), (0, 2), (2, 0), (0, -1), (-1, 0)]
        first_probe: Optional[Tuple[int, int]] = None
        for dr, dc in probes:
            row = label_row + dr
            col = label_col + dc
            if first_probe is None:
                first_probe = (row, col)
            value = snapshot_value(snapshot, row, col)
            if to_float(value) is not None:
                return row, col
        if allow_non_numeric and first_probe is not None:
            return first_probe

    if fallback_offset is None:
        return None

    row = anchor_row + fallback_offset[0]
    col = anchor_col + fallback_offset[1]
    if row < 1 or col < 1:
        return None
    if allow_non_numeric:
        return row, col
    value = get_sheet_float(sheet, (row, col))
    if value is None:
        return row, col
    return row, col


def infer_penetration_series(
    snapshot: SheetSnapshot, anchor_row: int, anchor_col: int
) -> Tuple[Optional[int], List[Tuple[int, float]]]:
    best_col: Optional[int] = None
    best_rows: List[Tuple[int, float]] = []
    best_score: Optional[Tuple[int, int, int]] = None

    col_start = max(snapshot.start_col, anchor_col - 20)
    col_end = min(snapshot.end_col, anchor_col + 2)

    for col in range(col_start, col_end + 1):
        numeric_rows: List[Tuple[int, float]] = []
        pct_like = 0
        for row in range(snapshot.start_row, anchor_row):
            num = to_float(snapshot_value(snapshot, row, col))
            if num is None:
                continue
            numeric_rows.append((row, num))
            if 0 <= abs(num) <= 150:
                pct_like += 1
        if len(numeric_rows) < 3:
            continue

        # Prefer columns with many rows and percentage-like values,
        # with a slight preference for columns close to the anchor.
        proximity_bonus = -abs(col - anchor_col)
        score = (pct_like, len(numeric_rows), proximity_bonus)
        if best_score is None or score > best_score:
            best_score = score
            best_col = col
            best_rows = numeric_rows

    return best_col, best_rows


def infer_last_quarter_label(
    snapshot: SheetSnapshot, start_row: int, candidate_cols: Sequence[int]
) -> Any:
    for col in candidate_cols:
        value = snapshot_value(snapshot, start_row, col)
        if isinstance(value, str) and value.strip():
            return value.strip()
    return start_row


def set_formula2(cell: xw.main.Range, formula: str) -> None:
    try:
        cell.formula2 = formula
    except Exception:
        # fallback for environments where formula2 is unavailable
        cell.formula = formula


def safe_close_workbook(wb: xw.main.Book) -> None:
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
        wb.close()


def parse_file_label(file_path: Path) -> Optional[Dict[str, str]]:
    stem = file_path.stem
    parts = [part.strip() for part in stem.split(" - ") if part.strip()]
    ticker = parts[1].upper() if len(parts) >= 2 else ""

    period_match = re.search(
        r"(Early|Mid|Late)\s*([A-Za-z]{3,9})\s*([12]\d{3})",
        stem,
        flags=re.IGNORECASE,
    )
    if not period_match:
        return None

    period_word = period_match.group(1).title()
    month_token = period_match.group(2).strip().lower()
    year = int(period_match.group(3))
    month = MONTHS.get(month_token[:4], MONTHS.get(month_token[:3]))
    if month is None:
        return None

    day = PERIOD_DAY[period_word.lower()]
    model_dt = date(year, month, day)
    month_abbrev = datetime(year, month, 1).strftime("%b")
    model_period = f"{period_word}{month_abbrev}_{year}"

    if not ticker:
        ticker = "UNKNOWN"
    model = f"{ticker}_{model_period}"

    return {
        "model": model,
        "ticker": ticker,
        "model_period": model_period,
        "model_date": model_dt.isoformat(),
    }


def calc_range_width(max_value: Optional[float], min_value: Optional[float]) -> Optional[float]:
    if max_value is None or min_value is None:
        return None
    return max_value - min_value


def rounded_signature(values: Iterable[Optional[float]]) -> Tuple[Optional[float], ...]:
    signature: List[Optional[float]] = []
    for value in values:
        if value is None:
            signature.append(None)
        else:
            signature.append(round(value, 10))
    return tuple(signature)


def process_empirical_sheet(
    wb: xw.main.Book,
    sheet: xw.main.Sheet,
    meta: Dict[str, str],
    source_file: str,
) -> List[Dict[str, Any]]:
    snapshot = make_sheet_snapshot(sheet)
    anchor = find_anchor_max(snapshot)
    if anchor is None:
        log(f"  - {EMPIRICAL_SHEET_NAME}: skipped (could not find 'max' anchor)")
        return []

    anchor_row, anchor_col = anchor
    min_col = find_min_value_col(snapshot, anchor_row, anchor_col)

    # Anchor-based value cells near the max/min output section.
    forecast_max_cell = (anchor_row + 1, anchor_col)
    forecast_min_cell = (anchor_row + 1, min_col)
    forecast_value_cell = resolve_metric_cell(
        snapshot,
        sheet,
        anchor_row,
        anchor_col,
        keywords=["estimated total sold", "tot fcst", "forecast"],
        fallback_offset=(1, -1),
    )
    reported_sales_cell = resolve_metric_cell(
        snapshot,
        sheet,
        anchor_row,
        anchor_col,
        keywords=["reported sales", "actual sales"],
        fallback_offset=(1, -2),
    )
    quarterly_sales_cell = resolve_metric_cell(
        snapshot,
        sheet,
        anchor_row,
        anchor_col,
        keywords=["quarterly sales"],
        fallback_offset=(1, -3),
    )
    growth_rate_cell = resolve_metric_cell(
        snapshot,
        sheet,
        anchor_row,
        anchor_col,
        keywords=["growth rate"],
        fallback_offset=(1, -4),
    )
    sales_captured_cell = resolve_metric_cell(
        snapshot,
        sheet,
        anchor_row,
        anchor_col,
        keywords=["sales captured", "captured in db"],
        fallback_offset=(1, -5),
    )
    num_quarters_input_cell = resolve_metric_cell(
        snapshot,
        sheet,
        anchor_row,
        anchor_col,
        keywords=["num quarters", "n quarters"],
        fallback_offset=(-1, -6),
        allow_non_numeric=True,
    )

    penetration_col, penetration_rows = infer_penetration_series(snapshot, anchor_row, anchor_col)
    if penetration_col is None or len(penetration_rows) < 1:
        log(f"  - {EMPIRICAL_SHEET_NAME}: skipped (could not infer penetration series)")
        return []

    max_quarters = min(N_QUARTERS, len(penetration_rows))
    calc_col = snapshot.end_col + 3
    calc_row = snapshot.start_row + 1
    avg_pen_cell = (calc_row, calc_col)

    extracted_rows: List[Dict[str, Any]] = []

    for num_quarters_used in range(1, max_quarters + 1):
        if num_quarters_input_cell is not None:
            sheet.range(num_quarters_input_cell).value = num_quarters_used

        start_row = penetration_rows[-num_quarters_used][0]
        end_row = penetration_rows[-1][0]
        avg_formula = (
            f"=AVERAGE(R{start_row}C{penetration_col}:R{end_row}C{penetration_col})"
        )
        set_formula2(sheet.range(avg_pen_cell), avg_formula)
        wb.app.calculate()

        avg_pen_raw = to_float(sheet.range(avg_pen_cell).value)
        avg_penetration_pct = to_math_pct(avg_pen_raw)

        quarterly_sales = get_sheet_float(sheet, quarterly_sales_cell)
        reported_sales = get_sheet_float(sheet, reported_sales_cell)
        growth_rate_pct = get_sheet_float(sheet, growth_rate_cell)
        sales_captured_in_db_pct = get_sheet_float(sheet, sales_captured_cell)
        sales_captured_in_db_pct = to_math_pct(sales_captured_in_db_pct)

        forecast_value = get_sheet_float(sheet, forecast_value_cell)
        if forecast_value is None and quarterly_sales is not None:
            if avg_penetration_pct not in (None, 0):
                forecast_value = quarterly_sales / avg_penetration_pct

        forecast_max = get_sheet_float(sheet, forecast_max_cell)
        forecast_min = get_sheet_float(sheet, forecast_min_cell)
        range_width = calc_range_width(forecast_max, forecast_min)

        last_quarter_used = infer_last_quarter_label(
            snapshot,
            start_row,
            candidate_cols=[penetration_col - 1, penetration_col - 2, penetration_col],
        )

        row: Dict[str, Any] = {
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
            "source_file": source_file,
        }
        extracted_rows.append(row)

    # Temporary formula writes are intentionally not saved to source files.
    sheet.range(avg_pen_cell).value = None
    return extracted_rows


def process_regression_sheet(
    wb: xw.main.Book,
    sheet: xw.main.Sheet,
    meta: Dict[str, str],
    source_file: str,
) -> List[Dict[str, Any]]:
    snapshot = make_sheet_snapshot(sheet)
    anchor = find_anchor_max(snapshot)
    if anchor is None:
        log(f"  - {REGRESSION_SHEET_NAME}: skipped (could not find 'max' anchor)")
        return []

    anchor_row, anchor_col = anchor
    min_col = find_min_value_col(snapshot, anchor_row, anchor_col)

    y_col = anchor_col - 7
    x_col = anchor_col - 11

    xy_rows: List[Tuple[int, float, float]] = []
    for row in range(snapshot.start_row, anchor_row):
        x_val = to_float(snapshot_value(snapshot, row, x_col))
        y_val = to_float(snapshot_value(snapshot, row, y_col))
        if x_val is None or y_val is None:
            continue
        xy_rows.append((row, x_val, y_val))

    if len(xy_rows) < 2:
        log(f"  - {REGRESSION_SHEET_NAME}: skipped (not enough X/Y history)")
        return []

    forecast_total_cell = resolve_metric_cell(
        snapshot,
        sheet,
        anchor_row,
        anchor_col,
        keywords=["tot fcst w/o sa", "tot fcst without sa", "forecast"],
        fallback_offset=(1, -1),
    )
    actual_value_cell = resolve_metric_cell(
        snapshot,
        sheet,
        anchor_row,
        anchor_col,
        keywords=["actual", "reported sales"],
        fallback_offset=None,
    )
    forecast_max_cell = (anchor_row + 1, anchor_col)
    forecast_min_cell = (anchor_row + 1, min_col)

    calc_col = snapshot.end_col + 3
    intercept_cell = (snapshot.start_row + 1, calc_col)
    slope_cell = (snapshot.start_row + 2, calc_col)

    max_quarters = min(N_QUARTERS, len(xy_rows))
    extracted_rows: List[Dict[str, Any]] = []
    prev_signature: Optional[Tuple[Optional[float], ...]] = None

    for num_quarters_used in range(2, max_quarters + 1):
        series_slice = xy_rows[-num_quarters_used:]
        start_row = series_slice[0][0]
        end_row = series_slice[-1][0]

        intercept_formula = (
            f"=INTERCEPT(R{start_row}C{y_col}:R{end_row}C{y_col},"
            f"R{start_row}C{x_col}:R{end_row}C{x_col})"
        )
        slope_formula = (
            f"=SLOPE(R{start_row}C{y_col}:R{end_row}C{y_col},"
            f"R{start_row}C{x_col}:R{end_row}C{x_col})"
        )
        set_formula2(sheet.range(intercept_cell), intercept_formula)
        set_formula2(sheet.range(slope_cell), slope_formula)
        wb.app.calculate()

        intercept = get_sheet_float(sheet, intercept_cell)
        slope = get_sheet_float(sheet, slope_cell)

        forecast_value = get_sheet_float(sheet, forecast_total_cell)
        if forecast_value is None and intercept is not None and slope is not None:
            # Fallback if workbook does not expose TOT FCST w/o SA cell directly.
            x_next = series_slice[-1][1] + 1.0
            forecast_value = intercept + (slope * x_next)

        forecast_max = get_sheet_float(sheet, forecast_max_cell)
        forecast_min = get_sheet_float(sheet, forecast_min_cell)
        range_width = calc_range_width(forecast_max, forecast_min)
        actual_value = get_sheet_float(sheet, actual_value_cell)

        signature = rounded_signature(
            [forecast_value, forecast_max, forecast_min, intercept, slope]
        )
        if signature == prev_signature:
            continue
        prev_signature = signature

        row: Dict[str, Any] = {
            "model": meta["model"],
            "ticker": meta["ticker"],
            "model_period": meta["model_period"],
            "model_date": meta["model_date"],
            "method": "regression",
            "parameter_name": "num_quarters_used",
            "parameter_value": num_quarters_used,
            "num_quarters_used": num_quarters_used,
            "forecast_value": forecast_value,
            "actual_value": actual_value,
            "forecast_max": forecast_max,
            "forecast_min": forecast_min,
            "range_width": range_width,
            "intercept": intercept,
            "slope": slope,
            "source_file": source_file,
        }
        extracted_rows.append(row)

    # Temporary formula writes are intentionally not saved to source files.
    sheet.range(intercept_cell).value = None
    sheet.range(slope_cell).value = None
    return extracted_rows


def unique_output_path(in_dir: Path, out_dir: Path) -> Path:
    base_name = f"{in_dir.name}_PARAM"
    candidate = out_dir / f"{base_name}.xlsx"
    idx = 1
    while candidate.exists():
        candidate = out_dir / f"{base_name}.{idx}.xlsx"
        idx += 1
    return candidate


def write_sheet(
    workbook: Workbook, sheet_name: str, columns: List[str], rows: List[Dict[str, Any]]
) -> None:
    ws = workbook.create_sheet(title=sheet_name)
    ws.append(columns)
    for row in rows:
        ws.append([row.get(col) for col in columns])

    for cell in ws[1]:
        cell.font = Font(bold=True)

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(columns))}{max(1, ws.max_row)}"

    for col_idx, column_name in enumerate(columns, start=1):
        max_len = len(column_name)
        for row_idx in range(2, ws.max_row + 1):
            value = ws.cell(row=row_idx, column=col_idx).value
            if value is None:
                continue
            max_len = max(max_len, len(str(value)))
        ws.column_dimensions[get_column_letter(col_idx)].width = min(max_len + 2, 45)


def write_output_workbook(
    out_path: Path,
    empirical_rows: List[Dict[str, Any]],
    regression_rows: List[Dict[str, Any]],
) -> None:
    wb = Workbook()
    default_sheet = wb.active
    wb.remove(default_sheet)
    write_sheet(wb, "empirical_candidates", EMPIRICAL_COLUMNS, empirical_rows)
    write_sheet(wb, "regression_candidates", REGRESSION_COLUMNS, regression_rows)
    wb.save(out_path)


def list_source_files(in_dir: Path) -> List[Path]:
    if not in_dir.exists():
        log(f"Input directory does not exist: {in_dir}")
        return []
    if not in_dir.is_dir():
        log(f"Input path is not a directory: {in_dir}")
        return []

    files: List[Path] = []
    for path in sorted(in_dir.iterdir()):
        if not path.is_file():
            continue
        if path.name.startswith("~"):
            log(f"Skipped {path.name}: temp file")
            continue
        if path.suffix.lower() != ".xlsx":
            log(f"Skipped {path.name}: not .xlsx")
            continue
        files.append(path)
    return files


def process_workbooks() -> None:
    output_dir.mkdir(parents=True, exist_ok=True)

    source_files = list_source_files(input_dir)
    if not source_files:
        log("No .xlsx files to process.")
        return

    empirical_rows: List[Dict[str, Any]] = []
    regression_rows: List[Dict[str, Any]] = []
    files_processed = 0

    app = xw.App(visible=False, add_book=False)
    app.display_alerts = False
    app.screen_updating = False

    try:
        for file_path in source_files:
            meta = parse_file_label(file_path)
            if meta is None:
                log(f"Skipped {file_path.name}: filename does not match expected label pattern")
                continue

            log(f"Processing {file_path.name}")
            try:
                src_wb = app.books.open(str(file_path), update_links=False)
            except Exception as exc:
                log(f"Skipped {file_path.name}: open failed ({exc})")
                continue

            file_empirical_count = 0
            file_regression_count = 0
            try:
                empirical_sheet = None
                regression_sheet = None
                try:
                    empirical_sheet = src_wb.sheets[EMPIRICAL_SHEET_NAME]
                except Exception:
                    log(f"  - {EMPIRICAL_SHEET_NAME}: skipped (sheet missing)")
                try:
                    regression_sheet = src_wb.sheets[REGRESSION_SHEET_NAME]
                except Exception:
                    log(f"  - {REGRESSION_SHEET_NAME}: skipped (sheet missing)")

                if empirical_sheet is not None:
                    extracted_empirical = process_empirical_sheet(
                        src_wb, empirical_sheet, meta, file_path.name
                    )
                    empirical_rows.extend(extracted_empirical)
                    file_empirical_count = len(extracted_empirical)

                if regression_sheet is not None:
                    extracted_regression = process_regression_sheet(
                        src_wb, regression_sheet, meta, file_path.name
                    )
                    regression_rows.extend(extracted_regression)
                    file_regression_count = len(extracted_regression)
            except Exception as exc:
                log(f"  - processing warning for {file_path.name}: {exc}")
            finally:
                safe_close_workbook(src_wb)

            files_processed += 1
            log(
                "  - rows extracted "
                f"(empirical={file_empirical_count}, regression={file_regression_count})"
            )
    finally:
        app.quit()

    output_path = unique_output_path(input_dir, output_dir)
    write_output_workbook(output_path, empirical_rows, regression_rows)

    log(f"Output path: {output_path}")
    log(f"Number of files processed: {files_processed}")
    log(f"Number of empirical rows: {len(empirical_rows)}")
    log(f"Number of regression rows: {len(regression_rows)}")


if __name__ == "__main__":
    process_workbooks()
