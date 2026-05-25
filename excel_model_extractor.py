from __future__ import annotations

from dataclasses import dataclass
from datetime import date
from pathlib import Path
import re
from typing import Any, Dict, Iterable, List, Optional, Sequence, Tuple

import xlwings as xw
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter


# =========================
# User-configurable paths
# =========================
input_dir = Path("input")
output_dir = Path("output")


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

PERIOD_PATTERN = re.compile(
    r"(?P<phase>Early|Mid|Late)\s*[_-]?\s*"
    r"(?P<month>Jan(?:uary)?|Feb(?:ruary)?|Mar(?:ch)?|Apr(?:il)?|May|"
    r"Jun(?:e)?|Jul(?:y)?|Aug(?:ust)?|Sep(?:t(?:ember)?)?|Oct(?:ober)?|"
    r"Nov(?:ember)?|Dec(?:ember)?)\s*[_-]?\s*"
    r"(?P<year>20\d{2})",
    flags=re.IGNORECASE,
)

PHASE_DAY = {"Early": 5, "Mid": 15, "Late": 25}
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


@dataclass(frozen=True)
class FileLabel:
    model: str
    ticker: str
    model_period: str
    model_date: str


def parse_file_label(file_name: str) -> FileLabel:
    """Parse ticker/period/date/model from file name."""
    stem = Path(file_name).stem

    # Ticker from common "AAA - TICKER - Period..." naming.
    parts = [p.strip() for p in stem.split(" - ") if p.strip()]
    ticker = ""
    if len(parts) >= 2:
        ticker = re.split(r"[\s_]+", parts[1])[0].upper()
    if not ticker:
        ticker_match = re.search(r"\b[A-Z]{2,8}\b", stem)
        ticker = ticker_match.group(0).upper() if ticker_match else "UNKNOWN"

    period_match = PERIOD_PATTERN.search(stem)
    if not period_match:
        raise ValueError(
            "could not parse model period from filename (need Early/Mid/Late + month + year)"
        )

    phase = period_match.group("phase").title()
    month_token = period_match.group("month")[:3].lower()
    month_abbrev = month_token.title()
    year = int(period_match.group("year"))

    day = PHASE_DAY[phase]
    month = MONTH_NUM[month_token]
    model_period = f"{phase}{month_abbrev}_{year}"
    model_date = date(year, month, day).isoformat()
    model = f"{ticker}_{model_period}"
    return FileLabel(model=model, ticker=ticker, model_period=model_period, model_date=model_date)


def normalize_text(value: Any) -> str:
    if value is None:
        return ""
    return re.sub(r"[^a-z0-9]+", "_", str(value).strip().lower()).strip("_")


def to_float(value: Any) -> Optional[float]:
    if value is None or value == "":
        return None
    if isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        return float(value)
    text = str(value).strip().replace(",", "")
    if not text:
        return None
    pct = False
    if text.endswith("%"):
        pct = True
        text = text[:-1].strip()
    try:
        parsed = float(text)
    except ValueError:
        return None
    return parsed * (0.01 if pct else 1.0)


def ensure_2d(values: Any) -> List[List[Any]]:
    if values is None:
        return []
    if isinstance(values, list):
        if not values:
            return []
        if isinstance(values[0], list):
            return values
        return [values]
    return [[values]]


def ensure_list(values: Any) -> List[Any]:
    if values is None:
        return []
    if isinstance(values, list):
        return values
    return [values]


def build_output_path(in_dir: Path, out_dir: Path) -> Path:
    base_name = f"{in_dir.name}_PARAM.xlsx"
    candidate = out_dir / base_name
    suffix = 1
    while candidate.exists():
        candidate = out_dir / f"{in_dir.name}_PARAM.{suffix}.xlsx"
        suffix += 1
    return candidate


def get_sheet(wb: xw.Book, sheet_name: str) -> Optional[xw.Sheet]:
    lowered = sheet_name.strip().lower()
    for sheet in wb.sheets:
        if sheet.name.strip().lower() == lowered:
            return sheet
    return None


def read_sheet_grid(sheet: xw.Sheet) -> Tuple[List[List[Any]], int, int, int, int]:
    used = sheet.used_range
    grid = ensure_2d(used.value)
    start_row = used.row
    start_col = used.column
    last_row = used.last_cell.row
    last_col = used.last_cell.column
    return grid, start_row, start_col, last_row, last_col


def find_anchor_max(grid: Sequence[Sequence[Any]], start_row: int, start_col: int) -> Optional[Tuple[int, int]]:
    for r_idx, row_values in enumerate(grid):
        for c_idx, cell_value in enumerate(row_values):
            if isinstance(cell_value, str) and normalize_text(cell_value) == "max":
                return start_row + r_idx, start_col + c_idx
    return None


def collect_history(
    sheet: xw.Sheet, anchor_row: int, x_col: int, y_col: int
) -> List[Dict[str, Any]]:
    """
    Collect contiguous numeric history ending immediately above anchor_row.
    Uses x_col and y_col from anchor-based offsets.
    """
    if anchor_row <= 1 or x_col <= 0 or y_col <= 0:
        return []

    x_values = ensure_list(sheet.range((1, x_col), (anchor_row - 1, x_col)).value)
    y_values = ensure_list(sheet.range((1, y_col), (anchor_row - 1, y_col)).value)
    q_col = x_col - 1
    q_values: List[Any] = []
    if q_col > 0:
        q_values = ensure_list(sheet.range((1, q_col), (anchor_row - 1, q_col)).value)

    total_rows = min(len(x_values), len(y_values))
    history: List[Dict[str, Any]] = []
    started = False

    for idx in range(total_rows - 1, -1, -1):
        x_val = to_float(x_values[idx])
        y_val = to_float(y_values[idx])
        if x_val is None or y_val is None:
            if started:
                break
            continue
        started = True
        history.append(
            {
                "row": idx + 1,
                "x": x_val,
                "y": y_val,
                "quarter": q_values[idx] if idx < len(q_values) else None,
            }
        )

    history.reverse()
    return history


def choose_helper_area(anchor_row: int, anchor_col: int, last_col: int, cols_needed: int) -> Tuple[int, int]:
    helper_row = max(anchor_row + 2, 2)
    max_start_col = 16384 - cols_needed - 1
    helper_col = max(last_col + 2, anchor_col + 2)
    helper_col = min(helper_col, max_start_col)
    helper_col = max(helper_col, 1)
    return helper_row, helper_col


def safe_close_source_workbook(wb: xw.Book) -> None:
    try:
        wb.close(save=False)
        return
    except TypeError:
        pass
    except Exception as exc:
        print(f"  close(save=False) failed: {exc}")

    # Fallback path for environments where save=False signature is not accepted.
    for closer in (lambda: wb.api.Close(False), lambda: wb.close(), lambda: wb.api.Close()):
        try:
            closer()
            return
        except Exception:
            continue
    print("  warning: unable to close workbook with safe fallback")


def signature(values: Iterable[Optional[float]]) -> Tuple[Optional[float], ...]:
    out: List[Optional[float]] = []
    for item in values:
        out.append(None if item is None else round(item, 10))
    return tuple(out)


def process_empirical_sheet(wb: xw.Book, info: FileLabel, source_file: str) -> List[Dict[str, Any]]:
    sheet = get_sheet(wb, "Empirical Model")
    if sheet is None:
        print("  skipped empirical sheet: 'Empirical Model' not found")
        return []

    grid, start_row, start_col, _last_row, last_col = read_sheet_grid(sheet)
    anchor = find_anchor_max(grid, start_row, start_col)
    if anchor is None:
        print("  skipped empirical sheet: max anchor not found")
        return []

    anchor_row, anchor_col = anchor
    x_col = anchor_col - 11
    y_col = anchor_col - 7
    history = collect_history(sheet, anchor_row=anchor_row, x_col=x_col, y_col=y_col)
    if len(history) < 2:
        print("  skipped empirical sheet: not enough numeric history")
        return []

    current_x = to_float(sheet.range((anchor_row, x_col)).value) or history[-1]["x"]
    current_y = to_float(sheet.range((anchor_row, y_col)).value) or history[-1]["y"]

    max_n = min(N_QUARTERS, len(history))
    helper_row, helper_col = choose_helper_area(anchor_row, anchor_col, last_col, cols_needed=3)

    # Use R1C1 formula2 to compute average/min/max penetration for each n.
    for n in range(1, max_n + 1):
        subset = history[-n:]
        start = subset[0]["row"]
        end = subset[-1]["row"]
        row = helper_row + (n - 1)

        avg_cell = sheet.range((row, helper_col))
        min_cell = sheet.range((row, helper_col + 1))
        max_cell = sheet.range((row, helper_col + 2))

        avg_cell.formula2 = (
            f"=AVERAGE(R{start}C{x_col}:R{end}C{x_col}/R{start}C{y_col}:R{end}C{y_col})"
        )
        min_cell.formula2 = (
            f"=MIN(R{start}C{x_col}:R{end}C{x_col}/R{start}C{y_col}:R{end}C{y_col})"
        )
        max_cell.formula2 = (
            f"=MAX(R{start}C{x_col}:R{end}C{x_col}/R{start}C{y_col}:R{end}C{y_col})"
        )

    wb.app.calculate()

    helper_values = ensure_2d(
        sheet.range((helper_row, helper_col), (helper_row + max_n - 1, helper_col + 2)).value
    )

    rows: List[Dict[str, Any]] = []
    for idx in range(max_n):
        n_used = idx + 1
        subset = history[-n_used:]
        avg_pen_ratio = to_float(helper_values[idx][0]) if idx < len(helper_values) else None
        min_pen_ratio = to_float(helper_values[idx][1]) if idx < len(helper_values) else None
        max_pen_ratio = to_float(helper_values[idx][2]) if idx < len(helper_values) else None

        if avg_pen_ratio is None or avg_pen_ratio == 0:
            continue

        forecast_value = current_x / avg_pen_ratio if current_x is not None else None
        forecast_max = (
            current_x / min_pen_ratio
            if current_x is not None and min_pen_ratio not in (None, 0)
            else None
        )
        forecast_min = (
            current_x / max_pen_ratio
            if current_x is not None and max_pen_ratio not in (None, 0)
            else None
        )
        range_width = (
            (forecast_max - forecast_min)
            if forecast_max is not None and forecast_min is not None
            else None
        )

        previous_x = subset[-2]["x"] if len(subset) >= 2 else None
        growth_rate_pct = (
            ((current_x / previous_x) - 1.0) * 100.0
            if current_x is not None and previous_x not in (None, 0)
            else None
        )
        sales_captured_in_db_pct = (
            (current_x / current_y) * 100.0
            if current_x is not None and current_y not in (None, 0)
            else None
        )

        quarter_value = subset[-1].get("quarter")
        last_quarter_used = str(quarter_value).strip() if quarter_value not in (None, "") else ""
        avg_penetration_pct = avg_pen_ratio * 100.0

        rows.append(
            {
                "model": info.model,
                "ticker": info.ticker,
                "model_period": info.model_period,
                "model_date": info.model_date,
                "method": "empirical",
                "parameter_name": "avg_penetration_pct",
                "parameter_value": avg_penetration_pct,
                "num_quarters_used": n_used,
                "last_quarter_used": last_quarter_used,
                "forecast_value": forecast_value,
                "actual_value": current_y,
                "forecast_max": forecast_max,
                "forecast_min": forecast_min,
                "range_width": range_width,
                "avg_penetration_pct": avg_penetration_pct,
                "quarterly_sales": current_x,
                "reported_sales": current_y,
                "growth_rate_pct": growth_rate_pct,
                "sales_captured_in_db_pct": sales_captured_in_db_pct,
                "source_file": source_file,
            }
        )

    return rows


def process_regression_sheet(wb: xw.Book, info: FileLabel, source_file: str) -> List[Dict[str, Any]]:
    sheet = get_sheet(wb, "Regression Model")
    if sheet is None:
        print("  skipped regression sheet: 'Regression Model' not found")
        return []

    grid, start_row, start_col, _last_row, last_col = read_sheet_grid(sheet)
    anchor = find_anchor_max(grid, start_row, start_col)
    if anchor is None:
        print("  skipped regression sheet: max anchor not found")
        return []

    anchor_row, anchor_col = anchor
    y_col = anchor_col - 7
    x_col = anchor_col - 11
    history = collect_history(sheet, anchor_row=anchor_row, x_col=x_col, y_col=y_col)
    if len(history) < 2:
        print("  skipped regression sheet: not enough numeric history")
        return []

    current_x = to_float(sheet.range((anchor_row, x_col)).value) or history[-1]["x"]
    current_y = to_float(sheet.range((anchor_row, y_col)).value)

    max_n = min(N_QUARTERS, len(history))
    helper_row, helper_col = choose_helper_area(anchor_row, anchor_col, last_col, cols_needed=6)

    # Use R1C1 formula2 for INTERCEPT/SLOPE and supporting forecast range values.
    for n in range(1, max_n + 1):
        subset = history[-n:]
        start = subset[0]["row"]
        end = subset[-1]["row"]
        row = helper_row + (n - 1)

        intercept_cell = sheet.range((row, helper_col))
        slope_cell = sheet.range((row, helper_col + 1))
        forecast_cell = sheet.range((row, helper_col + 2))
        steyx_cell = sheet.range((row, helper_col + 3))
        max_cell = sheet.range((row, helper_col + 4))
        min_cell = sheet.range((row, helper_col + 5))

        intercept_cell.formula2 = (
            f"=INTERCEPT(R{start}C{y_col}:R{end}C{y_col},R{start}C{x_col}:R{end}C{x_col})"
        )
        slope_cell.formula2 = (
            f"=SLOPE(R{start}C{y_col}:R{end}C{y_col},R{start}C{x_col}:R{end}C{x_col})"
        )
        forecast_cell.formula2 = f"={current_x}*R{row}C{helper_col + 1}+R{row}C{helper_col}"
        steyx_cell.formula2 = f"=STEYX(R{start}C{y_col}:R{end}C{y_col},R{start}C{x_col}:R{end}C{x_col})"
        max_cell.formula2 = f"=R{row}C{helper_col + 2}+R{row}C{helper_col + 3}"
        min_cell.formula2 = f"=R{row}C{helper_col + 2}-R{row}C{helper_col + 3}"

    wb.app.calculate()

    helper_values = ensure_2d(
        sheet.range((helper_row, helper_col), (helper_row + max_n - 1, helper_col + 5)).value
    )

    rows: List[Dict[str, Any]] = []
    previous_signature: Optional[Tuple[Optional[float], ...]] = None

    for idx in range(max_n):
        n_used = idx + 1
        if idx >= len(helper_values):
            continue
        values = helper_values[idx]

        intercept = to_float(values[0]) if len(values) > 0 else None
        slope = to_float(values[1]) if len(values) > 1 else None
        forecast_value = to_float(values[2]) if len(values) > 2 else None
        forecast_max = to_float(values[4]) if len(values) > 4 else None
        forecast_min = to_float(values[5]) if len(values) > 5 else None
        range_width = (
            (forecast_max - forecast_min)
            if forecast_max is not None and forecast_min is not None
            else None
        )

        row_signature = signature((intercept, slope, forecast_value, forecast_max, forecast_min))
        if previous_signature is not None and row_signature == previous_signature:
            # Prevent duplicate final row.
            continue
        previous_signature = row_signature

        rows.append(
            {
                "model": info.model,
                "ticker": info.ticker,
                "model_period": info.model_period,
                "model_date": info.model_date,
                "method": "regression",
                "parameter_name": "num_quarters_used",
                "parameter_value": n_used,
                "num_quarters_used": n_used,
                "forecast_value": forecast_value,
                "actual_value": current_y if current_y is not None else "",
                "forecast_max": forecast_max,
                "forecast_min": forecast_min,
                "range_width": range_width,
                "intercept": intercept,
                "slope": slope,
                "source_file": source_file,
            }
        )

    return rows


def write_sheet(ws: Any, columns: Sequence[str], rows: Sequence[Dict[str, Any]]) -> None:
    ws.append(list(columns))
    for row in rows:
        ws.append([row.get(column, "") for column in columns])

    bold = Font(bold=True)
    for cell in ws[1]:
        cell.font = bold

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions

    for col_idx, column_name in enumerate(columns, start=1):
        max_len = len(column_name)
        for row_idx in range(2, ws.max_row + 1):
            value = ws.cell(row=row_idx, column=col_idx).value
            if value is None:
                continue
            text = str(value)
            if len(text) > max_len:
                max_len = len(text)
        ws.column_dimensions[get_column_letter(col_idx)].width = min(max(max_len + 2, 12), 45)


def write_output_workbook(
    output_path: Path,
    empirical_rows: Sequence[Dict[str, Any]],
    regression_rows: Sequence[Dict[str, Any]],
) -> None:
    wb = Workbook()
    ws_empirical = wb.active
    ws_empirical.title = "empirical_candidates"
    write_sheet(ws_empirical, EMPIRICAL_COLUMNS, empirical_rows)

    ws_regression = wb.create_sheet("regression_candidates")
    write_sheet(ws_regression, REGRESSION_COLUMNS, regression_rows)
    wb.save(output_path)


def iter_valid_input_files(in_dir: Path) -> List[Path]:
    files: List[Path] = []
    for path in sorted(in_dir.iterdir(), key=lambda p: p.name.lower()):
        if not path.is_file():
            print(f"SKIPPED {path.name}: not a file")
            continue
        if path.name.startswith("~"):
            print(f"SKIPPED {path.name}: temp file")
            continue
        if path.suffix.lower() != ".xlsx":
            print(f"SKIPPED {path.name}: not .xlsx")
            continue
        files.append(path)
    return files


def main() -> None:
    in_dir = input_dir.expanduser().resolve()
    out_dir = output_dir.expanduser().resolve()
    out_dir.mkdir(parents=True, exist_ok=True)

    if not in_dir.exists():
        raise FileNotFoundError(f"Input directory does not exist: {in_dir}")

    files = iter_valid_input_files(in_dir)
    empirical_rows: List[Dict[str, Any]] = []
    regression_rows: List[Dict[str, Any]] = []
    processed_files = 0

    app: Optional[xw.App] = None
    try:
        app = xw.App(visible=False, add_book=False)
        app.display_alerts = False
        app.screen_updating = False

        for file_path in files:
            print(f"PROCESSING {file_path.name}")

            try:
                file_info = parse_file_label(file_path.name)
            except Exception as exc:
                print(f"SKIPPED {file_path.name}: filename parsing failed ({exc})")
                continue

            source_wb: Optional[xw.Book] = None
            try:
                source_wb = app.books.open(str(file_path), update_links=False)
                empirical = process_empirical_sheet(source_wb, file_info, file_path.name)
                regression = process_regression_sheet(source_wb, file_info, file_path.name)
                empirical_rows.extend(empirical)
                regression_rows.extend(regression)
                processed_files += 1
            except Exception as exc:
                print(f"SKIPPED {file_path.name}: workbook processing error ({exc})")
            finally:
                if source_wb is not None:
                    safe_close_source_workbook(source_wb)
    finally:
        if app is not None:
            app.quit()

    output_path = build_output_path(in_dir, out_dir)
    write_output_workbook(output_path, empirical_rows, regression_rows)

    print(f"OUTPUT {output_path}")
    print(f"FILES_PROCESSED {processed_files}")
    print(f"EMPIRICAL_ROWS {len(empirical_rows)}")
    print(f"REGRESSION_ROWS {len(regression_rows)}")


if __name__ == "__main__":
    main()
