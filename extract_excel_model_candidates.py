#!/usr/bin/env python3
"""Extract empirical and regression parameter candidates from Excel model files."""

from __future__ import annotations

import re
from datetime import date, datetime
from pathlib import Path
from statistics import pstdev
from typing import Any, Dict, List, Optional, Sequence, Tuple

import xlwings as xw
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

# Configure these two paths before running.
input_dir = Path("input")
output_dir = Path("output")

MAX_QUARTERS = 10

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

PERIOD_PATTERN = re.compile(r"(Early|Mid|Late)\s*([A-Za-z]{3,9})\s*([12][0-9]{3})", re.IGNORECASE)
DAY_BY_PERIOD = {"early": 5, "mid": 15, "late": 25}


def normalize_text(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, str):
        return value.strip().lower()
    return str(value).strip().lower()


def safe_float(value: Any) -> Optional[float]:
    if value is None:
        return None
    if isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        return float(value)
    if isinstance(value, str):
        cleaned = value.strip().replace(",", "")
        if not cleaned:
            return None
        is_percent = cleaned.endswith("%")
        if is_percent:
            cleaned = cleaned[:-1]
        try:
            numeric = float(cleaned)
            if is_percent:
                numeric /= 100.0
            return numeric
        except ValueError:
            return None
    return None


def safe_divide(numerator: Optional[float], denominator: Optional[float]) -> Optional[float]:
    if numerator is None or denominator in (None, 0):
        return None
    return numerator / denominator


def to_matrix(values: Any) -> List[List[Any]]:
    if values is None:
        return []
    if isinstance(values, tuple):
        values = list(values)
    if isinstance(values, list):
        if not values:
            return []
        first = values[0]
        if isinstance(first, tuple):
            values = [list(row) for row in values]
            first = values[0] if values else []
        if isinstance(first, list):
            return values
        return [values]
    return [[values]]


class SheetSnapshot:
    """Cache used-range values so each sheet is scanned once."""

    def __init__(self, sheet: xw.Sheet):
        self.sheet = sheet
        used_range = sheet.used_range
        self.start_row = used_range.row
        self.start_col = used_range.column
        self.values = to_matrix(used_range.value)
        self.text_positions: Dict[str, List[Tuple[int, int]]] = {}

        for row_offset, row_values in enumerate(self.values):
            for col_offset, cell_value in enumerate(row_values):
                key = normalize_text(cell_value)
                if not key:
                    continue
                self.text_positions.setdefault(key, []).append(
                    (self.start_row + row_offset, self.start_col + col_offset)
                )

    def get_value(self, row: int, col: int) -> Any:
        row_idx = row - self.start_row
        col_idx = col - self.start_col
        if row_idx < 0 or col_idx < 0:
            return self.sheet.range((row, col)).value
        if row_idx >= len(self.values):
            return self.sheet.range((row, col)).value
        row_values = self.values[row_idx]
        if col_idx >= len(row_values):
            return self.sheet.range((row, col)).value
        return row_values[col_idx]

    def find_positions(self, label: str) -> List[Tuple[int, int]]:
        return self.text_positions.get(normalize_text(label), [])


def parse_month(month_token: str) -> Optional[int]:
    token = month_token.strip()
    if not token:
        return None
    for fmt in ("%b", "%B"):
        try:
            return datetime.strptime(token[:3] if fmt == "%b" else token, fmt).month
        except ValueError:
            continue
    return None


def parse_file_labels(file_path: Path) -> Dict[str, str]:
    stem = file_path.stem
    parts = [part.strip() for part in stem.split(" - ")]
    ticker = parts[1] if len(parts) > 1 else "UNKNOWN"
    ticker = re.sub(r"[^A-Za-z0-9]+", "", ticker).upper() or "UNKNOWN"

    period_token = parts[2] if len(parts) > 2 else stem
    period_token = period_token.split("_")[0]
    match = PERIOD_PATTERN.search(period_token)

    if match:
        period_word = match.group(1).capitalize()
        month_token = match.group(2)
        year = int(match.group(3))
        month = parse_month(month_token)
        if month is not None:
            month_abbr = datetime(year, month, 1).strftime("%b")
            model_period = f"{period_word}{month_abbr}_{year}"
            day = DAY_BY_PERIOD[period_word.lower()]
            model_date = date(year, month, day).isoformat()
        else:
            model_period = re.sub(r"\s+", "", period_token)
            model_date = ""
    else:
        model_period = re.sub(r"\s+", "", period_token)
        model_date = ""

    model = f"{ticker}_{model_period}" if model_period else ticker
    return {
        "model": model,
        "ticker": ticker,
        "model_period": model_period,
        "model_date": model_date,
    }


def resolve_output_path(input_folder: Path, output_folder: Path) -> Path:
    output_folder.mkdir(parents=True, exist_ok=True)
    base_name = f"{input_folder.name}_PARAM.xlsx"
    candidate = output_folder / base_name
    if not candidate.exists():
        return candidate

    counter = 1
    while True:
        candidate = output_folder / f"{input_folder.name}_PARAM.{counter}.xlsx"
        if not candidate.exists():
            return candidate
        counter += 1


def get_sheet(workbook: xw.Book, sheet_name: str) -> Optional[xw.Sheet]:
    sheet_name_norm = normalize_text(sheet_name)
    for sheet in workbook.sheets:
        if normalize_text(sheet.name) == sheet_name_norm:
            return sheet
    return None


def safe_close_workbook(workbook: xw.Book) -> None:
    try:
        workbook.close(save=False)
        return
    except TypeError:
        pass
    except Exception:
        pass

    try:
        workbook.api.Close(False)
    except Exception as exc:
        print(f"Warning: failed to close workbook safely ({exc})")


def set_formula2(cell: xw.Range, formula: str) -> None:
    try:
        cell.formula2 = formula
    except Exception:
        cell.formula = formula


def choose_anchor_max(snapshot: SheetSnapshot) -> Optional[Tuple[int, int]]:
    max_positions = snapshot.find_positions("max")
    if not max_positions:
        return None
    # In these templates the relevant "max" is usually deepest in the model area.
    return sorted(max_positions, key=lambda pos: (pos[0], pos[1]))[-1]


def read_value_next_to_label(snapshot: SheetSnapshot, position: Tuple[int, int]) -> Optional[float]:
    row, col = position
    neighbor_offsets = [(0, 1), (0, 2), (1, 0), (1, 1), (-1, 1), (0, -1)]
    for row_offset, col_offset in neighbor_offsets:
        value = safe_float(snapshot.get_value(row + row_offset, col + col_offset))
        if value is not None:
            return value
    return None


def read_anchor_min_max(snapshot: SheetSnapshot, anchor: Tuple[int, int]) -> Tuple[Optional[float], Optional[float]]:
    anchor_row, anchor_col = anchor
    max_value = read_value_next_to_label(snapshot, anchor)

    min_positions = snapshot.find_positions("min")
    min_value = None
    if min_positions:
        nearest_min = min(
            min_positions,
            key=lambda pos: abs(pos[0] - anchor_row) + abs(pos[1] - anchor_col),
        )
        min_value = read_value_next_to_label(snapshot, nearest_min)

    return max_value, min_value


def quarter_to_text(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, datetime):
        return value.date().isoformat()
    if isinstance(value, date):
        return value.isoformat()
    text = str(value).strip()
    return text


def collect_recent_history(
    snapshot: SheetSnapshot,
    anchor_row: int,
    quarter_col: int,
    x_col: int,
    y_col: int,
) -> List[Dict[str, Any]]:
    numeric_rows: List[Dict[str, Any]] = []
    for row in range(snapshot.start_row, anchor_row):
        x_value = safe_float(snapshot.get_value(row, x_col))
        y_value = safe_float(snapshot.get_value(row, y_col))
        if x_value is None or y_value is None:
            continue
        quarter_value = snapshot.get_value(row, quarter_col)
        numeric_rows.append(
            {
                "row": row,
                "quarter": quarter_value,
                "x": x_value,
                "y": y_value,
            }
        )

    if not numeric_rows:
        return []

    blocks: List[List[Dict[str, Any]]] = []
    current_block = [numeric_rows[0]]
    for item in numeric_rows[1:]:
        if item["row"] == current_block[-1]["row"] + 1:
            current_block.append(item)
        else:
            blocks.append(current_block)
            current_block = [item]
    blocks.append(current_block)

    # Prefer the block nearest to the anchor; break ties by longer length.
    blocks.sort(key=lambda block: (block[-1]["row"], len(block)), reverse=True)
    return blocks[0]


def build_empirical_rows(
    workbook: xw.Book,
    sheet: xw.Sheet,
    file_meta: Dict[str, str],
    source_file: str,
) -> List[Dict[str, Any]]:
    snapshot = SheetSnapshot(sheet)
    anchor = choose_anchor_max(snapshot)
    if anchor is None:
        print(f"Skipped empirical extraction for {source_file}: 'max' anchor not found")
        return []

    anchor_row, anchor_col = anchor
    x_col = anchor_col - 11
    y_col = anchor_col - 7
    quarter_col = x_col - 1
    history = collect_recent_history(snapshot, anchor_row, quarter_col, x_col, y_col)

    if not history:
        print(f"Skipped empirical extraction for {source_file}: no numeric quarter history found")
        return []

    base_max, base_min = read_anchor_min_max(snapshot, anchor)
    helper_cell = sheet.range((anchor_row + 6, anchor_col + 24))
    rows: List[Dict[str, Any]] = []

    max_quarters = min(MAX_QUARTERS, len(history))
    for n_quarters in range(1, max_quarters + 1):
        window = history[-n_quarters:]
        start_row = window[0]["row"]
        end_row = window[-1]["row"]
        latest = window[-1]
        latest_x = latest["x"]
        latest_y = latest["y"]

        avg_formula = (
            f'=AVERAGE(IFERROR((R{start_row}C{x_col}:R{end_row}C{x_col})/'
            f'(R{start_row}C{y_col}:R{end_row}C{y_col}),""))'
        )
        set_formula2(helper_cell, avg_formula)
        workbook.app.calculate()

        avg_penetration = safe_float(helper_cell.value)
        forecast_value = safe_divide(latest_x, avg_penetration)

        penetration_window: List[float] = []
        for item in window:
            penetration = safe_divide(item["x"], item["y"])
            if penetration is not None and penetration > 0:
                penetration_window.append(penetration)

        forecast_max = base_max
        forecast_min = base_min
        if penetration_window:
            if forecast_max is None:
                forecast_max = safe_divide(latest_x, min(penetration_window))
            if forecast_min is None:
                forecast_min = safe_divide(latest_x, max(penetration_window))

        previous_y = window[-2]["y"] if len(window) > 1 else None
        growth_rate = None
        if previous_y not in (None, 0):
            growth_rate = safe_divide(latest_y, previous_y)
            if growth_rate is not None:
                growth_rate -= 1

        range_width = None
        if forecast_max is not None and forecast_min is not None:
            range_width = forecast_max - forecast_min

        rows.append(
            {
                "model": file_meta["model"],
                "ticker": file_meta["ticker"],
                "model_period": file_meta["model_period"],
                "model_date": file_meta["model_date"],
                "method": "empirical",
                "parameter_name": "avg_penetration_pct",
                "parameter_value": avg_penetration,
                "num_quarters_used": n_quarters,
                "last_quarter_used": quarter_to_text(window[0]["quarter"]),
                "forecast_value": forecast_value,
                "actual_value": latest_y,
                "forecast_max": forecast_max,
                "forecast_min": forecast_min,
                "range_width": range_width,
                "avg_penetration_pct": avg_penetration,
                "quarterly_sales": latest_x,
                "reported_sales": latest_y,
                "growth_rate_pct": growth_rate,
                "sales_captured_in_db_pct": safe_divide(latest_x, latest_y),
                "source_file": source_file,
            }
        )

    helper_cell.clear_contents()
    return rows


def build_regression_rows(
    workbook: xw.Book,
    sheet: xw.Sheet,
    file_meta: Dict[str, str],
    source_file: str,
) -> List[Dict[str, Any]]:
    snapshot = SheetSnapshot(sheet)
    anchor = choose_anchor_max(snapshot)
    if anchor is None:
        print(f"Skipped regression extraction for {source_file}: 'max' anchor not found")
        return []

    anchor_row, anchor_col = anchor
    y_col = anchor_col - 7
    x_col = anchor_col - 11
    quarter_col = x_col - 1
    history = collect_recent_history(snapshot, anchor_row, quarter_col, x_col, y_col)

    if len(history) < 2:
        print(f"Skipped regression extraction for {source_file}: not enough data points")
        return []

    base_max, base_min = read_anchor_min_max(snapshot, anchor)
    helper_intercept = sheet.range((anchor_row + 6, anchor_col + 24))
    helper_slope = sheet.range((anchor_row + 7, anchor_col + 24))
    rows: List[Dict[str, Any]] = []
    previous_signature: Optional[Tuple[Any, ...]] = None

    max_quarters = min(MAX_QUARTERS, len(history))
    for n_quarters in range(2, max_quarters + 1):
        window = history[-n_quarters:]
        start_row = window[0]["row"]
        end_row = window[-1]["row"]
        latest_x = window[-1]["x"]

        intercept_formula = (
            f"=INTERCEPT(R{start_row}C{y_col}:R{end_row}C{y_col},"
            f"R{start_row}C{x_col}:R{end_row}C{x_col})"
        )
        slope_formula = (
            f"=SLOPE(R{start_row}C{y_col}:R{end_row}C{y_col},"
            f"R{start_row}C{x_col}:R{end_row}C{x_col})"
        )
        set_formula2(helper_intercept, intercept_formula)
        set_formula2(helper_slope, slope_formula)
        workbook.app.calculate()

        intercept = safe_float(helper_intercept.value)
        slope = safe_float(helper_slope.value)
        if intercept is None or slope is None:
            continue

        forecast_value = intercept + (slope * latest_x)
        residuals = [item["y"] - (intercept + slope * item["x"]) for item in window]
        sigma = pstdev(residuals) if len(residuals) > 1 else 0.0

        forecast_max = base_max if base_max is not None else forecast_value + (2 * sigma)
        forecast_min = base_min if base_min is not None else forecast_value - (2 * sigma)
        range_width = None
        if forecast_max is not None and forecast_min is not None:
            range_width = forecast_max - forecast_min

        signature = (
            n_quarters,
            round(intercept, 8),
            round(slope, 8),
            round(forecast_value, 8),
            None if forecast_max is None else round(forecast_max, 8),
            None if forecast_min is None else round(forecast_min, 8),
        )
        if previous_signature == signature:
            continue
        previous_signature = signature

        rows.append(
            {
                "model": file_meta["model"],
                "ticker": file_meta["ticker"],
                "model_period": file_meta["model_period"],
                "model_date": file_meta["model_date"],
                "method": "regression",
                "parameter_name": "num_quarters_used",
                "parameter_value": n_quarters,
                "num_quarters_used": n_quarters,
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

    helper_intercept.clear_contents()
    helper_slope.clear_contents()
    return rows


def write_sheet_rows(sheet, headers: Sequence[str], rows: Sequence[Dict[str, Any]]) -> None:
    sheet.append(list(headers))
    for row in rows:
        sheet.append([row.get(column) for column in headers])

    for header_cell in sheet[1]:
        header_cell.font = Font(bold=True)

    sheet.freeze_panes = "A2"
    last_col = get_column_letter(len(headers))
    sheet.auto_filter.ref = f"A1:{last_col}{max(sheet.max_row, 1)}"

    for idx, column in enumerate(headers, start=1):
        max_length = len(column)
        for row_idx in range(2, sheet.max_row + 1):
            value = sheet.cell(row=row_idx, column=idx).value
            if value is None:
                continue
            value_length = len(str(value))
            if value_length > max_length:
                max_length = value_length
        sheet.column_dimensions[get_column_letter(idx)].width = min(max_length + 2, 44)


def write_output_workbook(
    output_path: Path,
    empirical_rows: Sequence[Dict[str, Any]],
    regression_rows: Sequence[Dict[str, Any]],
) -> None:
    workbook = Workbook()
    empirical_sheet = workbook.active
    empirical_sheet.title = "empirical_candidates"
    write_sheet_rows(empirical_sheet, EMPIRICAL_HEADERS, empirical_rows)

    regression_sheet = workbook.create_sheet("regression_candidates")
    write_sheet_rows(regression_sheet, REGRESSION_HEADERS, regression_rows)
    workbook.save(output_path)


def should_skip_file(file_path: Path) -> Optional[str]:
    if not file_path.is_file():
        return "not a file"
    if file_path.name.startswith("~"):
        return "temporary file"
    if file_path.suffix.lower() != ".xlsx":
        return "not an .xlsx file"
    if re.search(r"_PARAM(?:\.\d+)?\.xlsx$", file_path.name, re.IGNORECASE):
        return "prior output file"
    return None


def main() -> None:
    source_dir = Path(input_dir).expanduser().resolve()
    destination_dir = Path(output_dir).expanduser().resolve()

    if not source_dir.exists():
        print(f"Input directory does not exist: {source_dir}")
        return

    output_path = resolve_output_path(source_dir, destination_dir)

    processed_files = 0
    empirical_rows: List[Dict[str, Any]] = []
    regression_rows: List[Dict[str, Any]] = []

    app: Optional[xw.App] = None
    try:
        app = xw.App(visible=False, add_book=False)
        app.display_alerts = False
        app.screen_updating = False

        for file_path in sorted(source_dir.iterdir(), key=lambda p: p.name.lower()):
            skip_reason = should_skip_file(file_path)
            if skip_reason:
                print(f"Skipped {file_path.name}: {skip_reason}")
                continue

            print(f"Processing file: {file_path.name}")
            workbook: Optional[xw.Book] = None
            try:
                workbook = app.books.open(str(file_path), update_links=False)
                file_meta = parse_file_labels(file_path)

                empirical_sheet = get_sheet(workbook, "Empirical Model")
                if empirical_sheet is None:
                    print(f"Skipped empirical extraction for {file_path.name}: missing sheet 'Empirical Model'")
                else:
                    empirical_rows.extend(
                        build_empirical_rows(workbook, empirical_sheet, file_meta, file_path.name)
                    )

                regression_sheet = get_sheet(workbook, "Regression Model")
                if regression_sheet is None:
                    print(
                        f"Skipped regression extraction for {file_path.name}: missing sheet 'Regression Model'"
                    )
                else:
                    regression_rows.extend(
                        build_regression_rows(workbook, regression_sheet, file_meta, file_path.name)
                    )

                processed_files += 1
            except Exception as exc:
                print(f"Skipped {file_path.name}: processing error ({exc})")
            finally:
                if workbook is not None:
                    safe_close_workbook(workbook)
    finally:
        if app is not None:
            try:
                app.quit()
            except Exception as exc:
                print(f"Warning: failed to quit Excel app cleanly ({exc})")

    write_output_workbook(output_path, empirical_rows, regression_rows)

    print(f"Output path: {output_path}")
    print(f"Number of files processed: {processed_files}")
    print(f"Number of empirical rows: {len(empirical_rows)}")
    print(f"Number of regression rows: {len(regression_rows)}")


if __name__ == "__main__":
    main()
