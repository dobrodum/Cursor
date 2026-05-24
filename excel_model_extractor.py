from __future__ import annotations

from dataclasses import dataclass
from datetime import date
from pathlib import Path
import math
import re
from typing import Any, Iterable

import xlwings as xw
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter


# Configure these paths before running.
input_dir = r"/path/to/input/folder"
output_dir = r"/path/to/output/folder"


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

MONTH_MAP = {
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

PERIOD_DAY_MAP = {"early": 5, "mid": 15, "late": 25}


@dataclass
class ParsedFileLabel:
    model: str
    ticker: str
    model_period: str
    model_date: str


@dataclass
class SheetSnapshot:
    start_row: int
    start_col: int
    values: list[list[Any]]
    labels: dict[str, list[tuple[int, int]]]

    @property
    def n_rows(self) -> int:
        return len(self.values)

    @property
    def n_cols(self) -> int:
        return len(self.values[0]) if self.values else 0

    @property
    def end_row(self) -> int:
        return self.start_row + self.n_rows - 1

    @property
    def end_col(self) -> int:
        return self.start_col + self.n_cols - 1

    def get(self, row: int, col: int) -> Any:
        if row < self.start_row or col < self.start_col:
            return None
        r_idx = row - self.start_row
        c_idx = col - self.start_col
        if r_idx >= self.n_rows or c_idx >= self.n_cols:
            return None
        return self.values[r_idx][c_idx]

    def find_exact(self, label: str) -> list[tuple[int, int]]:
        return self.labels.get(normalize_label(label), [])

    def find_contains(self, needle: str) -> list[tuple[int, int]]:
        norm_needle = normalize_label(needle)
        matches: list[tuple[int, int]] = []
        for norm_label, coords in self.labels.items():
            if norm_needle in norm_label:
                matches.extend(coords)
        return matches


def normalize_label(value: Any) -> str:
    if value is None:
        return ""
    text = str(value).strip().lower()
    text = re.sub(r"\s+", " ", text)
    return text


def to_float(value: Any) -> float | None:
    if value is None or value == "":
        return None
    if isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        if isinstance(value, float) and (math.isnan(value) or math.isinf(value)):
            return None
        return float(value)

    text = str(value).strip()
    if not text:
        return None
    text = text.replace(",", "").replace("$", "")
    if text.endswith("%"):
        text = text[:-1]
    try:
        return float(text)
    except ValueError:
        return None


def as_pct_fraction(value: Any) -> float | None:
    numeric = to_float(value)
    if numeric is None:
        return None
    if numeric > 1:
        return numeric / 100.0
    if numeric < 0:
        return None
    return numeric


def safe_round(value: Any, ndigits: int = 6) -> float | None:
    numeric = to_float(value)
    if numeric is None:
        return None
    return round(numeric, ndigits)


def parse_model_label(file_path: Path) -> ParsedFileLabel:
    stem = file_path.stem
    parts = [part.strip() for part in stem.split(" - ")]

    ticker = "UNKNOWN"
    raw_period = ""
    if len(parts) >= 3:
        ticker = parts[1] or "UNKNOWN"
        raw_period = parts[2]
    elif len(parts) >= 2:
        ticker = parts[1] or "UNKNOWN"

    raw_period = raw_period.split("_")[0].strip()
    period_match = re.match(r"^(Early|Mid|Late)([A-Za-z]+)(\d{4})$", raw_period, flags=re.IGNORECASE)

    if period_match:
        period_bucket = period_match.group(1).title()
        month_token = period_match.group(2)
        year_text = period_match.group(3)
        month_key = month_token.strip().lower()
        month_num = MONTH_MAP.get(month_key)
        if month_num is None and len(month_key) >= 3:
            month_num = MONTH_MAP.get(month_key[:3])
        if month_num is None:
            month_num = 1

        month_label = date(2000, month_num, 1).strftime("%b")
        model_period = f"{period_bucket}{month_label}_{year_text}"
        day = PERIOD_DAY_MAP[period_bucket.lower()]
        model_date = date(int(year_text), month_num, day).isoformat()
    else:
        model_period = "Unknown_0000"
        model_date = ""

    model = f"{ticker}_{model_period}"
    return ParsedFileLabel(model=model, ticker=ticker, model_period=model_period, model_date=model_date)


def list_input_files(folder: Path) -> tuple[list[Path], list[tuple[str, str]]]:
    valid_files: list[Path] = []
    skipped: list[tuple[str, str]] = []
    if not folder.exists():
        return valid_files, [("", f"input directory does not exist: {folder}")]

    for file_path in sorted(folder.iterdir()):
        if not file_path.is_file():
            continue
        name = file_path.name
        lower_name = name.lower()

        if name.startswith("~"):
            skipped.append((name, "temporary workbook"))
            continue
        if not lower_name.endswith(".xlsx"):
            skipped.append((name, "not an .xlsx file"))
            continue
        if "_param" in lower_name:
            skipped.append((name, "already a PARAM output workbook"))
            continue
        valid_files.append(file_path)

    return valid_files, skipped


def build_output_path(input_folder: Path, output_folder: Path) -> Path:
    output_folder.mkdir(parents=True, exist_ok=True)
    base_name = f"{input_folder.name}_PARAM.xlsx"
    base_path = output_folder / base_name
    if not base_path.exists():
        return base_path

    idx = 1
    while True:
        candidate = output_folder / f"{input_folder.name}_PARAM.{idx}.xlsx"
        if not candidate.exists():
            return candidate
        idx += 1


def build_sheet_snapshot(sheet: xw.Sheet) -> SheetSnapshot:
    used = sheet.used_range
    start_row = used.row
    start_col = used.column
    raw_values = used.value

    if raw_values is None:
        values = [[]]
    elif isinstance(raw_values, list):
        if raw_values and not isinstance(raw_values[0], list):
            values = [raw_values]
        else:
            values = raw_values
    else:
        values = [[raw_values]]

    labels: dict[str, list[tuple[int, int]]] = {}
    for r_idx, row_values in enumerate(values):
        for c_idx, cell_value in enumerate(row_values):
            norm = normalize_label(cell_value)
            if norm:
                labels.setdefault(norm, []).append((start_row + r_idx, start_col + c_idx))

    return SheetSnapshot(start_row=start_row, start_col=start_col, values=values, labels=labels)


def find_max_anchor(snapshot: SheetSnapshot) -> tuple[int, int] | None:
    exact = snapshot.find_exact("max")
    if exact:
        return exact[0]

    starts_with_max: list[tuple[int, int]] = []
    for label, coords in snapshot.labels.items():
        if label.startswith("max "):
            starts_with_max.extend(coords)
    if starts_with_max:
        return starts_with_max[0]
    return None


def lookup_value_near_label(
    snapshot: SheetSnapshot,
    label_terms: Iterable[str],
    neighbor_offsets: Iterable[tuple[int, int]] = ((0, 1), (1, 0), (0, 2), (2, 0)),
) -> Any:
    for term in label_terms:
        matches = snapshot.find_contains(term)
        for row, col in matches:
            for row_offset, col_offset in neighbor_offsets:
                candidate = snapshot.get(row + row_offset, col + col_offset)
                if candidate not in (None, ""):
                    return candidate
    return None


def value_right_of(snapshot: SheetSnapshot, row: int, col: int, distance: int = 1) -> Any:
    return snapshot.get(row, col + distance)


def detect_anchor_min(snapshot: SheetSnapshot, anchor_row: int, anchor_col: int) -> Any:
    if normalize_label(snapshot.get(anchor_row + 1, anchor_col)) == "min":
        val = value_right_of(snapshot, anchor_row + 1, anchor_col)
        if val not in (None, ""):
            return val
    min_cells = snapshot.find_exact("min")
    if min_cells:
        min_row, min_col = min_cells[0]
        val = value_right_of(snapshot, min_row, min_col)
        if val not in (None, ""):
            return val
    return snapshot.get(anchor_row + 1, anchor_col + 1)


def get_row_numeric_series(snapshot: SheetSnapshot, row: int, col_start: int, col_end: int) -> list[tuple[int, float]]:
    values: list[tuple[int, float]] = []
    for col in range(col_start, col_end + 1):
        numeric = to_float(snapshot.get(row, col))
        if numeric is not None:
            values.append((col, numeric))
    return values


def find_empirical_series_rows(snapshot: SheetSnapshot, anchor_row: int, anchor_col: int) -> tuple[int | None, int | None]:
    row_lo = max(snapshot.start_row, anchor_row - 30)
    row_hi = min(snapshot.end_row, anchor_row + 6)
    col_lo = max(snapshot.start_col, anchor_col - 24)
    col_hi = max(col_lo, anchor_col - 1)

    best_pen_row = None
    best_pen_score = -1
    best_sales_row = None
    best_sales_score = -1

    for row in range(row_lo, row_hi + 1):
        series = get_row_numeric_series(snapshot, row, col_lo, col_hi)
        if len(series) < 3:
            continue
        numeric_values = [value for _, value in series]
        in_pct_shape = sum(1 for value in numeric_values if 0 <= value <= 100)
        if in_pct_shape == len(numeric_values):
            score = len(series)
            if score > best_pen_score:
                best_pen_score = score
                best_pen_row = row

        sales_like = sum(1 for value in numeric_values if abs(value) > 1)
        if sales_like >= max(3, len(numeric_values) // 2):
            score = len(series)
            if score > best_sales_score:
                best_sales_score = score
                best_sales_row = row

    return best_pen_row, best_sales_row


def row_header_value(snapshot: SheetSnapshot, row: int, col: int) -> str:
    for candidate_row in (row - 1, row - 2):
        header = snapshot.get(candidate_row, col)
        if header not in (None, ""):
            return str(header)
    return ""


def is_duplicate_regression_row(last_row: dict[str, Any] | None, current_row: dict[str, Any]) -> bool:
    if last_row is None:
        return False
    fields = ("num_quarters_used", "forecast_value", "forecast_max", "forecast_min", "intercept", "slope")
    for field in fields:
        left = to_float(last_row.get(field))
        right = to_float(current_row.get(field))
        if left is None and right is None:
            continue
        if left is None or right is None:
            return False
        if abs(left - right) > 1e-9:
            return False
    return True


def safe_close_workbook(wb: xw.Book) -> None:
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

    wb.app.display_alerts = False
    wb.close()


def process_empirical_sheet(
    wb: xw.Book,
    sheet: xw.Sheet,
    parsed: ParsedFileLabel,
    source_file: str,
) -> list[dict[str, Any]]:
    snapshot = build_sheet_snapshot(sheet)
    anchor = find_max_anchor(snapshot)
    if not anchor:
        return []
    anchor_row, anchor_col = anchor

    anchor_max = to_float(value_right_of(snapshot, anchor_row, anchor_col))
    anchor_min = to_float(detect_anchor_min(snapshot, anchor_row, anchor_col))
    baseline_forecast = to_float(
        lookup_value_near_label(
            snapshot,
            ("estimated total sold", "forecast total", "tot fcst", "total sold"),
        )
    )
    reported_sales_label_val = to_float(
        lookup_value_near_label(snapshot, ("reported sales", "actual sales", "actual"))
    )

    pen_row, sales_row = find_empirical_series_rows(snapshot, anchor_row, anchor_col)
    if pen_row is None:
        return []

    col_lo = max(snapshot.start_col, anchor_col - 24)
    col_hi = max(col_lo, anchor_col - 1)

    penetration_series_raw = get_row_numeric_series(snapshot, pen_row, col_lo, col_hi)
    penetration_series = [(col, as_pct_fraction(value)) for col, value in penetration_series_raw]
    penetration_series = [(col, value) for col, value in penetration_series if value is not None]
    if not penetration_series:
        return []

    sales_by_col: dict[int, float] = {}
    if sales_row is not None:
        for col, value in get_row_numeric_series(snapshot, sales_row, col_lo, col_hi):
            sales_by_col[col] = value

    n_quarters = min(10, len(penetration_series))
    scratch_col = max(anchor_col + 6, snapshot.end_col + 2)
    scratch_row_start = max(anchor_row + 6, snapshot.end_row + 2)

    avg_cells: list[tuple[int, xw.Range, int, int]] = []
    for idx in range(n_quarters):
        use_n = idx + 1
        selected = penetration_series[-use_n:]
        start_col = selected[0][0]
        end_col = selected[-1][0]
        cell = sheet.range((scratch_row_start + idx, scratch_col))
        cell.formula2 = f'=IFERROR(AVERAGE(R{pen_row}C{start_col}:R{pen_row}C{end_col}),"")'
        avg_cells.append((use_n, cell, start_col, end_col))

    wb.app.calculate()

    max_ratio = None
    min_ratio = None
    if baseline_forecast and baseline_forecast != 0 and anchor_max is not None and anchor_min is not None:
        max_ratio = anchor_max / baseline_forecast
        min_ratio = anchor_min / baseline_forecast

    rows: list[dict[str, Any]] = []
    for use_n, avg_cell, start_col, end_col in avg_cells:
        avg_pen = as_pct_fraction(avg_cell.value)
        if avg_pen is None or avg_pen <= 0:
            continue

        selected_cols = [col for col, _ in penetration_series[-use_n:]]
        selected_sales = [sales_by_col[col] for col in selected_cols if col in sales_by_col]

        reported_sales = reported_sales_label_val
        if reported_sales is None and selected_sales:
            reported_sales = selected_sales[-1]
        quarterly_sales = selected_sales[-1] if selected_sales else None

        forecast_value = None
        if reported_sales is not None and avg_pen > 0:
            forecast_value = reported_sales / avg_pen

        if forecast_value is not None:
            if max_ratio is not None and min_ratio is not None:
                forecast_max = forecast_value * max_ratio
                forecast_min = forecast_value * min_ratio
            elif anchor_max is not None and anchor_min is not None:
                forecast_max = anchor_max
                forecast_min = anchor_min
            else:
                forecast_max = forecast_value * 1.1
                forecast_min = forecast_value * 0.9
        else:
            forecast_max = anchor_max
            forecast_min = anchor_min

        range_width = None
        if forecast_max is not None and forecast_min is not None:
            range_width = forecast_max - forecast_min

        growth_rate_pct = None
        if len(selected_sales) >= 2 and selected_sales[-2] != 0:
            growth_rate_pct = ((selected_sales[-1] - selected_sales[-2]) / abs(selected_sales[-2])) * 100.0

        sales_captured_pct = None
        if forecast_value and reported_sales is not None and forecast_value != 0:
            sales_captured_pct = (reported_sales / forecast_value) * 100.0

        row = {
            "model": parsed.model,
            "ticker": parsed.ticker,
            "model_period": parsed.model_period,
            "model_date": parsed.model_date,
            "method": "empirical",
            "parameter_name": "avg_penetration_pct",
            "parameter_value": safe_round(avg_pen * 100.0, 6),
            "num_quarters_used": use_n,
            "last_quarter_used": row_header_value(snapshot, pen_row, end_col),
            "forecast_value": safe_round(forecast_value, 6),
            "actual_value": safe_round(reported_sales, 6),
            "forecast_max": safe_round(forecast_max, 6),
            "forecast_min": safe_round(forecast_min, 6),
            "range_width": safe_round(range_width, 6),
            "avg_penetration_pct": safe_round(avg_pen * 100.0, 6),
            "quarterly_sales": safe_round(quarterly_sales, 6),
            "reported_sales": safe_round(reported_sales, 6),
            "growth_rate_pct": safe_round(growth_rate_pct, 6),
            "sales_captured_in_db_pct": safe_round(sales_captured_pct, 6),
            "source_file": source_file,
        }
        rows.append(row)

    return rows


def process_regression_sheet(
    wb: xw.Book,
    sheet: xw.Sheet,
    parsed: ParsedFileLabel,
    source_file: str,
) -> list[dict[str, Any]]:
    snapshot = build_sheet_snapshot(sheet)
    anchor = find_max_anchor(snapshot)
    if not anchor:
        return []
    anchor_row, anchor_col = anchor

    x_col = anchor_col - 11
    y_col = anchor_col - 7
    if x_col < snapshot.start_col or y_col < snapshot.start_col:
        return []

    rows_with_xy: list[tuple[int, float, float]] = []
    for row in range(snapshot.start_row, min(anchor_row, snapshot.end_row + 1)):
        x_val = to_float(snapshot.get(row, x_col))
        y_val = to_float(snapshot.get(row, y_col))
        if x_val is None or y_val is None:
            continue
        rows_with_xy.append((row, x_val, y_val))

    if len(rows_with_xy) < 2:
        return []

    anchor_max = to_float(value_right_of(snapshot, anchor_row, anchor_col))
    anchor_min = to_float(detect_anchor_min(snapshot, anchor_row, anchor_col))
    baseline_forecast = to_float(
        lookup_value_near_label(snapshot, ("tot fcst w/o sa", "forecast w/o sa", "total forecast"))
    )
    actual_val = to_float(lookup_value_near_label(snapshot, ("actual", "reported sales")))

    n_quarters = min(10, len(rows_with_xy))
    scratch_col = max(anchor_col + 6, snapshot.end_col + 2)
    scratch_row_start = max(anchor_row + 6, snapshot.end_row + 2)

    calc_rows: list[dict[str, Any]] = []
    for idx in range(n_quarters):
        use_n = idx + 1
        selected = rows_with_xy[-use_n:]
        row_start = selected[0][0]
        row_end = selected[-1][0]
        next_x = selected[-1][1] + 1
        calc_row = scratch_row_start + idx

        intercept_cell = sheet.range((calc_row, scratch_col))
        slope_cell = sheet.range((calc_row, scratch_col + 1))
        forecast_cell = sheet.range((calc_row, scratch_col + 2))

        intercept_cell.formula2 = (
            f'=IFERROR(INTERCEPT(R{row_start}C{y_col}:R{row_end}C{y_col},'
            f'R{row_start}C{x_col}:R{row_end}C{x_col}),"")'
        )
        slope_cell.formula2 = (
            f'=IFERROR(SLOPE(R{row_start}C{y_col}:R{row_end}C{y_col},'
            f'R{row_start}C{x_col}:R{row_end}C{x_col}),"")'
        )
        forecast_cell.formula2 = (
            f'=IFERROR(R{calc_row}C{scratch_col}+R{calc_row}C{scratch_col + 1}*{next_x},"")'
        )

        calc_rows.append(
            {
                "n": use_n,
                "intercept_cell": intercept_cell,
                "slope_cell": slope_cell,
                "forecast_cell": forecast_cell,
            }
        )

    wb.app.calculate()

    max_ratio = None
    min_ratio = None
    if baseline_forecast and baseline_forecast != 0 and anchor_max is not None and anchor_min is not None:
        max_ratio = anchor_max / baseline_forecast
        min_ratio = anchor_min / baseline_forecast

    rows: list[dict[str, Any]] = []
    last_row: dict[str, Any] | None = None
    for calc in calc_rows:
        intercept_val = to_float(calc["intercept_cell"].value)
        slope_val = to_float(calc["slope_cell"].value)
        forecast_val = to_float(calc["forecast_cell"].value)
        if forecast_val is None and intercept_val is not None and slope_val is not None:
            forecast_val = intercept_val + slope_val

        if forecast_val is not None:
            if max_ratio is not None and min_ratio is not None:
                forecast_max = forecast_val * max_ratio
                forecast_min = forecast_val * min_ratio
            elif anchor_max is not None and anchor_min is not None:
                forecast_max = anchor_max
                forecast_min = anchor_min
            else:
                forecast_max = forecast_val * 1.1
                forecast_min = forecast_val * 0.9
        else:
            forecast_max = anchor_max
            forecast_min = anchor_min

        range_width = None
        if forecast_max is not None and forecast_min is not None:
            range_width = forecast_max - forecast_min

        out_row = {
            "model": parsed.model,
            "ticker": parsed.ticker,
            "model_period": parsed.model_period,
            "model_date": parsed.model_date,
            "method": "regression",
            "parameter_name": "num_quarters_used",
            "parameter_value": calc["n"],
            "num_quarters_used": calc["n"],
            "forecast_value": safe_round(forecast_val, 6),
            "actual_value": safe_round(actual_val, 6),
            "forecast_max": safe_round(forecast_max, 6),
            "forecast_min": safe_round(forecast_min, 6),
            "range_width": safe_round(range_width, 6),
            "intercept": safe_round(intercept_val, 6),
            "slope": safe_round(slope_val, 6),
            "source_file": source_file,
        }

        if is_duplicate_regression_row(last_row, out_row):
            continue
        rows.append(out_row)
        last_row = out_row

    return rows


def write_output_workbook(
    output_path: Path,
    empirical_rows: list[dict[str, Any]],
    regression_rows: list[dict[str, Any]],
) -> None:
    wb = Workbook()
    default_sheet = wb.active
    wb.remove(default_sheet)

    def write_sheet(title: str, columns: list[str], rows: list[dict[str, Any]]) -> None:
        ws = wb.create_sheet(title=title)
        ws.append(columns)
        for row in rows:
            ws.append([row.get(col, "") for col in columns])

        for cell in ws[1]:
            cell.font = Font(bold=True)

        ws.freeze_panes = "A2"
        ws.auto_filter.ref = f"A1:{get_column_letter(len(columns))}{max(1, ws.max_row)}"

        for col_idx, column in enumerate(columns, start=1):
            max_len = len(column)
            for value in ws.iter_cols(min_col=col_idx, max_col=col_idx, min_row=2, max_row=ws.max_row, values_only=True):
                for item in value:
                    if item is None:
                        continue
                    max_len = max(max_len, len(str(item)))
            ws.column_dimensions[get_column_letter(col_idx)].width = min(max_len + 2, 42)

    write_sheet("empirical_candidates", EMPIRICAL_COLUMNS, empirical_rows)
    write_sheet("regression_candidates", REGRESSION_COLUMNS, regression_rows)
    wb.save(output_path)


def main() -> None:
    input_folder = Path(input_dir).expanduser().resolve()
    output_folder = Path(output_dir).expanduser().resolve()

    files, skipped = list_input_files(input_folder)
    for skipped_name, reason in skipped:
        print(f"SKIPPED: {skipped_name or '(n/a)'} - {reason}")

    if not files:
        output_path = build_output_path(input_folder, output_folder)
        write_output_workbook(output_path, [], [])
        print(f"Output workbook written: {output_path}")
        print("Files processed: 0")
        print("Empirical rows: 0")
        print("Regression rows: 0")
        return

    app = xw.App(visible=False, add_book=False)
    app.display_alerts = False
    app.screen_updating = False

    empirical_rows: list[dict[str, Any]] = []
    regression_rows: list[dict[str, Any]] = []
    processed_count = 0

    try:
        for file_path in files:
            print(f"PROCESSING: {file_path.name}")
            wb = None
            try:
                wb = app.books.open(str(file_path), update_links=False)
                parsed = parse_model_label(file_path)

                empirical_sheet = None
                regression_sheet = None
                for sheet in wb.sheets:
                    normalized = normalize_label(sheet.name)
                    if normalized == "empirical model":
                        empirical_sheet = sheet
                    elif normalized == "regression model":
                        regression_sheet = sheet

                if empirical_sheet is None:
                    print(f"SKIPPED empirical extraction: {file_path.name} - missing sheet 'Empirical Model'")
                else:
                    empirical_rows.extend(
                        process_empirical_sheet(wb=wb, sheet=empirical_sheet, parsed=parsed, source_file=file_path.name)
                    )

                if regression_sheet is None:
                    print(f"SKIPPED regression extraction: {file_path.name} - missing sheet 'Regression Model'")
                else:
                    regression_rows.extend(
                        process_regression_sheet(wb=wb, sheet=regression_sheet, parsed=parsed, source_file=file_path.name)
                    )

                processed_count += 1
            except Exception as exc:
                print(f"SKIPPED: {file_path.name} - processing error: {exc}")
            finally:
                if wb is not None:
                    safe_close_workbook(wb)
    finally:
        app.quit()

    output_path = build_output_path(input_folder, output_folder)
    write_output_workbook(output_path, empirical_rows, regression_rows)

    print(f"Output workbook written: {output_path}")
    print(f"Files processed: {processed_count}")
    print(f"Empirical rows: {len(empirical_rows)}")
    print(f"Regression rows: {len(regression_rows)}")


if __name__ == "__main__":
    main()
