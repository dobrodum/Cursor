#!/usr/bin/env python3
from __future__ import annotations

import re
from dataclasses import dataclass
from datetime import date
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Sequence, Tuple

import xlwings as xw
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

# ----------------------------
# User-configurable directories
# ----------------------------
input_dir = Path("input")
output_dir = Path("output")


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

MONTH_TO_INT = {
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

PERIOD_TO_DAY = {
    "early": 5,
    "mid": 15,
    "late": 25,
}

EMPIRICAL_ALIASES = {
    "num_quarters_used": [
        "num quarters used",
        "number of quarters used",
        "quarters used",
        "num qtrs used",
        "n quarters",
        "n qtrs",
    ],
    "last_quarter_used": [
        "last quarter used",
        "last qtr used",
        "last quarter",
    ],
    "forecast_value": [
        "estimated total sold",
        "est total sold",
        "forecast value",
        "forecast",
        "tot fcst",
        "total forecast",
    ],
    "actual_value": [
        "reported sales",
        "actual sales",
        "actual value",
    ],
    "forecast_min": [
        "min",
        "forecast min",
        "minimum",
    ],
    "avg_penetration_pct": [
        "avg penetration pct",
        "avg penetration",
        "average penetration pct",
        "average penetration",
        "penetration pct",
    ],
    "quarterly_sales": [
        "quarterly sales",
        "quarter sales",
        "qtr sales",
    ],
    "reported_sales": [
        "reported sales",
        "reported",
    ],
    "growth_rate_pct": [
        "growth rate pct",
        "growth rate",
        "growth %",
    ],
    "sales_captured_in_db_pct": [
        "sales captured in db pct",
        "sales captured in db",
        "captured in db pct",
        "captured in db",
    ],
}

REGRESSION_ALIASES = {
    "num_quarters_used": [
        "num quarters used",
        "number of quarters used",
        "quarters used",
        "num qtrs used",
        "n quarters",
        "n qtrs",
    ],
    "forecast_value": [
        "tot fcst w o sa",
        "tot fcst wo sa",
        "tot fcst without sa",
        "total forecast w o sa",
        "forecast total without sa",
        "tot fcst",
        "total forecast",
    ],
    "forecast_min": [
        "min",
        "forecast min",
        "minimum",
    ],
    "actual_value": [
        "actual value",
        "actual sales",
        "reported sales",
    ],
}


@dataclass
class FileMetadata:
    model: str
    ticker: str
    model_period: str
    model_date: str


class SheetSnapshot:
    def __init__(self, sheet: xw.Sheet) -> None:
        self.sheet = sheet
        self.used_range = sheet.used_range
        self.first_row = self.used_range.row
        self.first_col = self.used_range.column
        self.row_count = self.used_range.rows.count
        self.col_count = self.used_range.columns.count
        self.last_row = self.first_row + self.row_count - 1
        self.last_col = self.first_col + self.col_count - 1
        self.matrix = self._to_matrix(self.used_range.value, self.row_count, self.col_count)

    @staticmethod
    def _to_matrix(values: Any, n_rows: int, n_cols: int) -> List[List[Any]]:
        if n_rows == 1 and n_cols == 1:
            return [[values]]
        if n_rows == 1:
            if isinstance(values, (list, tuple)):
                return [list(values)]
            return [[values]]
        if n_cols == 1:
            if isinstance(values, (list, tuple)):
                return [[v] for v in values]
            return [[values]]
        if not isinstance(values, (list, tuple)):
            return [[values]]
        matrix: List[List[Any]] = []
        for row in values:
            if isinstance(row, (list, tuple)):
                matrix.append(list(row))
            else:
                matrix.append([row])
        return matrix

    def value_at(self, row: int, col: int) -> Any:
        r_idx = row - self.first_row
        c_idx = col - self.first_col
        if r_idx < 0 or c_idx < 0:
            return self.sheet.range((row, col)).value
        if r_idx >= len(self.matrix):
            return self.sheet.range((row, col)).value
        row_values = self.matrix[r_idx]
        if c_idx >= len(row_values):
            return self.sheet.range((row, col)).value
        return row_values[c_idx]

    def row_values(self, row: int) -> List[Any]:
        r_idx = row - self.first_row
        if 0 <= r_idx < len(self.matrix):
            return self.matrix[r_idx]
        return []


def normalize_text(value: Any) -> str:
    if value is None:
        return ""
    text = str(value).strip().lower()
    text = text.replace("%", " pct ")
    text = text.replace("&", " and ")
    text = re.sub(r"[_\-/\n\r]+", " ", text)
    text = re.sub(r"[^a-z0-9\s]", " ", text)
    text = re.sub(r"\s+", " ", text).strip()
    return text


def is_blank(value: Any) -> bool:
    return value is None or (isinstance(value, str) and value.strip() == "")


def as_number(value: Any) -> Optional[float]:
    if isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        return float(value)
    if value is None:
        return None
    text = str(value).strip()
    if not text:
        return None
    is_pct = text.endswith("%")
    text = text.replace(",", "").replace("$", "").replace("%", "")
    try:
        number = float(text)
    except ValueError:
        return None
    if is_pct:
        number /= 100.0
    return number


def to_int_if_whole(value: Any) -> Any:
    number = as_number(value)
    if number is None:
        return value
    if abs(number - round(number)) < 1e-9:
        return int(round(number))
    return number


def parse_file_metadata(file_path: Path) -> FileMetadata:
    stem = file_path.stem

    ticker_match = re.search(r"Model\s*-\s*([A-Za-z0-9]+)\s*-\s*", stem, flags=re.IGNORECASE)
    ticker = ticker_match.group(1).upper() if ticker_match else "UNKNOWN"

    period_match = re.search(
        r"(Early|Mid|Late)(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)(\d{4})",
        stem,
        flags=re.IGNORECASE,
    )

    if not period_match:
        model_period = "UnknownPeriod"
        model_date = ""
        model = f"{ticker}_{model_period}"
        return FileMetadata(model=model, ticker=ticker, model_period=model_period, model_date=model_date)

    period_bucket = period_match.group(1).capitalize()
    month_short = period_match.group(2).capitalize()
    year = int(period_match.group(3))
    model_period = f"{period_bucket}{month_short}_{year}"

    month_num = MONTH_TO_INT[month_short.lower()]
    day = PERIOD_TO_DAY[period_bucket.lower()]
    model_date = date(year, month_num, day).isoformat()
    model = f"{ticker}_{model_period}"
    return FileMetadata(model=model, ticker=ticker, model_period=model_period, model_date=model_date)


def resolve_output_path(src_input_dir: Path, target_output_dir: Path) -> Path:
    target_output_dir.mkdir(parents=True, exist_ok=True)
    stem = f"{src_input_dir.name}_PARAM"
    candidate = target_output_dir / f"{stem}.xlsx"
    suffix = 1
    while candidate.exists():
        candidate = target_output_dir / f"{stem}.{suffix}.xlsx"
        suffix += 1
    return candidate


def safe_close_workbook(wb: xw.Book) -> None:
    try:
        wb.close(save=False)
        return
    except TypeError:
        pass
    except Exception:
        pass

    try:
        wb.api.Close(SaveChanges=False)  # type: ignore[attr-defined]
        return
    except Exception:
        pass

    try:
        wb.close()
    except Exception:
        pass


def find_anchor(snapshot: SheetSnapshot, anchor_label: str = "max") -> Optional[Tuple[int, int]]:
    target = normalize_text(anchor_label)
    for r_idx, row in enumerate(snapshot.matrix):
        for c_idx, value in enumerate(row):
            if normalize_text(value) == target:
                return snapshot.first_row + r_idx, snapshot.first_col + c_idx
    return None


def build_header_map(snapshot: SheetSnapshot, header_row: int) -> Dict[str, List[int]]:
    mapping: Dict[str, List[int]] = {}
    row_values = snapshot.row_values(header_row)
    for c_idx, value in enumerate(row_values):
        normalized = normalize_text(value)
        if not normalized:
            continue
        abs_col = snapshot.first_col + c_idx
        mapping.setdefault(normalized, []).append(abs_col)
    return mapping


def nearest_column(
    header_map: Dict[str, List[int]],
    aliases: Sequence[str],
    anchor_col: int,
    max_distance: int = 18,
) -> Optional[int]:
    alias_set = [normalize_text(alias) for alias in aliases]
    candidates: List[int] = []
    for header_text, cols in header_map.items():
        for alias in alias_set:
            if not alias:
                continue
            if header_text == alias or alias in header_text:
                candidates.extend(cols)
                break
    if not candidates:
        return None
    nearest = min(candidates, key=lambda col: abs(col - anchor_col))
    if abs(nearest - anchor_col) > max_distance:
        return None
    return nearest


def detect_data_rows(
    snapshot: SheetSnapshot,
    header_row: int,
    key_columns: Iterable[Optional[int]],
    max_rows: int = 10,
    scan_limit: int = 300,
) -> List[int]:
    usable_cols = [col for col in key_columns if col is not None]
    if not usable_cols:
        usable_cols = [snapshot.first_col]

    rows: List[int] = []
    empty_streak = 0
    start = header_row + 1
    end = min(snapshot.last_row, header_row + scan_limit)

    for row in range(start, end + 1):
        has_data = False
        for col in usable_cols:
            if not is_blank(snapshot.value_at(row, col)):
                has_data = True
                break
        if has_data:
            rows.append(row)
            empty_streak = 0
            if len(rows) >= max_rows:
                break
        else:
            if rows:
                empty_streak += 1
                if empty_streak >= 3:
                    break
    return rows


def compute_range_width(max_value: Any, min_value: Any) -> Any:
    max_num = as_number(max_value)
    min_num = as_number(min_value)
    if max_num is None or min_num is None:
        return ""
    return max_num - min_num


def read_vertical_values(sheet: xw.Sheet, start_row: int, end_row: int, col: int) -> List[Any]:
    if start_row > end_row:
        return []
    values = sheet.range((start_row, col), (end_row, col)).value
    if start_row == end_row:
        return [values]
    if isinstance(values, (list, tuple)):
        return [item[0] if isinstance(item, (list, tuple)) else item for item in values]
    return [values]


def extract_empirical_rows(wb: xw.Book, metadata: FileMetadata, source_file: str) -> List[Dict[str, Any]]:
    sheet_names = {sheet.name.lower(): sheet for sheet in wb.sheets}
    if "empirical model" not in sheet_names:
        print(f"  Skipped empirical extraction: sheet 'Empirical Model' not found in {source_file}")
        return []

    sheet = sheet_names["empirical model"]
    snapshot = SheetSnapshot(sheet)
    anchor = find_anchor(snapshot, "max")
    if not anchor:
        print(f"  Skipped empirical extraction: 'max' anchor not found in {source_file}")
        return []

    anchor_row, anchor_col = anchor
    header_map = build_header_map(snapshot, anchor_row)

    num_q_col = nearest_column(header_map, EMPIRICAL_ALIASES["num_quarters_used"], anchor_col)
    last_q_col = nearest_column(header_map, EMPIRICAL_ALIASES["last_quarter_used"], anchor_col)
    forecast_col = nearest_column(header_map, EMPIRICAL_ALIASES["forecast_value"], anchor_col)
    actual_col = nearest_column(header_map, EMPIRICAL_ALIASES["actual_value"], anchor_col)
    min_col = nearest_column(header_map, EMPIRICAL_ALIASES["forecast_min"], anchor_col)
    avg_pen_col = nearest_column(header_map, EMPIRICAL_ALIASES["avg_penetration_pct"], anchor_col)
    quarterly_col = nearest_column(header_map, EMPIRICAL_ALIASES["quarterly_sales"], anchor_col)
    reported_col = nearest_column(header_map, EMPIRICAL_ALIASES["reported_sales"], anchor_col)
    growth_col = nearest_column(header_map, EMPIRICAL_ALIASES["growth_rate_pct"], anchor_col)
    captured_col = nearest_column(header_map, EMPIRICAL_ALIASES["sales_captured_in_db_pct"], anchor_col)

    if min_col is None:
        min_col = anchor_col + 1

    data_rows = detect_data_rows(
        snapshot=snapshot,
        header_row=anchor_row,
        key_columns=[num_q_col, forecast_col, anchor_col, min_col],
        max_rows=10,
    )

    if not data_rows:
        print(f"  Skipped empirical extraction: no candidate rows found in {source_file}")
        return []

    avg_values: List[Any] = [None] * len(data_rows)
    if reported_col is not None and quarterly_col is not None:
        helper_col = max(snapshot.last_col + 2, anchor_col + 2)
        formula_values: List[List[str]] = []
        for _ in data_rows:
            rel_reported = reported_col - helper_col
            rel_quarterly = quarterly_col - helper_col
            formula_values.append([f'=IFERROR(RC[{rel_reported}]/RC[{rel_quarterly}], "")'])
        sheet.range((data_rows[0], helper_col), (data_rows[-1], helper_col)).formula2 = formula_values
        wb.app.calculate()
        avg_values = read_vertical_values(sheet, data_rows[0], data_rows[-1], helper_col)
    elif avg_pen_col is not None:
        avg_values = [snapshot.value_at(row, avg_pen_col) for row in data_rows]

    rows: List[Dict[str, Any]] = []
    for idx, row in enumerate(data_rows):
        num_q_value = snapshot.value_at(row, num_q_col) if num_q_col is not None else idx + 1
        last_q_value = snapshot.value_at(row, last_q_col) if last_q_col is not None else ""
        forecast_value = snapshot.value_at(row, forecast_col) if forecast_col is not None else ""
        actual_value = snapshot.value_at(row, actual_col) if actual_col is not None else ""
        forecast_max = snapshot.value_at(row, anchor_col)
        forecast_min = snapshot.value_at(row, min_col) if min_col is not None else ""
        avg_pen = avg_values[idx] if idx < len(avg_values) else ""
        quarterly_sales = snapshot.value_at(row, quarterly_col) if quarterly_col is not None else ""
        reported_sales = snapshot.value_at(row, reported_col) if reported_col is not None else actual_value
        growth_pct = snapshot.value_at(row, growth_col) if growth_col is not None else ""
        captured_pct = snapshot.value_at(row, captured_col) if captured_col is not None else ""

        rows.append(
            {
                "model": metadata.model,
                "ticker": metadata.ticker,
                "model_period": metadata.model_period,
                "model_date": metadata.model_date,
                "method": "empirical",
                "parameter_name": "avg_penetration_pct",
                "parameter_value": avg_pen,
                "num_quarters_used": to_int_if_whole(num_q_value),
                "last_quarter_used": last_q_value,
                "forecast_value": forecast_value,
                "actual_value": actual_value,
                "forecast_max": forecast_max,
                "forecast_min": forecast_min,
                "range_width": compute_range_width(forecast_max, forecast_min),
                "avg_penetration_pct": avg_pen,
                "quarterly_sales": quarterly_sales,
                "reported_sales": reported_sales,
                "growth_rate_pct": growth_pct,
                "sales_captured_in_db_pct": captured_pct,
                "source_file": source_file,
            }
        )
    return rows


def collect_regression_history_rows(
    snapshot: SheetSnapshot,
    x_col: int,
    y_col: int,
    data_end_row: int,
    max_scan_rows: int = 250,
) -> List[int]:
    rows: List[int] = []
    lower_bound = max(snapshot.first_row, data_end_row - max_scan_rows)
    for row in range(data_end_row, lower_bound - 1, -1):
        x_value = as_number(snapshot.value_at(row, x_col))
        y_value = as_number(snapshot.value_at(row, y_col))
        if x_value is None or y_value is None:
            if rows:
                break
            continue
        rows.append(row)
    rows.sort()
    return rows


def extract_regression_rows(wb: xw.Book, metadata: FileMetadata, source_file: str) -> List[Dict[str, Any]]:
    sheet_names = {sheet.name.lower(): sheet for sheet in wb.sheets}
    if "regression model" not in sheet_names:
        print(f"  Skipped regression extraction: sheet 'Regression Model' not found in {source_file}")
        return []

    sheet = sheet_names["regression model"]
    snapshot = SheetSnapshot(sheet)
    anchor = find_anchor(snapshot, "max")
    if not anchor:
        print(f"  Skipped regression extraction: 'max' anchor not found in {source_file}")
        return []

    anchor_row, anchor_col = anchor
    x_col = anchor_col - 11
    y_col = anchor_col - 7
    header_map = build_header_map(snapshot, anchor_row)

    num_q_col = nearest_column(header_map, REGRESSION_ALIASES["num_quarters_used"], anchor_col)
    forecast_col = nearest_column(header_map, REGRESSION_ALIASES["forecast_value"], anchor_col)
    min_col = nearest_column(header_map, REGRESSION_ALIASES["forecast_min"], anchor_col)
    actual_col = nearest_column(header_map, REGRESSION_ALIASES["actual_value"], anchor_col)

    if min_col is None:
        min_col = anchor_col + 1

    data_rows = detect_data_rows(
        snapshot=snapshot,
        header_row=anchor_row,
        key_columns=[num_q_col, forecast_col, anchor_col, min_col],
        max_rows=25,
    )

    row_for_quarters: Dict[int, int] = {}
    for row in data_rows:
        num_q = as_number(snapshot.value_at(row, num_q_col)) if num_q_col is not None else None
        if num_q is None:
            continue
        num_q_int = int(round(num_q))
        if 1 <= num_q_int <= 10 and num_q_int not in row_for_quarters:
            row_for_quarters[num_q_int] = row

    if not row_for_quarters:
        for quarter_count in range(1, 11):
            fallback_row = anchor_row + quarter_count
            row_for_quarters[quarter_count] = fallback_row

    history_rows = collect_regression_history_rows(snapshot, x_col=x_col, y_col=y_col, data_end_row=anchor_row - 1)
    quarter_values = sorted(q for q in row_for_quarters if q <= 10)
    quarter_values = [q for q in quarter_values if q <= len(history_rows)] if history_rows else quarter_values

    intercept_by_quarter: Dict[int, Any] = {}
    slope_by_quarter: Dict[int, Any] = {}
    if history_rows and quarter_values:
        helper_col = max(snapshot.last_col + 2, anchor_col + 2)
        intercept_col = helper_col
        slope_col = helper_col + 1
        for quarter_count in quarter_values:
            start_row = history_rows[-quarter_count]
            end_row = history_rows[-1]
            formula_row = anchor_row + quarter_count
            sheet.range((formula_row, intercept_col)).formula2 = (
                f'=IFERROR(INTERCEPT(R{start_row}C{y_col}:R{end_row}C{y_col},'
                f'R{start_row}C{x_col}:R{end_row}C{x_col}), "")'
            )
            sheet.range((formula_row, slope_col)).formula2 = (
                f'=IFERROR(SLOPE(R{start_row}C{y_col}:R{end_row}C{y_col},'
                f'R{start_row}C{x_col}:R{end_row}C{x_col}), "")'
            )
        wb.app.calculate()
        for quarter_count in quarter_values:
            formula_row = anchor_row + quarter_count
            intercept_by_quarter[quarter_count] = sheet.range((formula_row, intercept_col)).value
            slope_by_quarter[quarter_count] = sheet.range((formula_row, slope_col)).value

    rows: List[Dict[str, Any]] = []
    previous_signature: Optional[Tuple[Any, Any, Any, Any, Any]] = None
    for quarter_count in sorted(row_for_quarters):
        sheet_row = row_for_quarters[quarter_count]
        forecast_value = snapshot.value_at(sheet_row, forecast_col) if forecast_col is not None else ""
        actual_value = snapshot.value_at(sheet_row, actual_col) if actual_col is not None else ""
        forecast_max = snapshot.value_at(sheet_row, anchor_col)
        forecast_min = snapshot.value_at(sheet_row, min_col) if min_col is not None else ""
        intercept = intercept_by_quarter.get(quarter_count, "")
        slope = slope_by_quarter.get(quarter_count, "")

        row_signature = (
            to_int_if_whole(intercept),
            to_int_if_whole(slope),
            to_int_if_whole(forecast_value),
            to_int_if_whole(forecast_max),
            to_int_if_whole(forecast_min),
        )
        if previous_signature is not None and row_signature == previous_signature:
            continue
        previous_signature = row_signature

        if all(is_blank(value) for value in [forecast_value, forecast_max, forecast_min, intercept, slope]):
            continue

        rows.append(
            {
                "model": metadata.model,
                "ticker": metadata.ticker,
                "model_period": metadata.model_period,
                "model_date": metadata.model_date,
                "method": "regression",
                "parameter_name": "num_quarters_used",
                "parameter_value": quarter_count,
                "num_quarters_used": quarter_count,
                "forecast_value": forecast_value,
                "actual_value": actual_value,
                "forecast_max": forecast_max,
                "forecast_min": forecast_min,
                "range_width": compute_range_width(forecast_max, forecast_min),
                "intercept": intercept,
                "slope": slope,
                "source_file": source_file,
            }
        )
    return rows


def format_sheet(ws, columns: List[str]) -> None:
    for col_idx, column_name in enumerate(columns, start=1):
        ws.cell(row=1, column=col_idx, value=column_name)
        ws.cell(row=1, column=col_idx).font = Font(bold=True)
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions

    for col_idx in range(1, len(columns) + 1):
        max_width = len(columns[col_idx - 1])
        for row_idx in range(2, ws.max_row + 1):
            value = ws.cell(row=row_idx, column=col_idx).value
            if value is None:
                continue
            max_width = max(max_width, len(str(value)))
        ws.column_dimensions[get_column_letter(col_idx)].width = min(max_width + 2, 42)


def write_rows(ws, columns: List[str], rows: List[Dict[str, Any]]) -> None:
    for row_data in rows:
        ws.append([row_data.get(column, "") for column in columns])
    format_sheet(ws, columns)


def create_output_workbook(
    output_path: Path,
    empirical_rows: List[Dict[str, Any]],
    regression_rows: List[Dict[str, Any]],
) -> None:
    out_wb = Workbook()
    empirical_ws = out_wb.active
    empirical_ws.title = "empirical_candidates"
    write_rows(empirical_ws, EMPIRICAL_COLUMNS, empirical_rows)

    regression_ws = out_wb.create_sheet("regression_candidates")
    write_rows(regression_ws, REGRESSION_COLUMNS, regression_rows)
    out_wb.save(output_path)


def iter_source_files(src_dir: Path) -> Iterable[Path]:
    for file_path in sorted(src_dir.iterdir()):
        if not file_path.is_file():
            continue
        yield file_path


def main() -> None:
    src_dir = Path(input_dir)
    dst_dir = Path(output_dir)

    if not src_dir.exists():
        print(f"Input directory not found: {src_dir}")
        return

    output_path = resolve_output_path(src_dir, dst_dir)
    generated_pattern = re.compile(rf"^{re.escape(src_dir.name)}_PARAM(?:\.\d+)?\.xlsx$", flags=re.IGNORECASE)

    empirical_rows: List[Dict[str, Any]] = []
    regression_rows: List[Dict[str, Any]] = []
    processed_files = 0

    app: Optional[xw.App] = None
    try:
        app = xw.App(visible=False, add_book=False)
        app.display_alerts = False
        app.screen_updating = False

        for file_path in iter_source_files(src_dir):
            if file_path.name.startswith("~"):
                print(f"Skipped: {file_path.name} (temporary Excel file)")
                continue
            if generated_pattern.match(file_path.name):
                print(f"Skipped: {file_path.name} (generated output file)")
                continue
            if file_path.suffix.lower() != ".xlsx":
                print(f"Skipped: {file_path.name} (not an .xlsx file)")
                continue

            print(f"Processing: {file_path.name}")
            metadata = parse_file_metadata(file_path)
            wb: Optional[xw.Book] = None
            try:
                wb = app.books.open(str(file_path), update_links=False)
                empirical_rows.extend(extract_empirical_rows(wb, metadata, file_path.name))
                regression_rows.extend(extract_regression_rows(wb, metadata, file_path.name))
                processed_files += 1
            except Exception as exc:
                print(f"Skipped: {file_path.name} (error: {exc})")
            finally:
                if wb is not None:
                    safe_close_workbook(wb)
    finally:
        if app is not None:
            app.quit()

    create_output_workbook(output_path, empirical_rows, regression_rows)
    print(f"Output: {output_path}")
    print(f"Files processed: {processed_files}")
    print(f"Empirical rows: {len(empirical_rows)}")
    print(f"Regression rows: {len(regression_rows)}")


if __name__ == "__main__":
    main()
