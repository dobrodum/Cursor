#!/usr/bin/env python3
from __future__ import annotations

import math
import re
from dataclasses import dataclass
from datetime import date
from pathlib import Path
from typing import Any, Dict, List, Optional, Sequence, Tuple

import xlwings as xw
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

# -----------------------------
# User-configurable paths
# -----------------------------
input_dir = Path("/path/to/input")
output_dir = Path("/path/to/output")

# -----------------------------
# Constants
# -----------------------------
N_QUARTERS = 10
EMPIRICAL_MODEL_SHEET = "Empirical Model"
REGRESSION_MODEL_SHEET = "Regression Model"

EMPIRICAL_OUTPUT_SHEET = "empirical_candidates"
REGRESSION_OUTPUT_SHEET = "regression_candidates"

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

DAY_BY_PERIOD_PREFIX = {"early": 5, "mid": 15, "late": 25}
MONTH_BY_ABBREV = {
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


@dataclass
class FileLabels:
    model: str
    ticker: str
    model_period: str
    model_date: str


class SheetSnapshot:
    """In-memory view of a worksheet's used range to reduce COM round-trips."""

    def __init__(self, sheet: xw.Sheet) -> None:
        used = sheet.used_range
        self.start_row = used.row
        self.start_col = used.column
        self.values = self._to_matrix(used.value)
        self.n_rows = len(self.values)
        self.n_cols = len(self.values[0]) if self.values else 0
        self.max_row = self.start_row + self.n_rows - 1
        self.max_col = self.start_col + self.n_cols - 1

    @staticmethod
    def _to_matrix(raw: Any) -> List[List[Any]]:
        if isinstance(raw, tuple):
            raw = list(raw)
        if isinstance(raw, list):
            if not raw:
                return [[]]
            if isinstance(raw[0], (list, tuple)):
                return [list(row) if isinstance(row, tuple) else row for row in raw]
            return [raw]
        return [[raw]]

    def get(self, abs_row: int, abs_col: int) -> Any:
        row_idx = abs_row - self.start_row
        col_idx = abs_col - self.start_col
        if row_idx < 0 or col_idx < 0:
            return None
        if row_idx >= self.n_rows or col_idx >= self.n_cols:
            return None
        return self.values[row_idx][col_idx]

    def row_values(self, abs_row: int) -> List[Any]:
        row_idx = abs_row - self.start_row
        if row_idx < 0 or row_idx >= self.n_rows:
            return []
        return self.values[row_idx]

    def iter_cells(self) -> Sequence[Tuple[int, int, Any]]:
        for r_idx, row_values in enumerate(self.values):
            abs_row = self.start_row + r_idx
            for c_idx, value in enumerate(row_values):
                abs_col = self.start_col + c_idx
                yield abs_row, abs_col, value


def normalize_label(value: Any) -> str:
    if value is None:
        return ""
    label = str(value).strip().lower()
    label = re.sub(r"[^a-z0-9]+", " ", label)
    return label.strip()


def is_blank(value: Any) -> bool:
    if value is None:
        return True
    if isinstance(value, str) and value.strip() == "":
        return True
    return False


def as_float(value: Any) -> Optional[float]:
    if value is None:
        return None
    if isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        if isinstance(value, float) and math.isnan(value):
            return None
        return float(value)
    if isinstance(value, str):
        cleaned = value.strip().replace(",", "")
        if cleaned == "":
            return None
        try:
            return float(cleaned)
        except ValueError:
            return None
    return None


def as_int(value: Any) -> Optional[int]:
    number = as_float(value)
    if number is None:
        return None
    return int(round(number))


def safe_subtract(left: Any, right: Any) -> Optional[float]:
    left_num = as_float(left)
    right_num = as_float(right)
    if left_num is None or right_num is None:
        return None
    return left_num - right_num


def round_for_fingerprint(value: Any) -> Any:
    number = as_float(value)
    if number is None:
        return value
    return round(number, 10)


def parse_labels_from_filename(file_path: Path) -> Optional[FileLabels]:
    stem = file_path.stem
    parts = [part.strip() for part in stem.split(" - ")]
    if len(parts) < 3:
        return None

    ticker = re.sub(r"[^A-Za-z0-9]", "", parts[1]).upper()
    if not ticker:
        return None

    period_token = parts[2].split("_")[0].strip()
    match = re.fullmatch(r"(Early|Mid|Late)([A-Za-z]{3,9})(\d{4})", period_token, flags=re.IGNORECASE)
    if not match:
        return None

    period_prefix_raw, month_text_raw, year_text = match.groups()
    period_prefix = period_prefix_raw.capitalize()
    period_prefix_key = period_prefix.lower()
    month_abbrev = month_text_raw[:3].lower()
    month_num = MONTH_BY_ABBREV.get(month_abbrev)
    if month_num is None:
        return None

    day = DAY_BY_PERIOD_PREFIX[period_prefix_key]
    year = int(year_text)
    model_date = date(year, month_num, day).isoformat()
    month_period = month_abbrev.capitalize()
    model_period = f"{period_prefix}{month_period}_{year_text}"
    model = f"{ticker}_{model_period}"
    return FileLabels(model=model, ticker=ticker, model_period=model_period, model_date=model_date)


def find_anchor(snapshot: SheetSnapshot, anchor_text: str = "max") -> Optional[Tuple[int, int]]:
    target = normalize_label(anchor_text)
    for abs_row, abs_col, value in snapshot.iter_cells():
        if normalize_label(value) == target:
            return abs_row, abs_col
    return None


def build_header_map(snapshot: SheetSnapshot, header_row: int) -> Dict[str, int]:
    header_map: Dict[str, int] = {}
    for idx, value in enumerate(snapshot.row_values(header_row)):
        key = normalize_label(value)
        if key:
            col = snapshot.start_col + idx
            header_map.setdefault(key, col)
    return header_map


def find_col_by_aliases(header_map: Dict[str, int], aliases: Sequence[str], default: Optional[int] = None) -> Optional[int]:
    alias_norm = [normalize_label(alias) for alias in aliases]
    for alias in alias_norm:
        if alias in header_map:
            return header_map[alias]
    for key, col in header_map.items():
        for alias in alias_norm:
            if alias and (alias in key or key in alias):
                return col
    return default


def collect_candidate_rows(
    snapshot: SheetSnapshot,
    header_row: int,
    primary_col: Optional[int],
    row_limit: int = N_QUARTERS,
) -> List[int]:
    rows: List[int] = []
    blank_streak = 0

    max_scan_rows = max(row_limit * 4, row_limit + 2)
    for abs_row in range(header_row + 1, header_row + 1 + max_scan_rows):
        if abs_row > snapshot.max_row:
            break
        primary_value = snapshot.get(abs_row, primary_col) if primary_col else None
        if primary_col is not None and is_blank(primary_value):
            blank_streak += 1
            if blank_streak >= 2 and rows:
                break
            continue
        blank_streak = 0
        rows.append(abs_row)
        if len(rows) >= row_limit:
            break

    if not rows:
        rows = [header_row + idx for idx in range(1, row_limit + 1)]
    return rows[:row_limit]


def find_numeric_history_rows(snapshot: SheetSnapshot, x_col: int, y_col: int, anchor_row: int) -> List[int]:
    numeric_rows = []
    for abs_row in range(snapshot.start_row, snapshot.max_row + 1):
        x_val = as_float(snapshot.get(abs_row, x_col))
        y_val = as_float(snapshot.get(abs_row, y_col))
        if x_val is not None and y_val is not None:
            numeric_rows.append(abs_row)

    if not numeric_rows:
        return []

    blocks: List[List[int]] = []
    current_block: List[int] = [numeric_rows[0]]
    for row in numeric_rows[1:]:
        if row == current_block[-1] + 1:
            current_block.append(row)
        else:
            blocks.append(current_block)
            current_block = [row]
    blocks.append(current_block)

    blocks.sort(
        key=lambda block: (0 if block[0] <= anchor_row <= block[-1] else 1, -len(block), abs(anchor_row - block[-1]))
    )
    return blocks[0]


def build_empirical_avg_formula(
    row: int,
    avg_pen_col: Optional[int],
    quarterly_sales_col: Optional[int],
    reported_sales_col: Optional[int],
) -> Optional[str]:
    if avg_pen_col is not None:
        return f'=IFERROR(AVERAGE(R{row}C{avg_pen_col}:R{row}C{avg_pen_col}),"")'
    if quarterly_sales_col is not None and reported_sales_col is not None:
        return (
            f'=IFERROR('
            f'AVERAGE(R{row}C{reported_sales_col}:R{row}C{reported_sales_col})/'
            f'AVERAGE(R{row}C{quarterly_sales_col}:R{row}C{quarterly_sales_col}),'
            f'"")'
        )
    return None


def safe_close_source_workbook(wb: xw.Book) -> None:
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
        wb.close()
    except Exception as exc:
        print(f"  warning: workbook close fallback failed: {exc}")


def extract_empirical_rows(sheet: xw.Sheet, wb: xw.Book, labels: FileLabels, source_file: str) -> List[Dict[str, Any]]:
    snapshot = SheetSnapshot(sheet)
    anchor = find_anchor(snapshot, "max")
    if anchor is None:
        print("  skipped empirical extraction: 'max' anchor not found")
        return []

    anchor_row, anchor_col = anchor
    header_map = build_header_map(snapshot, anchor_row)

    num_q_col = find_col_by_aliases(header_map, ["num quarters used", "num_quarters_used", "quarters used", "num quarters"])
    last_q_col = find_col_by_aliases(header_map, ["last quarter used", "last_quarter_used", "last quarter"])
    forecast_col = find_col_by_aliases(
        header_map,
        ["estimated total sold", "forecast value", "forecast", "tot fcst", "tot fcst w/o sa"],
    )
    actual_col = find_col_by_aliases(header_map, ["reported sales", "actual value", "actual sales", "actual"])
    min_col = find_col_by_aliases(header_map, ["min", "forecast min"], default=anchor_col + 1)
    avg_pen_col = find_col_by_aliases(header_map, ["avg penetration pct", "avg penetration", "average penetration", "penetration"])
    quarterly_sales_col = find_col_by_aliases(header_map, ["quarterly sales", "quarterly_sales"])
    reported_sales_col = find_col_by_aliases(header_map, ["reported sales", "reported_sales"], default=actual_col)
    growth_col = find_col_by_aliases(header_map, ["growth rate pct", "growth rate", "growth_rate_pct"])
    captured_col = find_col_by_aliases(
        header_map,
        ["sales captured in db pct", "sales_captured_in_db_pct", "captured in db", "db capture pct"],
    )

    row_candidates = collect_candidate_rows(snapshot, anchor_row, primary_col=num_q_col, row_limit=N_QUARTERS)

    avg_values: Dict[int, Any] = {}
    scratch_col = snapshot.max_col + 2
    formula_rows: List[int] = []
    for row in row_candidates:
        formula = build_empirical_avg_formula(row, avg_pen_col, quarterly_sales_col, reported_sales_col)
        if formula:
            sheet.range((row, scratch_col)).formula2 = formula
            formula_rows.append(row)
    if formula_rows:
        wb.app.calculate()
        for row in formula_rows:
            avg_values[row] = sheet.range((row, scratch_col)).value

    rows: List[Dict[str, Any]] = []
    for idx, row in enumerate(row_candidates, start=1):
        num_quarters_used = as_int(snapshot.get(row, num_q_col)) if num_q_col is not None else idx
        if num_quarters_used is None:
            num_quarters_used = idx

        forecast_max = snapshot.get(row, anchor_col)
        forecast_min = snapshot.get(row, min_col) if min_col is not None else None
        forecast_value = snapshot.get(row, forecast_col) if forecast_col is not None else None
        actual_value = snapshot.get(row, actual_col) if actual_col is not None else None
        reported_sales = snapshot.get(row, reported_sales_col) if reported_sales_col is not None else actual_value

        # If the line looks empty across key fields, skip it.
        if all(is_blank(value) for value in (forecast_max, forecast_min, forecast_value, actual_value, reported_sales)):
            continue

        avg_penetration_pct = avg_values.get(row, snapshot.get(row, avg_pen_col) if avg_pen_col is not None else None)
        range_width = safe_subtract(forecast_max, forecast_min)

        row_data: Dict[str, Any] = {
            "model": labels.model,
            "ticker": labels.ticker,
            "model_period": labels.model_period,
            "model_date": labels.model_date,
            "method": "empirical",
            "parameter_name": "avg_penetration_pct",
            "parameter_value": avg_penetration_pct,
            "num_quarters_used": num_quarters_used,
            "last_quarter_used": snapshot.get(row, last_q_col) if last_q_col is not None else None,
            "forecast_value": forecast_value,
            "actual_value": actual_value,
            "forecast_max": forecast_max,
            "forecast_min": forecast_min,
            "range_width": range_width,
            "avg_penetration_pct": avg_penetration_pct,
            "quarterly_sales": snapshot.get(row, quarterly_sales_col) if quarterly_sales_col is not None else None,
            "reported_sales": reported_sales,
            "growth_rate_pct": snapshot.get(row, growth_col) if growth_col is not None else None,
            "sales_captured_in_db_pct": snapshot.get(row, captured_col) if captured_col is not None else None,
            "source_file": source_file,
        }
        rows.append(row_data)
    return rows


def extract_regression_rows(sheet: xw.Sheet, wb: xw.Book, labels: FileLabels, source_file: str) -> List[Dict[str, Any]]:
    snapshot = SheetSnapshot(sheet)
    anchor = find_anchor(snapshot, "max")
    if anchor is None:
        print("  skipped regression extraction: 'max' anchor not found")
        return []

    anchor_row, anchor_col = anchor
    header_map = build_header_map(snapshot, anchor_row)

    num_q_col = find_col_by_aliases(header_map, ["num quarters used", "num_quarters_used", "quarters used", "num quarters"])
    forecast_col = find_col_by_aliases(
        header_map,
        ["tot fcst w/o sa", "tot fcst without sa", "forecast total without sa", "forecast value", "forecast"],
    )
    actual_col = find_col_by_aliases(header_map, ["actual value", "actual sales", "reported sales", "actual"])
    min_col = find_col_by_aliases(header_map, ["min", "forecast min"], default=anchor_col + 1)

    # Required by spec.
    y_col = anchor_col - 7
    x_col = anchor_col - 11
    history_rows = find_numeric_history_rows(snapshot, x_col, y_col, anchor_row)

    row_candidates = collect_candidate_rows(snapshot, anchor_row, primary_col=num_q_col, row_limit=N_QUARTERS)

    scratch_col = snapshot.max_col + 2
    formula_specs: List[Tuple[int, int, int, int]] = []  # (row, n, start_row, end_row)
    for idx, row in enumerate(row_candidates, start=1):
        sheet_num_q = as_int(snapshot.get(row, num_q_col)) if num_q_col is not None else None
        n = sheet_num_q if sheet_num_q is not None else idx
        if not history_rows or len(history_rows) < 2:
            continue
        n = max(2, min(n, len(history_rows)))
        start_row = history_rows[-n]
        end_row = history_rows[-1]
        intercept_formula = (
            f'=IFERROR(INTERCEPT(R{start_row}C{y_col}:R{end_row}C{y_col},'
            f'R{start_row}C{x_col}:R{end_row}C{x_col}),"")'
        )
        slope_formula = (
            f'=IFERROR(SLOPE(R{start_row}C{y_col}:R{end_row}C{y_col},'
            f'R{start_row}C{x_col}:R{end_row}C{x_col}),"")'
        )
        sheet.range((row, scratch_col)).formula2 = intercept_formula
        sheet.range((row, scratch_col + 1)).formula2 = slope_formula
        formula_specs.append((row, n, start_row, end_row))

    if formula_specs:
        wb.app.calculate()

    rows: List[Dict[str, Any]] = []
    previous_fingerprint: Optional[Tuple[Any, ...]] = None

    for idx, row in enumerate(row_candidates, start=1):
        num_quarters_used = as_int(snapshot.get(row, num_q_col)) if num_q_col is not None else idx
        if num_quarters_used is None:
            num_quarters_used = idx

        intercept = sheet.range((row, scratch_col)).value if formula_specs else None
        slope = sheet.range((row, scratch_col + 1)).value if formula_specs else None

        forecast_value = snapshot.get(row, forecast_col) if forecast_col is not None else None

        if is_blank(forecast_value) and as_float(intercept) is not None and as_float(slope) is not None and history_rows:
            last_history_row = history_rows[-1]
            x_next = as_float(snapshot.get(last_history_row + 1, x_col))
            if x_next is None:
                x_last = as_float(snapshot.get(last_history_row, x_col))
                if x_last is not None:
                    x_next = x_last + 1
            if x_next is not None:
                forecast_value = as_float(intercept) + as_float(slope) * x_next

        forecast_max = snapshot.get(row, anchor_col)
        forecast_min = snapshot.get(row, min_col) if min_col is not None else None
        actual_value = snapshot.get(row, actual_col) if actual_col is not None else ""
        range_width = safe_subtract(forecast_max, forecast_min)

        # Skip empty lines.
        if all(is_blank(value) for value in (forecast_value, forecast_max, forecast_min, intercept, slope)):
            continue

        row_data: Dict[str, Any] = {
            "model": labels.model,
            "ticker": labels.ticker,
            "model_period": labels.model_period,
            "model_date": labels.model_date,
            "method": "regression",
            "parameter_name": "num_quarters_used",
            "parameter_value": num_quarters_used,
            "num_quarters_used": num_quarters_used,
            "forecast_value": forecast_value,
            "actual_value": actual_value if actual_value is not None else "",
            "forecast_max": forecast_max,
            "forecast_min": forecast_min,
            "range_width": range_width,
            "intercept": intercept,
            "slope": slope,
            "source_file": source_file,
        }

        fingerprint = (
            row_data["num_quarters_used"],
            round_for_fingerprint(row_data["forecast_value"]),
            round_for_fingerprint(row_data["forecast_max"]),
            round_for_fingerprint(row_data["forecast_min"]),
            round_for_fingerprint(row_data["intercept"]),
            round_for_fingerprint(row_data["slope"]),
        )
        if previous_fingerprint is not None and fingerprint == previous_fingerprint:
            continue
        previous_fingerprint = fingerprint
        rows.append(row_data)

    return rows


def build_output_path(input_dir_path: Path, output_dir_path: Path) -> Path:
    base_name = f"{input_dir_path.name}_PARAM"
    primary = output_dir_path / f"{base_name}.xlsx"
    if not primary.exists():
        return primary

    suffix = 1
    while True:
        candidate = output_dir_path / f"{base_name}.{suffix}.xlsx"
        if not candidate.exists():
            return candidate
        suffix += 1


def write_sheet(ws: Any, columns: Sequence[str], rows: Sequence[Dict[str, Any]]) -> None:
    ws.append(list(columns))
    for row in rows:
        ws.append([row.get(col) for col in columns])

    # Formatting requirements.
    for cell in ws[1]:
        cell.font = Font(bold=True)
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(columns))}{max(1, ws.max_row)}"

    for col_idx in range(1, len(columns) + 1):
        max_len = 0
        for row_idx in range(1, ws.max_row + 1):
            cell_value = ws.cell(row=row_idx, column=col_idx).value
            cell_len = len(str(cell_value)) if cell_value is not None else 0
            if cell_len > max_len:
                max_len = cell_len
        ws.column_dimensions[get_column_letter(col_idx)].width = min(max(max_len + 2, 12), 42)


def write_output_workbook(
    output_path: Path, empirical_rows: Sequence[Dict[str, Any]], regression_rows: Sequence[Dict[str, Any]]
) -> None:
    workbook = Workbook()
    default_sheet = workbook.active
    workbook.remove(default_sheet)

    empirical_ws = workbook.create_sheet(EMPIRICAL_OUTPUT_SHEET)
    regression_ws = workbook.create_sheet(REGRESSION_OUTPUT_SHEET)

    write_sheet(empirical_ws, EMPIRICAL_COLUMNS, empirical_rows)
    write_sheet(regression_ws, REGRESSION_COLUMNS, regression_rows)

    workbook.save(output_path)


def should_skip_output_artifact(file_path: Path, input_folder_name: str) -> bool:
    pattern = re.compile(rf"^{re.escape(input_folder_name)}_PARAM(?:\.\d+)?\.xlsx$", flags=re.IGNORECASE)
    return bool(pattern.fullmatch(file_path.name))


def main() -> None:
    input_dir_path = Path(input_dir)
    output_dir_path = Path(output_dir)

    if not input_dir_path.exists() or not input_dir_path.is_dir():
        raise FileNotFoundError(f"Input directory does not exist or is not a folder: {input_dir_path}")

    output_dir_path.mkdir(parents=True, exist_ok=True)
    output_path = build_output_path(input_dir_path, output_dir_path)

    empirical_rows: List[Dict[str, Any]] = []
    regression_rows: List[Dict[str, Any]] = []

    files_processed = 0
    all_entries = sorted(input_dir_path.iterdir(), key=lambda p: p.name.lower())

    app = xw.App(visible=False, add_book=False)
    app.display_alerts = False
    app.screen_updating = False
    try:
        app.calculation = "manual"
    except Exception:
        pass

    try:
        for file_path in all_entries:
            if not file_path.is_file():
                print(f"skipped file: {file_path.name} (not a file)")
                continue
            if file_path.name.startswith("~"):
                print(f"skipped file: {file_path.name} (temp file)")
                continue
            if file_path.suffix.lower() != ".xlsx":
                print(f"skipped file: {file_path.name} (not .xlsx)")
                continue
            if should_skip_output_artifact(file_path, input_dir_path.name):
                print(f"skipped file: {file_path.name} (existing output artifact)")
                continue

            labels = parse_labels_from_filename(file_path)
            if labels is None:
                print(f"skipped file: {file_path.name} (filename pattern not recognized)")
                continue

            print(f"processed file: {file_path.name}")
            wb: Optional[xw.Book] = None
            try:
                wb = app.books.open(str(file_path), update_links=False)
                sheet_names = {sheet.name for sheet in wb.sheets}

                if EMPIRICAL_MODEL_SHEET in sheet_names:
                    empirical_sheet = wb.sheets[EMPIRICAL_MODEL_SHEET]
                    empirical_rows.extend(
                        extract_empirical_rows(
                            empirical_sheet,
                            wb,
                            labels=labels,
                            source_file=file_path.name,
                        )
                    )
                else:
                    print(f"  skipped empirical extraction: missing '{EMPIRICAL_MODEL_SHEET}' sheet")

                if REGRESSION_MODEL_SHEET in sheet_names:
                    regression_sheet = wb.sheets[REGRESSION_MODEL_SHEET]
                    regression_rows.extend(
                        extract_regression_rows(
                            regression_sheet,
                            wb,
                            labels=labels,
                            source_file=file_path.name,
                        )
                    )
                else:
                    print(f"  skipped regression extraction: missing '{REGRESSION_MODEL_SHEET}' sheet")

                files_processed += 1
            except Exception as exc:
                print(f"skipped file: {file_path.name} (processing error: {exc})")
            finally:
                if wb is not None:
                    safe_close_source_workbook(wb)
    finally:
        try:
            app.quit()
        except Exception:
            pass

    write_output_workbook(output_path, empirical_rows=empirical_rows, regression_rows=regression_rows)

    print(f"output path: {output_path}")
    print(f"number of files processed: {files_processed}")
    print(f"number of empirical rows: {len(empirical_rows)}")
    print(f"number of regression rows: {len(regression_rows)}")


if __name__ == "__main__":
    main()
