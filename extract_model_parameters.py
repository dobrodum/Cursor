from __future__ import annotations

from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Sequence, Tuple
import re

import xlwings as xw
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter


# ========= User-configurable paths =========
input_dir = "/workspace/input"
output_dir = "/workspace/output"
# ==========================================


EMPIRICAL_COLUMNS: List[str] = [
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

REGRESSION_COLUMNS: List[str] = [
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


MONTH_TO_NUM = {
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

WINDOW_TO_DAY = {"early": 5, "mid": 15, "late": 25}


@dataclass
class FileMeta:
    model: str
    ticker: str
    model_period: str
    model_date: str
    source_file: str


@dataclass
class SheetSnapshot:
    values: List[List[Any]]
    start_row: int
    start_col: int
    n_rows: int
    n_cols: int
    label_index: Dict[str, List[Tuple[int, int]]]

    @property
    def end_row(self) -> int:
        return self.start_row + self.n_rows - 1

    @property
    def end_col(self) -> int:
        return self.start_col + self.n_cols - 1

    def get_value(self, row: int, col: int) -> Any:
        row_idx = row - self.start_row
        col_idx = col - self.start_col
        if row_idx < 0 or col_idx < 0 or row_idx >= self.n_rows or col_idx >= self.n_cols:
            return None
        row_values = self.values[row_idx]
        if col_idx >= len(row_values):
            return None
        return row_values[col_idx]

    @classmethod
    def from_sheet(cls, sheet: xw.Sheet) -> "SheetSnapshot":
        used = sheet.used_range
        start_row = used.row
        start_col = used.column
        values_2d = normalize_2d(used.value)
        if not values_2d:
            return cls(values=[], start_row=start_row, start_col=start_col, n_rows=0, n_cols=0, label_index={})
        n_rows = len(values_2d)
        n_cols = max(len(r) for r in values_2d)
        label_index: Dict[str, List[Tuple[int, int]]] = {}
        for r_offset, row_values in enumerate(values_2d):
            for c_offset, value in enumerate(row_values):
                norm = normalize_text(value)
                if not norm:
                    continue
                abs_row = start_row + r_offset
                abs_col = start_col + c_offset
                label_index.setdefault(norm, []).append((abs_row, abs_col))
        return cls(
            values=values_2d,
            start_row=start_row,
            start_col=start_col,
            n_rows=n_rows,
            n_cols=n_cols,
            label_index=label_index,
        )


def normalize_2d(value: Any) -> List[List[Any]]:
    if value is None:
        return []
    if isinstance(value, list):
        if not value:
            return []
        if isinstance(value[0], list):
            return value
        return [value]
    return [[value]]


def normalize_text(value: Any) -> str:
    if value is None:
        return ""
    return str(value).strip().lower()


def is_number(value: Any) -> bool:
    return isinstance(value, (int, float)) and not isinstance(value, bool)


def to_float(value: Any) -> Optional[float]:
    if value is None:
        return None
    if isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        return float(value)
    if isinstance(value, str):
        txt = value.strip().replace(",", "")
        if not txt:
            return None
        try:
            return float(txt)
        except ValueError:
            return None
    return None


def to_int(value: Any) -> Optional[int]:
    num = to_float(value)
    if num is None:
        return None
    try:
        return int(round(num))
    except (TypeError, ValueError):
        return None


def to_excel_safe(value: Any) -> Any:
    if isinstance(value, datetime):
        return value.strftime("%Y-%m-%d")
    if isinstance(value, date):
        return value.isoformat()
    return value


def parse_file_meta(file_path: Path) -> Optional[FileMeta]:
    stem = file_path.stem
    parts = [p.strip() for p in stem.split(" - ")]
    if len(parts) < 3:
        return None

    ticker = parts[1].strip().upper()
    period_part = parts[2].replace("_Send", "").replace("_send", "").strip()
    match = re.search(r"(Early|Mid|Late)(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)(\d{4})", period_part, re.IGNORECASE)
    if not match:
        return None

    window_raw, month_raw, year_raw = match.groups()
    window = window_raw.capitalize()
    month = month_raw.capitalize()
    year = int(year_raw)

    day = WINDOW_TO_DAY[window.lower()]
    month_num = MONTH_TO_NUM[month.lower()]
    model_date = f"{year:04d}-{month_num:02d}-{day:02d}"
    model_period = f"{window}{month}_{year:04d}"
    model = f"{ticker}_{model_period}"

    return FileMeta(
        model=model,
        ticker=ticker,
        model_period=model_period,
        model_date=model_date,
        source_file=file_path.name,
    )


def build_output_path(input_path: Path, output_path: Path) -> Path:
    output_path.mkdir(parents=True, exist_ok=True)
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


def find_max_anchor(snapshot: SheetSnapshot) -> Optional[Tuple[int, int]]:
    max_cells = snapshot.label_index.get("max", [])
    if not max_cells:
        return None
    for row, col in max_cells:
        if normalize_text(snapshot.get_value(row, col + 1)) == "min":
            return (row, col)
    return max_cells[0]


def find_adjacent_value(snapshot: SheetSnapshot, labels: Sequence[str]) -> Any:
    for label in labels:
        for row, col in snapshot.label_index.get(label, []):
            right = snapshot.get_value(row, col + 1)
            if right not in (None, ""):
                return right
            below = snapshot.get_value(row + 1, col)
            if below not in (None, ""):
                return below
    return None


def growth_rate(first_val: Any, last_val: Any) -> Optional[float]:
    first_num = to_float(first_val)
    last_num = to_float(last_val)
    if first_num is None or last_num is None or first_num == 0:
        return None
    return (last_num - first_num) / first_num


def diff_or_none(a: Any, b: Any) -> Optional[float]:
    a_num = to_float(a)
    b_num = to_float(b)
    if a_num is None or b_num is None:
        return None
    return a_num - b_num


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
    except Exception:
        try:
            wb.close()
        except Exception:
            pass


def read_cell(snapshot: SheetSnapshot, row: int, col: int, fallback: Any = None) -> Any:
    value = snapshot.get_value(row, col)
    return fallback if value in (None, "") else value


def extract_empirical_rows(sheet: xw.Sheet, meta: FileMeta, wb: xw.Book) -> List[Dict[str, Any]]:
    snapshot = SheetSnapshot.from_sheet(sheet)
    anchor = find_max_anchor(snapshot)
    if snapshot.n_rows == 0 or anchor is None:
        return []

    anchor_row, anchor_col = anchor
    n_quarters = 10

    # Offsets are anchored to the located "max" column.
    num_q_col = anchor_col - 5
    forecast_col = anchor_col - 1
    forecast_max_col = anchor_col
    forecast_min_col = anchor_col + 1
    sales_col = anchor_col - 7
    penetration_col = anchor_col - 11
    quarter_label_col = anchor_col - 12
    growth_col = anchor_col - 8
    db_capture_col = anchor_col - 9

    reported_sales_default = find_adjacent_value(
        snapshot,
        labels=("reported sales", "reported_sale", "actual sales", "actual"),
    )

    scratch_col = snapshot.end_col + 3
    row_payloads: List[Dict[str, Any]] = []
    formula_rows: List[int] = []

    for i in range(n_quarters):
        row = anchor_row + 1 + i
        num_quarters_used = to_int(read_cell(snapshot, row, num_q_col, i + 1)) or (i + 1)

        hist_end_row = anchor_row - 1
        hist_start_row = max(snapshot.start_row, hist_end_row - num_quarters_used + 1)
        if hist_start_row > hist_end_row:
            hist_start_row = hist_end_row

        avg_formula = f"=AVERAGE(R{hist_start_row}C{penetration_col}:R{hist_end_row}C{penetration_col})"
        sheet.range((row, scratch_col)).formula2 = avg_formula
        formula_rows.append(row)

        forecast_value = read_cell(snapshot, row, forecast_col)
        forecast_max = read_cell(snapshot, row, forecast_max_col)
        forecast_min = read_cell(snapshot, row, forecast_min_col)
        quarterly_sales = read_cell(snapshot, hist_end_row, sales_col)
        first_quarter_sales = read_cell(snapshot, hist_start_row, sales_col)
        growth_rate_pct = read_cell(snapshot, row, growth_col, growth_rate(first_quarter_sales, quarterly_sales))
        sales_captured_in_db_pct = read_cell(snapshot, row, db_capture_col, read_cell(snapshot, hist_end_row, penetration_col))
        last_quarter_used = read_cell(snapshot, hist_end_row, quarter_label_col)
        reported_sales = read_cell(snapshot, row, forecast_col + 4, reported_sales_default)

        row_payloads.append(
            {
                "row": row,
                "num_quarters_used": num_quarters_used,
                "last_quarter_used": to_excel_safe(last_quarter_used),
                "forecast_value": forecast_value,
                "actual_value": reported_sales,
                "forecast_max": forecast_max,
                "forecast_min": forecast_min,
                "range_width": diff_or_none(forecast_max, forecast_min),
                "quarterly_sales": quarterly_sales,
                "reported_sales": reported_sales,
                "growth_rate_pct": growth_rate_pct,
                "sales_captured_in_db_pct": sales_captured_in_db_pct,
            }
        )

    if formula_rows:
        wb.app.calculate()

    rows: List[Dict[str, Any]] = []
    for payload in row_payloads:
        row = payload["row"]
        avg_penetration_pct = sheet.range((row, scratch_col)).value
        forecast_value = payload["forecast_value"]
        if forecast_value in (None, ""):
            avg_num = to_float(avg_penetration_pct)
            q_num = to_float(payload["quarterly_sales"])
            forecast_value = (avg_num * q_num) if (avg_num is not None and q_num is not None) else None

        forecast_max = payload["forecast_max"]
        forecast_min = payload["forecast_min"]
        if forecast_max in (None, "") and forecast_value is not None:
            forecast_max = forecast_value
        if forecast_min in (None, "") and forecast_value is not None:
            forecast_min = forecast_value

        has_signal = any(
            value not in (None, "")
            for value in (forecast_value, forecast_max, forecast_min, avg_penetration_pct)
        )
        if not has_signal:
            continue

        rows.append(
            {
                "model": meta.model,
                "ticker": meta.ticker,
                "model_period": meta.model_period,
                "model_date": meta.model_date,
                "method": "empirical",
                "parameter_name": "avg_penetration_pct",
                "parameter_value": avg_penetration_pct,
                "num_quarters_used": payload["num_quarters_used"],
                "last_quarter_used": payload["last_quarter_used"],
                "forecast_value": forecast_value,
                "actual_value": payload["actual_value"],
                "forecast_max": forecast_max,
                "forecast_min": forecast_min,
                "range_width": diff_or_none(forecast_max, forecast_min),
                "avg_penetration_pct": avg_penetration_pct,
                "quarterly_sales": payload["quarterly_sales"],
                "reported_sales": payload["reported_sales"],
                "growth_rate_pct": payload["growth_rate_pct"],
                "sales_captured_in_db_pct": payload["sales_captured_in_db_pct"],
                "source_file": meta.source_file,
            }
        )
    return rows


def rows_match_for_dedup(a: Dict[str, Any], b: Dict[str, Any]) -> bool:
    keys = (
        "num_quarters_used",
        "forecast_value",
        "forecast_max",
        "forecast_min",
        "intercept",
        "slope",
    )
    for key in keys:
        av = a.get(key)
        bv = b.get(key)
        if av in (None, "") and bv in (None, ""):
            continue
        av_num = to_float(av)
        bv_num = to_float(bv)
        if av_num is not None and bv_num is not None:
            if abs(av_num - bv_num) > 1e-12:
                return False
        elif av != bv:
            return False
    return True


def extract_regression_rows(sheet: xw.Sheet, meta: FileMeta, wb: xw.Book) -> List[Dict[str, Any]]:
    snapshot = SheetSnapshot.from_sheet(sheet)
    anchor = find_max_anchor(snapshot)
    if snapshot.n_rows == 0 or anchor is None:
        return []

    anchor_row, anchor_col = anchor
    n_quarters = 10

    y_col = anchor_col - 7
    x_col = anchor_col - 11
    num_q_col = anchor_col - 5
    forecast_col = anchor_col - 1
    forecast_max_col = anchor_col
    forecast_min_col = anchor_col + 1

    actual_value_default = find_adjacent_value(
        snapshot,
        labels=("reported sales", "actual sales", "actual"),
    )

    scratch_intercept_col = snapshot.end_col + 3
    scratch_slope_col = snapshot.end_col + 4

    row_payloads: List[Dict[str, Any]] = []
    formula_rows: List[int] = []

    for i in range(n_quarters):
        row = anchor_row + 1 + i
        num_quarters_used = to_int(read_cell(snapshot, row, num_q_col, i + 1)) or (i + 1)

        hist_end_row = anchor_row - 1
        hist_start_row = max(snapshot.start_row, hist_end_row - num_quarters_used + 1)
        if hist_start_row > hist_end_row:
            hist_start_row = hist_end_row

        intercept_formula = (
            f"=INTERCEPT(R{hist_start_row}C{y_col}:R{hist_end_row}C{y_col},"
            f"R{hist_start_row}C{x_col}:R{hist_end_row}C{x_col})"
        )
        slope_formula = (
            f"=SLOPE(R{hist_start_row}C{y_col}:R{hist_end_row}C{y_col},"
            f"R{hist_start_row}C{x_col}:R{hist_end_row}C{x_col})"
        )
        sheet.range((row, scratch_intercept_col)).formula2 = intercept_formula
        sheet.range((row, scratch_slope_col)).formula2 = slope_formula
        formula_rows.append(row)

        row_payloads.append(
            {
                "row": row,
                "num_quarters_used": num_quarters_used,
                "forecast_value": read_cell(snapshot, row, forecast_col),
                "forecast_max": read_cell(snapshot, row, forecast_max_col),
                "forecast_min": read_cell(snapshot, row, forecast_min_col),
            }
        )

    if formula_rows:
        wb.app.calculate()

    rows: List[Dict[str, Any]] = []
    previous_row: Optional[Dict[str, Any]] = None

    for payload in row_payloads:
        row = payload["row"]
        intercept = sheet.range((row, scratch_intercept_col)).value
        slope = sheet.range((row, scratch_slope_col)).value
        forecast_value = payload["forecast_value"]

        if forecast_value in (None, ""):
            x_next = read_cell(snapshot, anchor_row, x_col, read_cell(snapshot, anchor_row - 1, x_col))
            x_next_num = to_float(x_next)
            intercept_num = to_float(intercept)
            slope_num = to_float(slope)
            if x_next_num is not None and intercept_num is not None and slope_num is not None:
                forecast_value = intercept_num + (slope_num * x_next_num)

        forecast_max = payload["forecast_max"] if payload["forecast_max"] not in (None, "") else forecast_value
        forecast_min = payload["forecast_min"] if payload["forecast_min"] not in (None, "") else forecast_value
        range_width = diff_or_none(forecast_max, forecast_min)

        candidate = {
            "model": meta.model,
            "ticker": meta.ticker,
            "model_period": meta.model_period,
            "model_date": meta.model_date,
            "method": "regression",
            "parameter_name": "num_quarters_used",
            "parameter_value": payload["num_quarters_used"],
            "num_quarters_used": payload["num_quarters_used"],
            "forecast_value": forecast_value,
            "actual_value": actual_value_default,
            "forecast_max": forecast_max,
            "forecast_min": forecast_min,
            "range_width": range_width,
            "intercept": intercept,
            "slope": slope,
            "source_file": meta.source_file,
        }

        has_signal = any(
            value not in (None, "")
            for value in (candidate["forecast_value"], candidate["intercept"], candidate["slope"])
        )
        if not has_signal:
            continue

        if previous_row is not None and rows_match_for_dedup(previous_row, candidate):
            continue

        rows.append(candidate)
        previous_row = candidate

    return rows


def write_rows_to_sheet(ws, columns: Sequence[str], rows: Sequence[Dict[str, Any]]) -> None:
    ws.append(list(columns))
    for row in rows:
        ws.append([to_excel_safe(row.get(col)) for col in columns])

    for cell in ws[1]:
        cell.font = Font(bold=True)

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions

    for col_idx, col_name in enumerate(columns, start=1):
        max_len = len(col_name)
        for row_idx in range(2, ws.max_row + 1):
            value = ws.cell(row=row_idx, column=col_idx).value
            value_len = len(str(value)) if value is not None else 0
            if value_len > max_len:
                max_len = value_len
        ws.column_dimensions[get_column_letter(col_idx)].width = min(max(max_len + 2, 12), 48)


def create_output_workbook(
    output_path: Path,
    empirical_rows: Sequence[Dict[str, Any]],
    regression_rows: Sequence[Dict[str, Any]],
) -> None:
    wb_out = Workbook()
    ws_emp = wb_out.active
    ws_emp.title = "empirical_candidates"
    ws_reg = wb_out.create_sheet("regression_candidates")

    write_rows_to_sheet(ws_emp, EMPIRICAL_COLUMNS, empirical_rows)
    write_rows_to_sheet(ws_reg, REGRESSION_COLUMNS, regression_rows)
    wb_out.save(output_path)


def iter_source_files(path: Path) -> Iterable[Path]:
    for file_path in sorted(path.iterdir()):
        if not file_path.is_file():
            continue
        if file_path.name.startswith("~"):
            print(f"skipped file: {file_path.name} (temp file)")
            continue
        if file_path.suffix.lower() != ".xlsx":
            print(f"skipped file: {file_path.name} (not .xlsx)")
            continue
        yield file_path


def main() -> None:
    input_path = Path(input_dir).expanduser().resolve()
    output_path = Path(output_dir).expanduser().resolve()

    if not input_path.exists() or not input_path.is_dir():
        raise FileNotFoundError(f"input_dir does not exist or is not a folder: {input_path}")

    output_file = build_output_path(input_path, output_path)

    empirical_rows: List[Dict[str, Any]] = []
    regression_rows: List[Dict[str, Any]] = []
    files_processed = 0

    app = xw.App(visible=False, add_book=False)
    app.display_alerts = False
    app.screen_updating = False

    try:
        for file_path in iter_source_files(input_path):
            meta = parse_file_meta(file_path)
            if meta is None:
                print(f"skipped file: {file_path.name} (filename parse failed)")
                continue

            print(f"processing file: {file_path.name}")
            wb = None
            try:
                wb = app.books.open(str(file_path), update_links=False)
                sheet_names = {sheet.name for sheet in wb.sheets}

                if "Empirical Model" in sheet_names:
                    empirical_rows.extend(extract_empirical_rows(wb.sheets["Empirical Model"], meta, wb))
                else:
                    print(f"skipped sheet: {file_path.name} -> Empirical Model missing")

                if "Regression Model" in sheet_names:
                    regression_rows.extend(extract_regression_rows(wb.sheets["Regression Model"], meta, wb))
                else:
                    print(f"skipped sheet: {file_path.name} -> Regression Model missing")

                files_processed += 1
            except Exception as exc:
                print(f"skipped file: {file_path.name} (processing error: {exc})")
            finally:
                if wb is not None:
                    safe_close_workbook(wb)
    finally:
        app.quit()

    create_output_workbook(output_file, empirical_rows, regression_rows)

    print(f"output path: {output_file}")
    print(f"files processed: {files_processed}")
    print(f"empirical rows: {len(empirical_rows)}")
    print(f"regression rows: {len(regression_rows)}")


if __name__ == "__main__":
    main()
