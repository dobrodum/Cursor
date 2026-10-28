#!/usr/bin/env python3
"""
Extract empirical/regression parameter candidates from Excel model workbooks.

The script reads every .xlsx file in ``input_dir`` and writes one aggregated
output workbook in ``output_dir`` with two sheets:
  - empirical_candidates
  - regression_candidates
"""

from __future__ import annotations

from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path
import re
from typing import Any

import xlwings as xw
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter


# ---------------------------------------------------------------------------
# Configure these two paths for your environment.
# ---------------------------------------------------------------------------
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

MODEL_DAY_LOOKUP = {
    "early": 5,
    "mid": 15,
    "late": 25,
}


@dataclass(frozen=True)
class FileLabel:
    model: str
    ticker: str
    model_period: str
    model_date: str


class SheetSnapshot:
    """Fast in-memory view of a sheet's used range to reduce COM calls."""

    def __init__(self, sheet: xw.main.Sheet) -> None:
        used = sheet.used_range
        self.start_row = used.row
        self.start_col = used.column
        self.row_count = used.rows.count
        self.col_count = used.columns.count
        self.end_row = self.start_row + self.row_count - 1
        self.end_col = self.start_col + self.col_count - 1
        self.matrix = self._normalize_2d(used.value, self.row_count, self.col_count)

    @staticmethod
    def _normalize_2d(raw: Any, rows: int, cols: int) -> list[list[Any]]:
        if rows <= 0 or cols <= 0:
            return []

        if rows == 1 and cols == 1:
            return [[raw]]

        if rows == 1:
            row = raw if isinstance(raw, list) else [raw]
            row = list(row)[:cols] + [None] * max(0, cols - len(row))
            return [row]

        if cols == 1:
            if isinstance(raw, list):
                col_vals = raw[:rows] + [None] * max(0, rows - len(raw))
            else:
                col_vals = [raw] + [None] * (rows - 1)
            return [[v] for v in col_vals]

        if not isinstance(raw, list):
            return [[None for _ in range(cols)] for _ in range(rows)]

        normalized: list[list[Any]] = []
        for i in range(rows):
            if i < len(raw) and isinstance(raw[i], list):
                row = list(raw[i])[:cols] + [None] * max(0, cols - len(raw[i]))
            elif i < len(raw):
                row = [raw[i]] + [None] * (cols - 1)
            else:
                row = [None] * cols
            normalized.append(row)
        return normalized

    def get(self, row: int, col: int) -> Any:
        if row < self.start_row or row > self.end_row:
            return None
        if col < self.start_col or col > self.end_col:
            return None
        return self.matrix[row - self.start_row][col - self.start_col]

    def row_values(self, row: int) -> list[Any]:
        if row < self.start_row or row > self.end_row:
            return []
        return self.matrix[row - self.start_row]


def parse_file_label(file_name: str) -> FileLabel | None:
    """Parse ticker/model period/date from names like:
       MedMiner_Model - AORT - MidJan2026_Send.xlsx
    """
    stem = Path(file_name).stem
    parts = [p.strip() for p in stem.split(" - ")]
    if len(parts) < 3:
        return None

    ticker = parts[1]
    period_chunk = parts[2]
    match = re.search(r"(Early|Mid|Late)([A-Za-z]+)(\d{4})", period_chunk, flags=re.IGNORECASE)
    if not match:
        return None

    timing = match.group(1).title()
    month_token = match.group(2).lower()[:3]
    year = int(match.group(3))

    month = MONTH_LOOKUP.get(month_token)
    if month is None:
        return None

    day = MODEL_DAY_LOOKUP[timing.lower()]
    model_period = f"{timing}{datetime(year, month, 1).strftime('%b')}_{year}"
    model_date = date(year, month, day).isoformat()
    model = f"{ticker}_{model_period}"

    return FileLabel(
        model=model,
        ticker=ticker,
        model_period=model_period,
        model_date=model_date,
    )


def normalize_header(value: Any) -> str:
    if value is None:
        return ""
    text = str(value).strip().lower()
    text = text.replace("%", " pct ")
    text = re.sub(r"[^a-z0-9]+", " ", text)
    return re.sub(r"\s+", " ", text).strip()


def build_header_map(snapshot: SheetSnapshot, header_row: int) -> dict[str, int]:
    header_map: dict[str, int] = {}
    for idx, value in enumerate(snapshot.row_values(header_row)):
        if isinstance(value, str):
            normalized = normalize_header(value)
            if normalized and normalized not in header_map:
                header_map[normalized] = snapshot.start_col + idx
    return header_map


def pick_header_column(header_map: dict[str, int], *patterns: str) -> int | None:
    for pattern in patterns:
        regex = re.compile(pattern)
        for header, col in header_map.items():
            if regex.search(header):
                return col
    return None


def find_anchor(snapshot: SheetSnapshot, anchor_text: str = "max") -> tuple[int, int] | None:
    needle = anchor_text.strip().lower()
    for r_idx, row in enumerate(snapshot.matrix):
        for c_idx, value in enumerate(row):
            if isinstance(value, str) and value.strip().lower() == needle:
                return snapshot.start_row + r_idx, snapshot.start_col + c_idx
    return None


def has_content(value: Any) -> bool:
    if value is None:
        return False
    if isinstance(value, str):
        return bool(value.strip())
    return True


def to_float(value: Any) -> float | None:
    if value is None:
        return None
    if isinstance(value, (int, float)):
        return float(value)
    if isinstance(value, str):
        text = value.strip().replace(",", "")
        if not text:
            return None
        is_pct = text.endswith("%")
        if is_pct:
            text = text[:-1].strip()
        try:
            parsed = float(text)
            return parsed / 100.0 if is_pct else parsed
        except ValueError:
            return None
    return None


def to_int(value: Any) -> int | None:
    parsed = to_float(value)
    if parsed is None:
        return None
    return int(round(parsed))


def collect_candidate_rows(
    snapshot: SheetSnapshot,
    start_row: int,
    key_columns: list[int | None],
    max_rows: int,
) -> list[int]:
    rows: list[int] = []
    blank_streak = 0

    for row in range(start_row, snapshot.end_row + 1):
        row_has_data = False
        for col in key_columns:
            if col is None:
                continue
            if has_content(snapshot.get(row, col)):
                row_has_data = True
                break

        if row_has_data:
            rows.append(row)
            blank_streak = 0
            if len(rows) >= max_rows:
                break
        elif rows:
            blank_streak += 1
            if blank_streak >= 3:
                break

    return rows


def collect_xy_history_rows(snapshot: SheetSnapshot, start_row: int, x_col: int, y_col: int) -> list[int]:
    rows: list[int] = []
    blank_streak = 0
    for row in range(start_row, snapshot.end_row + 1):
        x_val = to_float(snapshot.get(row, x_col))
        y_val = to_float(snapshot.get(row, y_col))
        if x_val is not None and y_val is not None:
            rows.append(row)
            blank_streak = 0
        elif rows:
            blank_streak += 1
            if blank_streak >= 5:
                break
    return rows


def set_formula2(cell: xw.main.Range, formula: str) -> None:
    try:
        cell.formula2 = formula
    except Exception:
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
        wb.api.Close(SaveChanges=False)
        return
    except Exception:
        pass

    try:
        wb.close()
    except Exception:
        pass


def get_output_path(input_folder: Path, output_folder: Path) -> Path:
    output_folder.mkdir(parents=True, exist_ok=True)
    base_name = f"{input_folder.name}_PARAM"
    candidate = output_folder / f"{base_name}.xlsx"
    if not candidate.exists():
        return candidate

    idx = 1
    while True:
        candidate = output_folder / f"{base_name}.{idx}.xlsx"
        if not candidate.exists():
            return candidate
        idx += 1


def extract_empirical_rows(wb: xw.main.Book, labels: FileLabel, source_file: str) -> list[dict[str, Any]]:
    if "Empirical Model" not in [sheet.name for sheet in wb.sheets]:
        return []

    sheet = wb.sheets["Empirical Model"]
    snapshot = SheetSnapshot(sheet)
    anchor = find_anchor(snapshot, "max")
    if anchor is None:
        return []

    anchor_row, anchor_col = anchor
    header_map = build_header_map(snapshot, anchor_row)

    max_col = anchor_col
    min_col = pick_header_column(header_map, r"\bmin\b")
    if min_col is None:
        min_col = anchor_col + 1 if anchor_col + 1 <= snapshot.end_col else None

    num_quarters_col = pick_header_column(header_map, r"num.*quarter", r"quarters? used")
    last_quarter_col = pick_header_column(header_map, r"last.*quarter")
    forecast_value_col = pick_header_column(
        header_map,
        r"estimated total sold",
        r"tot.*fcst",
        r"forecast",
    )
    avg_penetration_col = pick_header_column(header_map, r"avg.*penetration", r"average.*penetration")
    reported_sales_col = pick_header_column(header_map, r"reported sales")
    quarterly_sales_col = pick_header_column(header_map, r"quarterly sales")
    growth_rate_col = pick_header_column(header_map, r"growth.*rate")
    sales_capture_col = pick_header_column(header_map, r"sales.*captured.*db", r"captured.*db")

    candidate_rows = collect_candidate_rows(
        snapshot=snapshot,
        start_row=anchor_row + 1,
        key_columns=[num_quarters_col, forecast_value_col, max_col, min_col],
        max_rows=N_QUARTERS,
    )
    if not candidate_rows:
        return []

    avg_pen_helper_col = snapshot.end_col + 2
    needs_avg_formula = avg_penetration_col is None
    if needs_avg_formula:
        left_col = max(snapshot.start_col, anchor_col - 12)
        right_col = max(left_col, anchor_col - 3)
        for row in candidate_rows:
            formula = f"=AVERAGE(R{row}C{left_col}:R{row}C{right_col})"
            set_formula2(sheet.range((row, avg_pen_helper_col)), formula)
        wb.app.calculate()

    results: list[dict[str, Any]] = []
    for index, row in enumerate(candidate_rows, start=1):
        num_quarters = to_int(snapshot.get(row, num_quarters_col)) if num_quarters_col else index
        if not num_quarters:
            num_quarters = index

        forecast_value = (
            to_float(snapshot.get(row, forecast_value_col))
            if forecast_value_col is not None
            else None
        )
        forecast_max = to_float(snapshot.get(row, max_col))
        forecast_min = to_float(snapshot.get(row, min_col)) if min_col else None
        range_width = (
            forecast_max - forecast_min
            if forecast_max is not None and forecast_min is not None
            else None
        )

        reported_sales = (
            to_float(snapshot.get(row, reported_sales_col))
            if reported_sales_col is not None
            else None
        )

        if avg_penetration_col is not None:
            avg_penetration = to_float(snapshot.get(row, avg_penetration_col))
        else:
            avg_penetration = to_float(sheet.range((row, avg_pen_helper_col)).value)

        if (
            forecast_value is None
            and forecast_max is None
            and forecast_min is None
            and avg_penetration is None
        ):
            continue

        results.append(
            {
                "model": labels.model,
                "ticker": labels.ticker,
                "model_period": labels.model_period,
                "model_date": labels.model_date,
                "method": "empirical",
                "parameter_name": "avg_penetration_pct",
                "parameter_value": avg_penetration,
                "num_quarters_used": num_quarters,
                "last_quarter_used": snapshot.get(row, last_quarter_col) if last_quarter_col else None,
                "forecast_value": forecast_value,
                "actual_value": reported_sales,
                "forecast_max": forecast_max,
                "forecast_min": forecast_min,
                "range_width": range_width,
                "avg_penetration_pct": avg_penetration,
                "quarterly_sales": (
                    to_float(snapshot.get(row, quarterly_sales_col))
                    if quarterly_sales_col is not None
                    else None
                ),
                "reported_sales": reported_sales,
                "growth_rate_pct": (
                    to_float(snapshot.get(row, growth_rate_col)) if growth_rate_col is not None else None
                ),
                "sales_captured_in_db_pct": (
                    to_float(snapshot.get(row, sales_capture_col))
                    if sales_capture_col is not None
                    else None
                ),
                "source_file": source_file,
            }
        )

    return results


def extract_regression_rows(wb: xw.main.Book, labels: FileLabel, source_file: str) -> list[dict[str, Any]]:
    if "Regression Model" not in [sheet.name for sheet in wb.sheets]:
        return []

    sheet = wb.sheets["Regression Model"]
    snapshot = SheetSnapshot(sheet)
    anchor = find_anchor(snapshot, "max")
    if anchor is None:
        return []

    anchor_row, anchor_col = anchor
    header_map = build_header_map(snapshot, anchor_row)

    # Required anchor-based source columns from the existing logic.
    y_col = anchor_col - 7
    x_col = anchor_col - 11

    max_col = anchor_col
    min_col = pick_header_column(header_map, r"\bmin\b")
    if min_col is None:
        min_col = anchor_col + 1 if anchor_col + 1 <= snapshot.end_col else None

    num_quarters_col = pick_header_column(header_map, r"num.*quarter", r"quarters? used")
    forecast_value_col = pick_header_column(
        header_map,
        r"tot.*fcst.*without.*sa",
        r"tot.*fcst.*w o.*sa",
        r"fcst.*w o.*sa",
        r"forecast",
    )
    actual_value_col = pick_header_column(header_map, r"\bactual\b", r"reported sales")

    candidate_rows = collect_candidate_rows(
        snapshot=snapshot,
        start_row=anchor_row + 1,
        key_columns=[num_quarters_col, forecast_value_col, max_col, min_col],
        max_rows=N_QUARTERS,
    )
    if not candidate_rows:
        return []

    history_rows = collect_xy_history_rows(snapshot, anchor_row + 1, x_col=x_col, y_col=y_col)

    intercept_helper_col = snapshot.end_col + 2
    slope_helper_col = snapshot.end_col + 3
    points_used_by_row: dict[int, int] = {}

    if len(history_rows) >= 2:
        for index, row in enumerate(candidate_rows, start=1):
            num_quarters = (
                to_int(snapshot.get(row, num_quarters_col)) if num_quarters_col is not None else index
            )
            if not num_quarters or num_quarters < 2:
                num_quarters = max(2, index)

            points_to_use = min(num_quarters, len(history_rows))
            start_hist_row = history_rows[-points_to_use]
            end_hist_row = history_rows[-1]

            intercept_formula = (
                f"=INTERCEPT(R{start_hist_row}C{y_col}:R{end_hist_row}C{y_col},"
                f"R{start_hist_row}C{x_col}:R{end_hist_row}C{x_col})"
            )
            slope_formula = (
                f"=SLOPE(R{start_hist_row}C{y_col}:R{end_hist_row}C{y_col},"
                f"R{start_hist_row}C{x_col}:R{end_hist_row}C{x_col})"
            )
            set_formula2(sheet.range((row, intercept_helper_col)), intercept_formula)
            set_formula2(sheet.range((row, slope_helper_col)), slope_formula)
            points_used_by_row[row] = points_to_use

        wb.app.calculate()

    results: list[dict[str, Any]] = []
    previous_signature: tuple[Any, ...] | None = None

    for index, row in enumerate(candidate_rows, start=1):
        num_quarters = (
            to_int(snapshot.get(row, num_quarters_col)) if num_quarters_col is not None else index
        )
        if not num_quarters:
            num_quarters = index

        forecast_value = (
            to_float(snapshot.get(row, forecast_value_col))
            if forecast_value_col is not None
            else None
        )
        actual_value = (
            to_float(snapshot.get(row, actual_value_col))
            if actual_value_col is not None
            else None
        )
        forecast_max = to_float(snapshot.get(row, max_col))
        forecast_min = to_float(snapshot.get(row, min_col)) if min_col else None
        range_width = (
            forecast_max - forecast_min
            if forecast_max is not None and forecast_min is not None
            else None
        )

        intercept = (
            to_float(sheet.range((row, intercept_helper_col)).value)
            if row in points_used_by_row
            else None
        )
        slope = (
            to_float(sheet.range((row, slope_helper_col)).value)
            if row in points_used_by_row
            else None
        )

        if (
            forecast_value is None
            and forecast_max is None
            and forecast_min is None
            and intercept is None
            and slope is None
        ):
            continue

        signature = (
            num_quarters,
            round(forecast_value, 10) if forecast_value is not None else None,
            round(forecast_max, 10) if forecast_max is not None else None,
            round(forecast_min, 10) if forecast_min is not None else None,
            round(intercept, 10) if intercept is not None else None,
            round(slope, 10) if slope is not None else None,
        )
        if previous_signature is not None and signature == previous_signature:
            continue
        previous_signature = signature

        results.append(
            {
                "model": labels.model,
                "ticker": labels.ticker,
                "model_period": labels.model_period,
                "model_date": labels.model_date,
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
        )

    return results


def write_sheet(
    wb: Workbook,
    sheet_name: str,
    columns: list[str],
    rows: list[dict[str, Any]],
) -> None:
    ws = wb.create_sheet(sheet_name)
    ws.append(columns)

    for col_idx, header in enumerate(columns, start=1):
        ws.cell(row=1, column=col_idx, value=header).font = Font(bold=True)

    for row in rows:
        ws.append([row.get(col) for col in columns])

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(columns))}{max(1, ws.max_row)}"

    for col_idx, header in enumerate(columns, start=1):
        max_len = len(header)
        for row_idx in range(2, ws.max_row + 1):
            value = ws.cell(row=row_idx, column=col_idx).value
            if value is None:
                continue
            max_len = max(max_len, len(str(value)))
        ws.column_dimensions[get_column_letter(col_idx)].width = min(max(12, max_len + 2), 42)


def write_output_workbook(
    destination: Path,
    empirical_rows: list[dict[str, Any]],
    regression_rows: list[dict[str, Any]],
) -> None:
    wb = Workbook()
    default_sheet = wb.active
    wb.remove(default_sheet)

    write_sheet(wb, "empirical_candidates", EMPIRICAL_COLUMNS, empirical_rows)
    write_sheet(wb, "regression_candidates", REGRESSION_COLUMNS, regression_rows)
    wb.save(destination)


def iter_input_files(folder: Path) -> tuple[list[Path], int]:
    processable: list[Path] = []
    skipped_count = 0

    if not folder.exists():
        raise FileNotFoundError(f"Input folder does not exist: {folder}")

    for file_path in sorted(folder.iterdir(), key=lambda p: p.name.lower()):
        if not file_path.is_file():
            continue
        if file_path.name.startswith("~"):
            print(f"skipped: {file_path.name} (temporary file)")
            skipped_count += 1
            continue
        if file_path.suffix.lower() != ".xlsx":
            print(f"skipped: {file_path.name} (not .xlsx)")
            skipped_count += 1
            continue
        processable.append(file_path)

    return processable, skipped_count


def main() -> None:
    source_files, skipped_count = iter_input_files(input_dir)
    empirical_rows: list[dict[str, Any]] = []
    regression_rows: list[dict[str, Any]] = []
    processed_count = 0

    app: xw.main.App | None = None
    try:
        app = xw.App(visible=False, add_book=False)
        app.display_alerts = False
        app.screen_updating = False

        for file_path in source_files:
            labels = parse_file_label(file_path.name)
            if labels is None:
                print(f"skipped: {file_path.name} (filename pattern not recognized)")
                skipped_count += 1
                continue

            wb: xw.main.Book | None = None
            try:
                wb = app.books.open(str(file_path), update_links=False)
                empirical_rows.extend(extract_empirical_rows(wb, labels, file_path.name))
                regression_rows.extend(extract_regression_rows(wb, labels, file_path.name))
                processed_count += 1
                print(f"processed: {file_path.name}")
            except Exception as exc:
                print(f"skipped: {file_path.name} (error: {exc})")
                skipped_count += 1
            finally:
                if wb is not None:
                    safe_close_workbook(wb)
    finally:
        if app is not None:
            app.quit()

    out_path = get_output_path(input_dir, output_dir)
    write_output_workbook(out_path, empirical_rows, regression_rows)

    print(f"output_path: {out_path}")
    print(f"files_processed: {processed_count}")
    print(f"files_skipped: {skipped_count}")
    print(f"empirical_rows: {len(empirical_rows)}")
    print(f"regression_rows: {len(regression_rows)}")


if __name__ == "__main__":
    main()
