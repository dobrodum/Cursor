from __future__ import annotations

import calendar
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

# User-configurable paths.
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

# Anchor-based offsets relative to the cell containing text "max".
EMPIRICAL_OFFSETS = {
    "last_quarter": -12,
    "quarterly_sales": -11,
    "reported_sales": -10,
    "growth_rate_pct": -9,
    "sales_captured_in_db_pct": -8,
    "forecast_value": -2,
    "forecast_min": -1,
    "forecast_max": 0,
}

REGRESSION_OFFSETS = {
    "forecast_total_without_sa": -2,
    "forecast_min": -1,
    "forecast_max": 0,
}

MODEL_DAY_BY_PHASE = {"Early": 5, "Mid": 15, "Late": 25}
MONTH_TO_INT = {abbr: idx for idx, abbr in enumerate(calendar.month_abbr) if abbr}


@dataclass
class FileLabel:
    model: str
    ticker: str
    model_period: str
    model_date: str


@dataclass
class SheetSnapshot:
    start_row: int
    start_col: int
    values: List[List[Any]]

    @property
    def row_count(self) -> int:
        return len(self.values)

    @property
    def col_count(self) -> int:
        return max((len(row) for row in self.values), default=0)

    @property
    def end_row(self) -> int:
        return self.start_row + self.row_count - 1

    @property
    def end_col(self) -> int:
        return self.start_col + self.col_count - 1

    def get(self, row: int, col: int) -> Any:
        r_idx = row - self.start_row
        c_idx = col - self.start_col
        if r_idx < 0 or c_idx < 0:
            return None
        if r_idx >= self.row_count:
            return None
        row_values = self.values[r_idx]
        if c_idx >= len(row_values):
            return None
        return row_values[c_idx]


def normalize_2d(values: Any) -> List[List[Any]]:
    if values is None:
        return []
    if not isinstance(values, list):
        return [[values]]
    if not values:
        return []
    if isinstance(values[0], list):
        return values
    return [values]


def to_snapshot(sheet: xw.Sheet) -> SheetSnapshot:
    used = sheet.used_range
    return SheetSnapshot(
        start_row=used.row,
        start_col=used.column,
        values=normalize_2d(used.value),
    )


def to_float(value: Any) -> Optional[float]:
    if value is None:
        return None
    if isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        if isinstance(value, float) and math.isnan(value):
            return None
        return float(value)
    if isinstance(value, str):
        text = value.strip().replace(",", "")
        if not text:
            return None
        try:
            return float(text)
        except ValueError:
            return None
    return None


def almost_equal(left: Any, right: Any, tol: float = 1e-9) -> bool:
    left_num = to_float(left)
    right_num = to_float(right)
    if left_num is None and right_num is None:
        return True
    if left_num is None or right_num is None:
        return False
    return abs(left_num - right_num) <= tol


def find_anchor(snapshot: SheetSnapshot, anchor_text: str = "max") -> Optional[Tuple[int, int]]:
    target = anchor_text.strip().lower()
    for r_idx, row_vals in enumerate(snapshot.values):
        for c_idx, raw in enumerate(row_vals):
            if isinstance(raw, str) and raw.strip().lower() == target:
                return snapshot.start_row + r_idx, snapshot.start_col + c_idx
    return None


def set_formula2_r1c1(cell: xw.Range, formula_r1c1: str) -> None:
    try:
        cell.formula2 = formula_r1c1
        return
    except Exception:
        pass
    try:
        cell.api.Formula2R1C1 = formula_r1c1
        return
    except Exception:
        cell.formula = formula_r1c1


def safe_close_source_workbook(wb: xw.Book) -> None:
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


def get_sheet_if_present(wb: xw.Book, sheet_name: str) -> Optional[xw.Sheet]:
    for sheet in wb.sheets:
        if sheet.name == sheet_name:
            return sheet
    return None


def parse_file_label(file_name: str) -> Optional[FileLabel]:
    stem = Path(file_name).stem
    parts = [part.strip() for part in stem.split("-")]
    if len(parts) < 3:
        return None

    ticker = parts[1].strip().upper().replace(" ", "")
    period_source = parts[2]
    period_match = re.search(r"(Early|Mid|Late)\s*([A-Za-z]{3})\s*(\d{4})", period_source, flags=re.I)
    if not period_match:
        return None

    phase_raw, month_raw, year_raw = period_match.groups()
    phase = phase_raw.capitalize()
    month_abbr = month_raw.capitalize()
    year_int = int(year_raw)
    month_num = MONTH_TO_INT.get(month_abbr)
    if month_num is None:
        return None

    day = MODEL_DAY_BY_PHASE[phase]
    model_period = f"{phase}{month_abbr}_{year_int}"
    model_date = date(year_int, month_num, day).isoformat()
    model = f"{ticker}_{model_period}"
    return FileLabel(
        model=model,
        ticker=ticker,
        model_period=model_period,
        model_date=model_date,
    )


def derive_empirical_rows(
    wb: xw.Book,
    sheet: xw.Sheet,
    snapshot: SheetSnapshot,
    label: FileLabel,
    source_file: str,
) -> List[Dict[str, Any]]:
    anchor = find_anchor(snapshot, "max")
    if anchor is None:
        print(f"Skipped empirical extraction in {source_file}: anchor 'max' not found")
        return []

    anchor_row, anchor_col = anchor
    quarter_col = anchor_col + EMPIRICAL_OFFSETS["last_quarter"]
    quarterly_sales_col = anchor_col + EMPIRICAL_OFFSETS["quarterly_sales"]
    reported_sales_col = anchor_col + EMPIRICAL_OFFSETS["reported_sales"]
    growth_col = anchor_col + EMPIRICAL_OFFSETS["growth_rate_pct"]
    captured_col = anchor_col + EMPIRICAL_OFFSETS["sales_captured_in_db_pct"]
    forecast_col = anchor_col + EMPIRICAL_OFFSETS["forecast_value"]
    forecast_min_col = anchor_col + EMPIRICAL_OFFSETS["forecast_min"]
    forecast_max_col = anchor_col + EMPIRICAL_OFFSETS["forecast_max"]

    data_rows: List[int] = []
    for row in range(snapshot.start_row, anchor_row):
        if to_float(snapshot.get(row, quarterly_sales_col)) is None:
            continue
        if to_float(snapshot.get(row, reported_sales_col)) is None:
            continue
        data_rows.append(row)

    if not data_rows:
        print(f"Skipped empirical extraction in {source_file}: no valid quarter rows found")
        return []

    helper_row = snapshot.end_row + 2
    helper_col = snapshot.end_col + 2
    avg_cell = sheet.range((helper_row, helper_col))
    forecast_cell = sheet.range((helper_row, helper_col + 1))

    rows: List[Dict[str, Any]] = []
    max_windows = min(N_QUARTERS, len(data_rows))
    for num_quarters in range(1, max_windows + 1):
        first_row = data_rows[-num_quarters]
        last_row = data_rows[-1]

        avg_formula = (
            f"=IFERROR(AVERAGE(R{first_row}C{captured_col}:R{last_row}C{captured_col}),NA())"
        )
        forecast_formula = (
            f"=IFERROR(R{last_row}C{reported_sales_col}/R{helper_row}C{helper_col},NA())"
        )
        set_formula2_r1c1(avg_cell, avg_formula)
        set_formula2_r1c1(forecast_cell, forecast_formula)
        wb.app.calculate()

        avg_penetration = to_float(avg_cell.value)
        forecast_value = to_float(forecast_cell.value)

        window_row = first_row
        forecast_max = to_float(snapshot.get(window_row, forecast_max_col))
        forecast_min = to_float(snapshot.get(window_row, forecast_min_col))
        if forecast_max is None and forecast_value is not None:
            forecast_max = forecast_value
        if forecast_min is None and forecast_value is not None:
            forecast_min = forecast_value

        range_width = None
        if forecast_max is not None and forecast_min is not None:
            range_width = forecast_max - forecast_min

        reported_sales = to_float(snapshot.get(last_row, reported_sales_col))
        quarterly_sales = to_float(snapshot.get(last_row, quarterly_sales_col))
        growth_rate = to_float(snapshot.get(last_row, growth_col))
        sales_captured_pct = to_float(snapshot.get(last_row, captured_col))
        last_quarter_used = snapshot.get(last_row, quarter_col)

        rows.append(
            {
                "model": label.model,
                "ticker": label.ticker,
                "model_period": label.model_period,
                "model_date": label.model_date,
                "method": "empirical",
                "parameter_name": "avg_penetration_pct",
                "parameter_value": avg_penetration,
                "num_quarters_used": num_quarters,
                "last_quarter_used": last_quarter_used,
                "forecast_value": forecast_value,
                "actual_value": reported_sales,
                "forecast_max": forecast_max,
                "forecast_min": forecast_min,
                "range_width": range_width,
                "avg_penetration_pct": avg_penetration,
                "quarterly_sales": quarterly_sales,
                "reported_sales": reported_sales,
                "growth_rate_pct": growth_rate,
                "sales_captured_in_db_pct": sales_captured_pct,
                "source_file": source_file,
            }
        )

    avg_cell.value = None
    forecast_cell.value = None
    return rows


def derive_regression_rows(
    wb: xw.Book,
    sheet: xw.Sheet,
    snapshot: SheetSnapshot,
    label: FileLabel,
    source_file: str,
) -> List[Dict[str, Any]]:
    anchor = find_anchor(snapshot, "max")
    if anchor is None:
        print(f"Skipped regression extraction in {source_file}: anchor 'max' not found")
        return []

    anchor_row, anchor_col = anchor
    y_col = anchor_col - 7
    x_col = anchor_col - 11

    data_rows: List[int] = []
    for row in range(snapshot.start_row, anchor_row):
        if to_float(snapshot.get(row, y_col)) is None:
            continue
        if to_float(snapshot.get(row, x_col)) is None:
            continue
        data_rows.append(row)

    if not data_rows:
        print(f"Skipped regression extraction in {source_file}: no valid x/y rows found")
        return []

    helper_row = snapshot.end_row + 2
    helper_col = snapshot.end_col + 4
    intercept_cell = sheet.range((helper_row, helper_col))
    slope_cell = sheet.range((helper_row, helper_col + 1))
    forecast_cell = sheet.range((helper_row, helper_col + 2))

    rows: List[Dict[str, Any]] = []
    max_windows = min(N_QUARTERS, len(data_rows))
    for num_quarters in range(1, max_windows + 1):
        first_row = data_rows[-num_quarters]
        last_row = data_rows[-1]

        intercept_formula = (
            f"=IFERROR(INTERCEPT(R{first_row}C{y_col}:R{last_row}C{y_col},"
            f"R{first_row}C{x_col}:R{last_row}C{x_col}),NA())"
        )
        slope_formula = (
            f"=IFERROR(SLOPE(R{first_row}C{y_col}:R{last_row}C{y_col},"
            f"R{first_row}C{x_col}:R{last_row}C{x_col}),NA())"
        )
        forecast_formula = (
            f"=IFERROR(R{helper_row}C{helper_col}+R{helper_row}C{helper_col + 1}*R{last_row}C{x_col},NA())"
        )

        set_formula2_r1c1(intercept_cell, intercept_formula)
        set_formula2_r1c1(slope_cell, slope_formula)
        set_formula2_r1c1(forecast_cell, forecast_formula)
        wb.app.calculate()

        intercept = to_float(intercept_cell.value)
        slope = to_float(slope_cell.value)
        forecast_total = to_float(forecast_cell.value)

        window_row = first_row
        forecast_from_grid = to_float(
            snapshot.get(window_row, anchor_col + REGRESSION_OFFSETS["forecast_total_without_sa"])
        )
        if forecast_from_grid is not None:
            forecast_total = forecast_from_grid

        forecast_max = to_float(snapshot.get(window_row, anchor_col + REGRESSION_OFFSETS["forecast_max"]))
        forecast_min = to_float(snapshot.get(window_row, anchor_col + REGRESSION_OFFSETS["forecast_min"]))
        if forecast_max is None and forecast_total is not None:
            forecast_max = forecast_total
        if forecast_min is None and forecast_total is not None:
            forecast_min = forecast_total

        range_width = None
        if forecast_max is not None and forecast_min is not None:
            range_width = forecast_max - forecast_min

        candidate = {
            "model": label.model,
            "ticker": label.ticker,
            "model_period": label.model_period,
            "model_date": label.model_date,
            "method": "regression",
            "parameter_name": "num_quarters_used",
            "parameter_value": num_quarters,
            "num_quarters_used": num_quarters,
            "forecast_value": forecast_total,
            "actual_value": None,
            "forecast_max": forecast_max,
            "forecast_min": forecast_min,
            "range_width": range_width,
            "intercept": intercept,
            "slope": slope,
            "source_file": source_file,
        }

        if rows:
            previous = rows[-1]
            duplicate = (
                almost_equal(previous.get("forecast_value"), candidate["forecast_value"])
                and almost_equal(previous.get("forecast_max"), candidate["forecast_max"])
                and almost_equal(previous.get("forecast_min"), candidate["forecast_min"])
                and almost_equal(previous.get("intercept"), candidate["intercept"])
                and almost_equal(previous.get("slope"), candidate["slope"])
            )
            if duplicate:
                continue

        rows.append(candidate)

    intercept_cell.value = None
    slope_cell.value = None
    forecast_cell.value = None
    return rows


def build_output_path(input_path: Path, output_path: Path) -> Path:
    output_path.mkdir(parents=True, exist_ok=True)
    base_name = f"{input_path.resolve().name}_PARAM"
    candidate = output_path / f"{base_name}.xlsx"
    if not candidate.exists():
        return candidate

    counter = 1
    while True:
        candidate = output_path / f"{base_name}.{counter}.xlsx"
        if not candidate.exists():
            return candidate
        counter += 1


def write_sheet(ws: Any, columns: Sequence[str], rows: Sequence[Dict[str, Any]]) -> None:
    ws.append(list(columns))
    for row_data in rows:
        ws.append([row_data.get(col) for col in columns])

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    for cell in ws[1]:
        cell.font = Font(bold=True)

    for col_idx, col_name in enumerate(columns, start=1):
        max_len = len(col_name)
        for row_idx in range(2, ws.max_row + 1):
            value = ws.cell(row=row_idx, column=col_idx).value
            if value is None:
                continue
            max_len = max(max_len, len(str(value)))
        ws.column_dimensions[get_column_letter(col_idx)].width = min(max(max_len + 2, 12), 40)


def run() -> None:
    if not input_dir.exists():
        print(f"Skipped: input directory not found -> {input_dir.resolve()}")
        return

    empirical_rows: List[Dict[str, Any]] = []
    regression_rows: List[Dict[str, Any]] = []
    processed_files = 0

    app = xw.App(visible=False, add_book=False)
    app.display_alerts = False
    app.screen_updating = False

    try:
        for file_path in sorted(input_dir.iterdir(), key=lambda p: p.name.lower()):
            if not file_path.is_file():
                continue
            if file_path.name.startswith("~"):
                print(f"Skipped {file_path.name}: temporary workbook")
                continue
            if file_path.suffix.lower() != ".xlsx":
                print(f"Skipped {file_path.name}: unsupported extension")
                continue

            label = parse_file_label(file_path.name)
            if label is None:
                print(f"Skipped {file_path.name}: could not parse ticker/model period")
                continue

            print(f"Processing {file_path.name}")
            wb: Optional[xw.Book] = None
            try:
                wb = app.books.open(str(file_path), update_links=False)

                empirical_sheet = get_sheet_if_present(wb, "Empirical Model")
                if empirical_sheet is None:
                    print(f"Skipped empirical in {file_path.name}: sheet not found")
                else:
                    empirical_snapshot = to_snapshot(empirical_sheet)
                    empirical_rows.extend(
                        derive_empirical_rows(
                            wb=wb,
                            sheet=empirical_sheet,
                            snapshot=empirical_snapshot,
                            label=label,
                            source_file=file_path.name,
                        )
                    )

                regression_sheet = get_sheet_if_present(wb, "Regression Model")
                if regression_sheet is None:
                    print(f"Skipped regression in {file_path.name}: sheet not found")
                else:
                    regression_snapshot = to_snapshot(regression_sheet)
                    regression_rows.extend(
                        derive_regression_rows(
                            wb=wb,
                            sheet=regression_sheet,
                            snapshot=regression_snapshot,
                            label=label,
                            source_file=file_path.name,
                        )
                    )

                processed_files += 1
            except Exception as exc:
                print(f"Skipped {file_path.name}: processing error -> {exc}")
            finally:
                if wb is not None:
                    safe_close_source_workbook(wb)
    finally:
        app.quit()

    output_file = build_output_path(input_dir, output_dir)
    out_wb = Workbook()
    empirical_ws = out_wb.active
    empirical_ws.title = "empirical_candidates"
    write_sheet(empirical_ws, EMPIRICAL_COLUMNS, empirical_rows)

    regression_ws = out_wb.create_sheet("regression_candidates")
    write_sheet(regression_ws, REGRESSION_COLUMNS, regression_rows)
    out_wb.save(output_file)

    print(f"Output file: {output_file.resolve()}")
    print(f"Files processed: {processed_files}")
    print(f"Empirical rows: {len(empirical_rows)}")
    print(f"Regression rows: {len(regression_rows)}")


if __name__ == "__main__":
    run()
