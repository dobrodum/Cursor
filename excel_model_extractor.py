from __future__ import annotations

import math
import re
from dataclasses import dataclass
from datetime import date
from pathlib import Path
from typing import Any, Iterable, Sequence

import xlwings as xw
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

# Configure these two folders before running.
input_dir = Path("/path/to/input")
output_dir = Path("/path/to/output")

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

PHASE_DAY = {"early": 5, "mid": 15, "late": 25}
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
PERIOD_RE = re.compile(r"(Early|Mid|Late)\s*([A-Za-z]{3,9})\s*[_-]?\s*(20\d{2})", re.IGNORECASE)


@dataclass
class FileLabel:
    ticker: str
    model_period: str
    model_date: str
    model: str


@dataclass
class SheetSnapshot:
    start_row: int
    start_col: int
    end_row: int
    end_col: int
    values: list[list[Any]]

    def get(self, row: int, col: int) -> Any:
        row_idx = row - self.start_row
        col_idx = col - self.start_col
        if row_idx < 0 or col_idx < 0:
            return None
        if row_idx >= len(self.values):
            return None
        row_values = self.values[row_idx]
        if col_idx >= len(row_values):
            return None
        return row_values[col_idx]

    def row_values(self, row: int) -> list[Any]:
        row_idx = row - self.start_row
        if row_idx < 0 or row_idx >= len(self.values):
            return []
        return self.values[row_idx]


def to_2d(values: Any) -> list[list[Any]]:
    if values is None:
        return []
    if isinstance(values, list):
        if not values:
            return []
        if isinstance(values[0], list):
            return values
        return [values]
    return [[values]]


def normalize_text(value: Any) -> str:
    if value is None:
        return ""
    text = str(value).strip().lower()
    return re.sub(r"[^a-z0-9]+", " ", text).strip()


def coerce_float(value: Any) -> float | None:
    if value is None or value == "":
        return None
    if isinstance(value, bool):
        return None
    try:
        num = float(value)
    except (TypeError, ValueError):
        return None
    if math.isnan(num) or math.isinf(num):
        return None
    return num


def normalize_ratio(value: Any) -> float | None:
    num = coerce_float(value)
    if num is None:
        return None
    if abs(num) > 1.5:
        return num / 100.0
    return num


def first_non_empty(*values: Any) -> Any:
    for value in values:
        if value is None:
            continue
        if isinstance(value, str) and value.strip() == "":
            continue
        return value
    return None


def first_numeric(*values: Any) -> float | None:
    for value in values:
        num = coerce_float(value)
        if num is not None:
            return num
    return None


def snapshot_sheet(sheet: xw.Sheet) -> SheetSnapshot:
    used = sheet.used_range
    raw_values = to_2d(used.value)
    row_count = int(used.rows.count)
    col_count = int(used.columns.count)
    if not raw_values:
        raw_values = [[None] * max(1, col_count)]
        row_count = max(1, row_count)
        col_count = max(1, col_count)
    normalized_rows: list[list[Any]] = []
    for row in raw_values:
        row_values = row[:] if isinstance(row, list) else [row]
        if len(row_values) < col_count:
            row_values.extend([None] * (col_count - len(row_values)))
        normalized_rows.append(row_values)
    if len(normalized_rows) < row_count:
        for _ in range(row_count - len(normalized_rows)):
            normalized_rows.append([None] * col_count)
    start_row = int(used.row)
    start_col = int(used.column)
    end_row = start_row + row_count - 1
    end_col = start_col + col_count - 1
    return SheetSnapshot(
        start_row=start_row,
        start_col=start_col,
        end_row=end_row,
        end_col=end_col,
        values=normalized_rows,
    )


def find_max_anchor(snapshot: SheetSnapshot) -> tuple[int, int] | None:
    for row_idx, row_values in enumerate(snapshot.values):
        for col_idx, value in enumerate(row_values):
            if normalize_text(value) == "max":
                return snapshot.start_row + row_idx, snapshot.start_col + col_idx
    return None


def resolve_column(
    snapshot: SheetSnapshot,
    anchor_row: int,
    keyword_groups: Sequence[tuple[str, ...]],
    fallback_col: int,
) -> int:
    search_rows: list[int] = [
        anchor_row,
        anchor_row - 1,
        anchor_row + 1,
        anchor_row - 2,
        anchor_row + 2,
    ]
    best_col: int | None = None
    best_distance = 10**9
    for row in search_rows:
        row_values = snapshot.row_values(row)
        if not row_values:
            continue
        for idx, value in enumerate(row_values):
            text = normalize_text(value)
            if not text:
                continue
            if any(all(keyword in text for keyword in group) for group in keyword_groups):
                distance = abs(row - anchor_row)
                if distance < best_distance:
                    best_distance = distance
                    best_col = snapshot.start_col + idx
    if best_col is not None:
        return best_col
    return fallback_col


def collect_numeric_rows(
    snapshot: SheetSnapshot,
    col: int,
    start_row: int,
    end_row: int,
    required_pair_col: int | None = None,
) -> list[int]:
    rows: list[int] = []
    for row in range(max(start_row, snapshot.start_row), min(end_row, snapshot.end_row) + 1):
        first = coerce_float(snapshot.get(row, col))
        if first is None:
            continue
        if required_pair_col is not None and coerce_float(snapshot.get(row, required_pair_col)) is None:
            continue
        rows.append(row)
    return rows


def parse_file_label(file_path: Path) -> FileLabel:
    stem = file_path.stem
    parts = [part.strip() for part in stem.split(" - ")]
    ticker = ""
    if len(parts) >= 2 and parts[1]:
        ticker = re.sub(r"[^A-Za-z0-9]+", "", parts[1]).upper()
    if not ticker:
        fallback_match = re.search(r"\b[A-Z]{2,6}\b", stem)
        ticker = fallback_match.group(0) if fallback_match else "UNKNOWN"

    period_match = PERIOD_RE.search(stem)
    if period_match:
        phase = period_match.group(1).title()
        month_token = period_match.group(2)[:3].lower()
        year = period_match.group(3)
        month_num = MONTH_NUM.get(month_token)
        if month_num is None:
            model_period = "UNKNOWN"
            model_date = ""
        else:
            model_period = f"{phase}{month_token.title()}_{year}"
            day = PHASE_DAY[phase.lower()]
            model_date = date(int(year), month_num, day).isoformat()
    else:
        model_period = "UNKNOWN"
        model_date = ""

    model = f"{ticker}_{model_period}" if model_period else ticker
    return FileLabel(
        ticker=ticker,
        model_period=model_period,
        model_date=model_date,
        model=model,
    )


def next_output_path(input_dir: Path, output_dir: Path) -> Path:
    base_name = f"{input_dir.name}_PARAM"
    candidate = output_dir / f"{base_name}.xlsx"
    counter = 1
    while candidate.exists():
        candidate = output_dir / f"{base_name}.{counter}.xlsx"
        counter += 1
    return candidate


def close_source_workbook_safely(wb: xw.Book) -> None:
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


def make_empirical_rows(
    sheet: xw.Sheet,
    wb_app: xw.App,
    file_label: FileLabel,
    source_file: str,
) -> list[dict[str, Any]]:
    snapshot = snapshot_sheet(sheet)
    anchor = find_max_anchor(snapshot)
    if anchor is None:
        print(f"  skipped empirical extraction (max anchor not found): {source_file}")
        return []

    anchor_row, anchor_col = anchor
    forecast_min_col = resolve_column(snapshot, anchor_row, [("min",)], anchor_col + 1)
    forecast_value_col = resolve_column(
        snapshot,
        anchor_row,
        [
            ("estimated", "total", "sold"),
            ("forecast", "value"),
            ("forecast", "total"),
            ("tot", "fcst"),
        ],
        anchor_col - 1,
    )
    num_quarters_col = resolve_column(
        snapshot,
        anchor_row,
        [
            ("num", "quarter"),
            ("quarters", "used"),
            ("n", "quarters"),
        ],
        anchor_col - 6,
    )
    last_quarter_col = resolve_column(
        snapshot,
        anchor_row,
        [("last", "quarter"), ("quarter", "used")],
        anchor_col - 10,
    )
    penetration_col = resolve_column(
        snapshot,
        anchor_row,
        [("avg", "penetration"), ("penetration",)],
        anchor_col - 4,
    )
    quarterly_sales_col = resolve_column(
        snapshot,
        anchor_row,
        [("quarterly", "sales"), ("qtr", "sales"), ("quarter", "sales")],
        anchor_col - 9,
    )
    reported_sales_col = resolve_column(
        snapshot,
        anchor_row,
        [("reported", "sales"), ("actual", "sales")],
        anchor_col - 8,
    )
    growth_rate_col = resolve_column(
        snapshot,
        anchor_row,
        [("growth", "rate"), ("growth",)],
        anchor_col - 7,
    )
    sales_captured_col = resolve_column(
        snapshot,
        anchor_row,
        [("captured", "db"), ("sales", "captured"), ("captured", "in", "db")],
        anchor_col - 5,
    )

    penetration_rows = collect_numeric_rows(snapshot, penetration_col, anchor_row + 1, snapshot.end_row)
    if not penetration_rows and sales_captured_col != penetration_col:
        penetration_rows = collect_numeric_rows(snapshot, sales_captured_col, anchor_row + 1, snapshot.end_row)
        penetration_col = sales_captured_col
    if not penetration_rows:
        print(f"  skipped empirical extraction (no numeric penetration series): {source_file}")
        return []

    latest_row = penetration_rows[-1]
    fallback_last_quarter = snapshot.get(latest_row, last_quarter_col)
    fallback_quarterly_sales = coerce_float(snapshot.get(latest_row, quarterly_sales_col))
    fallback_reported_sales = coerce_float(snapshot.get(latest_row, reported_sales_col))
    fallback_growth_rate = coerce_float(snapshot.get(latest_row, growth_rate_col))
    fallback_sales_captured = coerce_float(snapshot.get(latest_row, sales_captured_col))

    temp_avg_col = snapshot.end_col + 5
    avg_cell = sheet.range((anchor_row, temp_avg_col))

    rows: list[dict[str, Any]] = []
    max_iterations = min(N_QUARTERS, len(penetration_rows))
    for n in range(1, max_iterations + 1):
        start_row = penetration_rows[-n]
        end_row = penetration_rows[-1]
        avg_cell.formula2 = f"=AVERAGE(R{start_row}C{penetration_col}:R{end_row}C{penetration_col})"
        wb_app.calculate()

        avg_penetration = coerce_float(avg_cell.value)
        output_row = anchor_row + n

        num_quarters_used = int(
            round(first_numeric(snapshot.get(output_row, num_quarters_col), n))
        )
        last_quarter_used = first_non_empty(
            snapshot.get(output_row, last_quarter_col),
            fallback_last_quarter,
        )
        quarterly_sales = first_numeric(
            snapshot.get(output_row, quarterly_sales_col),
            fallback_quarterly_sales,
        )
        reported_sales = first_numeric(
            snapshot.get(output_row, reported_sales_col),
            fallback_reported_sales,
        )
        growth_rate_pct = first_numeric(
            snapshot.get(output_row, growth_rate_col),
            fallback_growth_rate,
        )
        sales_captured_in_db_pct = first_numeric(
            snapshot.get(output_row, sales_captured_col),
            fallback_sales_captured,
            avg_penetration,
        )

        forecast_value = first_numeric(snapshot.get(output_row, forecast_value_col))
        penetration_ratio = normalize_ratio(avg_penetration)
        if forecast_value is None and quarterly_sales is not None and penetration_ratio not in (None, 0):
            forecast_value = quarterly_sales / penetration_ratio

        forecast_max = first_numeric(snapshot.get(output_row, anchor_col))
        forecast_min = first_numeric(snapshot.get(output_row, forecast_min_col))
        growth_ratio = normalize_ratio(growth_rate_pct) or 0.0
        if forecast_max is None and forecast_value is not None:
            forecast_max = forecast_value * (1 + growth_ratio)
        if forecast_min is None and forecast_value is not None:
            forecast_min = forecast_value * (1 - growth_ratio)

        range_width = (
            forecast_max - forecast_min
            if forecast_max is not None and forecast_min is not None
            else None
        )

        rows.append(
            {
                "model": file_label.model,
                "ticker": file_label.ticker,
                "model_period": file_label.model_period,
                "model_date": file_label.model_date,
                "method": "empirical",
                "parameter_name": "avg_penetration_pct",
                "parameter_value": avg_penetration,
                "num_quarters_used": num_quarters_used,
                "last_quarter_used": last_quarter_used,
                "forecast_value": forecast_value,
                "actual_value": reported_sales,
                "forecast_max": forecast_max,
                "forecast_min": forecast_min,
                "range_width": range_width,
                "avg_penetration_pct": avg_penetration,
                "quarterly_sales": quarterly_sales,
                "reported_sales": reported_sales,
                "growth_rate_pct": growth_rate_pct,
                "sales_captured_in_db_pct": sales_captured_in_db_pct,
                "source_file": source_file,
            }
        )

    return rows


def make_regression_rows(
    sheet: xw.Sheet,
    wb_app: xw.App,
    file_label: FileLabel,
    source_file: str,
) -> list[dict[str, Any]]:
    snapshot = snapshot_sheet(sheet)
    anchor = find_max_anchor(snapshot)
    if anchor is None:
        print(f"  skipped regression extraction (max anchor not found): {source_file}")
        return []

    anchor_row, anchor_col = anchor
    x_col = anchor_col - 11
    y_col = anchor_col - 7
    num_quarters_col = resolve_column(
        snapshot,
        anchor_row,
        [("num", "quarter"), ("quarters", "used"), ("n", "quarters")],
        anchor_col - 6,
    )
    forecast_col = resolve_column(
        snapshot,
        anchor_row,
        [
            ("tot", "fcst"),
            ("forecast", "without", "sa"),
            ("forecast", "total"),
            ("forecast", "value"),
        ],
        anchor_col - 1,
    )
    min_col = resolve_column(snapshot, anchor_row, [("min",)], anchor_col + 1)
    actual_col = resolve_column(
        snapshot,
        anchor_row,
        [("actual", "sales"), ("actual",)],
        anchor_col - 2,
    )

    series_rows = collect_numeric_rows(
        snapshot,
        x_col,
        anchor_row + 1,
        snapshot.end_row,
        required_pair_col=y_col,
    )
    if not series_rows:
        print(f"  skipped regression extraction (no numeric x/y series): {source_file}")
        return []

    temp_base_col = snapshot.end_col + 8
    max_iterations = min(N_QUARTERS, len(series_rows))
    rows: list[dict[str, Any]] = []
    previous_signature: tuple[Any, ...] | None = None

    for n in range(1, max_iterations + 1):
        start_row = series_rows[-n]
        end_row = series_rows[-1]
        temp_row = anchor_row + n

        intercept_cell = sheet.range((temp_row, temp_base_col))
        slope_cell = sheet.range((temp_row, temp_base_col + 1))
        forecast_cell = sheet.range((temp_row, temp_base_col + 2))

        intercept_cell.formula2 = (
            f"=INTERCEPT(R{start_row}C{y_col}:R{end_row}C{y_col},"
            f"R{start_row}C{x_col}:R{end_row}C{x_col})"
        )
        slope_cell.formula2 = (
            f"=SLOPE(R{start_row}C{y_col}:R{end_row}C{y_col},"
            f"R{start_row}C{x_col}:R{end_row}C{x_col})"
        )

        next_x_row = end_row + 1
        has_next_x = coerce_float(snapshot.get(next_x_row, x_col)) is not None
        x_ref_row = next_x_row if has_next_x else end_row
        forecast_cell.formula2 = (
            f"=R{temp_row}C{temp_base_col}+R{temp_row}C{temp_base_col + 1}*R{x_ref_row}C{x_col}"
        )
        wb_app.calculate()

        intercept = coerce_float(intercept_cell.value)
        slope = coerce_float(slope_cell.value)
        forecast_from_formula = coerce_float(forecast_cell.value)

        output_row = anchor_row + n
        num_quarters_used = int(
            round(first_numeric(snapshot.get(output_row, num_quarters_col), n))
        )
        forecast_value = first_numeric(
            snapshot.get(output_row, forecast_col),
            forecast_from_formula,
        )
        forecast_max = first_numeric(snapshot.get(output_row, anchor_col))
        forecast_min = first_numeric(snapshot.get(output_row, min_col))
        if forecast_max is None and forecast_value is not None:
            forecast_max = forecast_value * 1.05
        if forecast_min is None and forecast_value is not None:
            forecast_min = forecast_value * 0.95
        range_width = (
            forecast_max - forecast_min
            if forecast_max is not None and forecast_min is not None
            else None
        )
        actual_value = first_numeric(snapshot.get(output_row, actual_col))

        signature = (
            round(intercept, 10) if intercept is not None else None,
            round(slope, 10) if slope is not None else None,
            round(forecast_value, 10) if forecast_value is not None else None,
            round(forecast_max, 10) if forecast_max is not None else None,
            round(forecast_min, 10) if forecast_min is not None else None,
        )
        if signature == previous_signature:
            continue
        previous_signature = signature

        rows.append(
            {
                "model": file_label.model,
                "ticker": file_label.ticker,
                "model_period": file_label.model_period,
                "model_date": file_label.model_date,
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
        )

    return rows


def write_output_sheet(
    ws: Any,
    columns: Sequence[str],
    rows: Sequence[dict[str, Any]],
) -> None:
    ws.append(list(columns))
    for col_idx, _ in enumerate(columns, start=1):
        ws.cell(row=1, column=col_idx).font = Font(bold=True)

    for row in rows:
        ws.append([row.get(col) for col in columns])

    ws.freeze_panes = "A2"
    max_row = max(1, ws.max_row)
    max_col = len(columns)
    ws.auto_filter.ref = f"A1:{get_column_letter(max_col)}{max_row}"

    for col_idx, col_name in enumerate(columns, start=1):
        max_len = len(col_name)
        for row_idx in range(2, max_row + 1):
            value = ws.cell(row=row_idx, column=col_idx).value
            if value is None:
                continue
            value_len = len(str(value))
            if value_len > max_len:
                max_len = value_len
        ws.column_dimensions[get_column_letter(col_idx)].width = min(48, max(12, max_len + 2))


def iter_input_entries(folder: Path) -> Iterable[Path]:
    return sorted(folder.iterdir(), key=lambda path: path.name.lower())


def run() -> None:
    if not input_dir.exists() or not input_dir.is_dir():
        raise FileNotFoundError(f"input_dir does not exist or is not a directory: {input_dir}")
    output_dir.mkdir(parents=True, exist_ok=True)

    output_path = next_output_path(input_dir, output_dir)
    empirical_rows: list[dict[str, Any]] = []
    regression_rows: list[dict[str, Any]] = []
    processed_files = 0

    print(f"input_dir={input_dir}")
    print(f"output_dir={output_dir}")

    app = xw.App(visible=False, add_book=False)
    app.display_alerts = False
    app.screen_updating = False
    try:
        try:
            app.api.Calculation = -4135  # xlCalculationManual
        except Exception:
            pass

        for path in iter_input_entries(input_dir):
            if not path.is_file():
                print(f"skipped file: {path.name} (not a regular file)")
                continue
            if path.name.startswith("~"):
                print(f"skipped file: {path.name} (temp file)")
                continue
            if path.suffix.lower() != ".xlsx":
                print(f"skipped file: {path.name} (not .xlsx)")
                continue

            print(f"processed file: {path.name}")
            file_label = parse_file_label(path)
            wb: xw.Book | None = None
            try:
                wb = app.books.open(str(path), update_links=False)
                processed_files += 1

                if "Empirical Model" in [sheet.name for sheet in wb.sheets]:
                    empirical_rows.extend(
                        make_empirical_rows(
                            sheet=wb.sheets["Empirical Model"],
                            wb_app=wb.app,
                            file_label=file_label,
                            source_file=path.name,
                        )
                    )
                else:
                    print(f"  skipped empirical extraction (missing 'Empirical Model'): {path.name}")

                if "Regression Model" in [sheet.name for sheet in wb.sheets]:
                    regression_rows.extend(
                        make_regression_rows(
                            sheet=wb.sheets["Regression Model"],
                            wb_app=wb.app,
                            file_label=file_label,
                            source_file=path.name,
                        )
                    )
                else:
                    print(f"  skipped regression extraction (missing 'Regression Model'): {path.name}")

            except Exception as exc:
                print(f"skipped file: {path.name} (error opening/processing: {exc})")
            finally:
                if wb is not None:
                    close_source_workbook_safely(wb)

    finally:
        app.quit()

    out_wb = Workbook()
    empirical_ws = out_wb.active
    empirical_ws.title = "empirical_candidates"
    regression_ws = out_wb.create_sheet("regression_candidates")

    write_output_sheet(empirical_ws, EMPIRICAL_COLUMNS, empirical_rows)
    write_output_sheet(regression_ws, REGRESSION_COLUMNS, regression_rows)
    out_wb.save(output_path)

    print(f"output_path={output_path}")
    print(f"files_processed={processed_files}")
    print(f"empirical_rows={len(empirical_rows)}")
    print(f"regression_rows={len(regression_rows)}")


if __name__ == "__main__":
    run()
