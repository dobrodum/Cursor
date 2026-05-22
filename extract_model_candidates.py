#!/usr/bin/env python3
from __future__ import annotations

import re
import statistics
from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path
from typing import Any, Iterable

import xlwings as xw
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

# User-configurable paths
input_dir = Path("./input")
output_dir = Path("./output")

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

DAY_BY_PERIOD = {"early": 5, "mid": 15, "late": 25}


@dataclass
class SheetSnapshot:
    start_row: int
    start_col: int
    values: list[list[Any]]

    @property
    def end_row(self) -> int:
        return self.start_row + len(self.values) - 1

    @property
    def end_col(self) -> int:
        if not self.values:
            return self.start_col
        return self.start_col + len(self.values[0]) - 1

    def row_values(self, abs_row: int) -> list[Any]:
        idx = abs_row - self.start_row
        if idx < 0 or idx >= len(self.values):
            return []
        return self.values[idx]

    def get(self, abs_row: int, abs_col: int) -> Any:
        row_idx = abs_row - self.start_row
        col_idx = abs_col - self.start_col
        if row_idx < 0 or col_idx < 0:
            return None
        if row_idx >= len(self.values):
            return None
        row = self.values[row_idx]
        if col_idx >= len(row):
            return None
        return row[col_idx]


@dataclass
class HistoryPoint:
    row: int
    x: float
    y: float


def normalize_2d(values: Any) -> list[list[Any]]:
    if values is None:
        return [[]]
    if not isinstance(values, list):
        return [[values]]
    if values and not isinstance(values[0], list):
        return [values]

    matrix = values
    max_cols = max((len(row) for row in matrix), default=0)
    if max_cols == 0:
        return [[]]
    normalized: list[list[Any]] = []
    for row in matrix:
        padded = list(row) + [None] * (max_cols - len(row))
        normalized.append(padded)
    return normalized


def safe_float(value: Any) -> float | None:
    if value in (None, ""):
        return None
    try:
        return float(value)
    except (TypeError, ValueError):
        return None


def safe_int(value: Any) -> int | None:
    number = safe_float(value)
    if number is None:
        return None
    if abs(number - round(number)) > 1e-9:
        return None
    return int(round(number))


def round_or_none(value: float | None, digits: int = 8) -> float | None:
    if value is None:
        return None
    return round(value, digits)


def normalize_header(value: Any) -> str:
    if value is None:
        return ""
    text = str(value).strip().lower()
    text = re.sub(r"[^a-z0-9]+", " ", text)
    return re.sub(r"\s+", " ", text).strip()


def parse_month(month_token: str) -> int:
    month_token = month_token.strip()
    for fmt in ("%b", "%B"):
        try:
            return datetime.strptime(month_token.title(), fmt).month
        except ValueError:
            continue
    # Fallback for custom tokens like "Sept"
    return datetime.strptime(month_token[:3].title(), "%b").month


def parse_model_metadata(file_path: Path) -> dict[str, str]:
    stem = file_path.stem
    parts = [part.strip() for part in stem.split(" - ")]
    ticker = parts[1].upper() if len(parts) >= 2 and parts[1] else "UNKNOWN"

    period_token = parts[2] if len(parts) >= 3 else stem
    period_token = re.sub(r"(?i)_send.*$", "", period_token).strip()

    period_match = re.search(r"(?i)(Early|Mid|Late)\s*([A-Za-z]+)\s*(\d{4})", period_token)
    if period_match:
        period_tag = period_match.group(1).title()
        month_token = period_match.group(2)
        year = int(period_match.group(3))
        month = parse_month(month_token)
        month_abbrev = date(year, month, 1).strftime("%b")
        model_period = f"{period_tag}{month_abbrev}_{year}"
        day = DAY_BY_PERIOD[period_tag.lower()]
        model_date = date(year, month, day).isoformat()
    else:
        clean_period = re.sub(r"\s+", "", period_token)
        model_period = clean_period.replace("-", "_")
        model_date = ""

    model = f"{ticker}_{model_period}" if model_period else ticker
    return {
        "ticker": ticker,
        "model_period": model_period,
        "model_date": model_date,
        "model": model,
    }


def ensure_unique_output_path(in_dir: Path, out_dir: Path) -> Path:
    base_name = f"{in_dir.name}_PARAM"
    candidate = out_dir / f"{base_name}.xlsx"
    suffix = 1
    while candidate.exists():
        candidate = out_dir / f"{base_name}.{suffix}.xlsx"
        suffix += 1
    return candidate


def set_formula2(cell: Any, formula: str) -> None:
    try:
        cell.formula2 = formula
    except Exception:
        # Fallback for Excel builds without Formula2 support.
        cell.formula = formula


def close_source_workbook(wb: xw.Book) -> None:
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
        pass


def get_sheet(workbook: xw.Book, name: str) -> xw.Sheet | None:
    lookup = name.strip().lower()
    for sheet in workbook.sheets:
        if sheet.name.strip().lower() == lookup:
            return sheet
    return None


def take_snapshot(sheet: xw.Sheet) -> SheetSnapshot:
    used = sheet.used_range
    matrix = normalize_2d(used.value)
    return SheetSnapshot(start_row=used.row, start_col=used.column, values=matrix)


def find_anchor(snapshot: SheetSnapshot, anchor_text: str = "max") -> tuple[int, int] | None:
    needle = anchor_text.strip().lower()
    for r_offset, row in enumerate(snapshot.values):
        for c_offset, value in enumerate(row):
            if isinstance(value, str) and value.strip().lower() == needle:
                return snapshot.start_row + r_offset, snapshot.start_col + c_offset
    return None


def header_cells(snapshot: SheetSnapshot, header_row: int) -> list[tuple[str, int]]:
    values = snapshot.row_values(header_row)
    cells: list[tuple[str, int]] = []
    for idx, value in enumerate(values):
        normalized = normalize_header(value)
        if normalized:
            cells.append((normalized, snapshot.start_col + idx))
    return cells


def find_column(
    headers: list[tuple[str, int]],
    required_tokens: Iterable[str],
    forbidden_tokens: Iterable[str] = (),
) -> int | None:
    required = [token.lower() for token in required_tokens]
    forbidden = [token.lower() for token in forbidden_tokens]
    for text, col in headers:
        if all(token in text for token in required) and all(token not in text for token in forbidden):
            return col
    return None


def collect_history_points(
    snapshot: SheetSnapshot,
    row_start: int,
    row_end: int,
    x_col: int,
    y_col: int,
) -> list[HistoryPoint]:
    points: list[HistoryPoint] = []
    for row in range(max(row_start, snapshot.start_row), min(row_end, snapshot.end_row) + 1):
        x_val = safe_float(snapshot.get(row, x_col))
        y_val = safe_float(snapshot.get(row, y_col))
        if x_val is None or y_val in (None, 0):
            continue
        points.append(HistoryPoint(row=row, x=x_val, y=y_val))
    return points


def first_numeric(snapshot: SheetSnapshot, rows: Iterable[int], col: int) -> float | None:
    for row in rows:
        value = safe_float(snapshot.get(row, col))
        if value is not None:
            return value
    return None


def compute_padding(forecast_value: float | None, ratios: list[float]) -> float:
    if forecast_value is None:
        return 0.0
    if len(ratios) >= 2:
        volatility = statistics.pstdev(ratios)
        return abs(forecast_value) * max(volatility, 0.01)
    return abs(forecast_value) * 0.05


def extract_empirical_candidates(
    sheet: xw.Sheet,
    workbook: xw.Book,
    metadata: dict[str, str],
    source_file: str,
) -> list[dict[str, Any]]:
    snapshot = take_snapshot(sheet)
    anchor = find_anchor(snapshot, "max")
    if anchor is None:
        print(f"  skipped {sheet.name}: no 'max' anchor")
        return []

    anchor_row, anchor_col = anchor
    headers = header_cells(snapshot, anchor_row)

    x_col = find_column(headers, ["quarterly", "sales"], ["reported"]) or (anchor_col - 11)
    y_col = (
        find_column(headers, ["reported", "sales"])
        or find_column(headers, ["actual", "sales"])
        or (anchor_col - 7)
    )
    num_quarters_col = (
        find_column(headers, ["num", "quarter"])
        or find_column(headers, ["quarters", "used"])
        or (anchor_col - 6)
    )
    last_quarter_col = find_column(headers, ["last", "quarter"]) or (anchor_col - 5)
    forecast_col = (
        find_column(headers, ["estimated", "total"])
        or find_column(headers, ["forecast"])
        or (anchor_col - 4)
    )
    avg_pen_col = (
        find_column(headers, ["avg", "penetration"])
        or find_column(headers, ["penetration"])
        or (anchor_col - 3)
    )
    growth_col = find_column(headers, ["growth"]) or (anchor_col + 2)
    captured_col = find_column(headers, ["captured", "db"]) or find_column(headers, ["captured"]) or (anchor_col + 3)
    min_col = find_column(headers, ["min"]) or (anchor_col + 1)

    history = collect_history_points(snapshot, snapshot.start_row, anchor_row - 1, x_col, y_col)
    if not history:
        history = collect_history_points(snapshot, anchor_row + 1, snapshot.end_row, x_col, y_col)
    if not history:
        print(f"  skipped {sheet.name}: no numeric history for empirical extraction")
        return []

    max_quarters = min(10, len(history))
    helper_row = anchor_row
    helper_col = snapshot.end_col + 4
    helper_avg = sheet.cells(helper_row, helper_col)
    helper_forecast = sheet.cells(helper_row, helper_col + 1)

    forecast_x = first_numeric(snapshot, [anchor_row + 1, anchor_row], x_col) or history[-1].x
    base_actual = first_numeric(snapshot, [anchor_row + 1, anchor_row], y_col)

    rows: list[dict[str, Any]] = []
    for n_quarters in range(1, max_quarters + 1):
        subset = history[-n_quarters:]
        first_row = subset[0].row
        last_row = subset[-1].row

        avg_formula = (
            f'=IFERROR(AVERAGE((R{first_row}C{x_col}:R{last_row}C{x_col})/'
            f'(R{first_row}C{y_col}:R{last_row}C{y_col})), "")'
        )
        forecast_formula = f'=IFERROR({forecast_x}/R{helper_row}C{helper_col}, "")'
        set_formula2(helper_avg, avg_formula)
        set_formula2(helper_forecast, forecast_formula)
        workbook.app.calculate()

        table_row = anchor_row + n_quarters
        avg_penetration = safe_float(snapshot.get(table_row, avg_pen_col))
        if avg_penetration is None:
            avg_penetration = safe_float(helper_avg.value)

        forecast_value = safe_float(snapshot.get(table_row, forecast_col))
        if forecast_value is None:
            forecast_value = safe_float(helper_forecast.value)

        actual_value = safe_float(snapshot.get(table_row, y_col))
        if actual_value is None:
            actual_value = base_actual

        quarterly_sales = safe_float(snapshot.get(table_row, x_col))
        if quarterly_sales is None:
            quarterly_sales = forecast_x

        reported_sales = safe_float(snapshot.get(table_row, y_col))
        if reported_sales is None:
            reported_sales = actual_value

        growth_rate_pct = safe_float(snapshot.get(table_row, growth_col))
        if growth_rate_pct is None and len(subset) >= 2 and subset[-2].y != 0:
            growth_rate_pct = (subset[-1].y / subset[-2].y) - 1

        sales_captured_pct = safe_float(snapshot.get(table_row, captured_col))
        if sales_captured_pct is None and reported_sales not in (None, 0) and quarterly_sales is not None:
            sales_captured_pct = quarterly_sales / reported_sales

        forecast_max = safe_float(snapshot.get(table_row, anchor_col))
        forecast_min = safe_float(snapshot.get(table_row, min_col))
        if forecast_max is None or forecast_min is None:
            ratios = [point.x / point.y for point in subset if point.y]
            spread = compute_padding(forecast_value, ratios)
            if forecast_max is None and forecast_value is not None:
                forecast_max = forecast_value + spread
            if forecast_min is None and forecast_value is not None:
                forecast_min = forecast_value - spread

        range_width = None
        if forecast_max is not None and forecast_min is not None:
            range_width = forecast_max - forecast_min

        last_quarter_used = snapshot.get(table_row, last_quarter_col)
        if last_quarter_used in (None, ""):
            last_quarter_used = snapshot.get(subset[-1].row, last_quarter_col)

        table_num_quarters = safe_int(snapshot.get(table_row, num_quarters_col))
        if table_num_quarters is not None:
            n_quarters = table_num_quarters

        row = {
            "model": metadata["model"],
            "ticker": metadata["ticker"],
            "model_period": metadata["model_period"],
            "model_date": metadata["model_date"],
            "method": "empirical",
            "parameter_name": "avg_penetration_pct",
            "parameter_value": avg_penetration,
            "num_quarters_used": n_quarters,
            "last_quarter_used": last_quarter_used,
            "forecast_value": forecast_value,
            "actual_value": actual_value,
            "forecast_max": forecast_max,
            "forecast_min": forecast_min,
            "range_width": range_width,
            "avg_penetration_pct": avg_penetration,
            "quarterly_sales": quarterly_sales,
            "reported_sales": reported_sales,
            "growth_rate_pct": growth_rate_pct,
            "sales_captured_in_db_pct": sales_captured_pct,
            "source_file": source_file,
        }
        rows.append(row)
    return rows


def extract_regression_candidates(
    sheet: xw.Sheet,
    workbook: xw.Book,
    metadata: dict[str, str],
    source_file: str,
) -> list[dict[str, Any]]:
    snapshot = take_snapshot(sheet)
    anchor = find_anchor(snapshot, "max")
    if anchor is None:
        print(f"  skipped {sheet.name}: no 'max' anchor")
        return []

    anchor_row, anchor_col = anchor
    headers = header_cells(snapshot, anchor_row)

    # Required anchor-relative source columns
    y_col = anchor_col - 7
    x_col = anchor_col - 11

    num_quarters_col = (
        find_column(headers, ["num", "quarter"])
        or find_column(headers, ["quarters", "used"])
        or (anchor_col - 6)
    )
    forecast_col = (
        find_column(headers, ["tot", "fcst", "w", "o", "sa"])
        or find_column(headers, ["forecast"])
        or (anchor_col - 2)
    )
    intercept_col = find_column(headers, ["intercept"]) or (anchor_col - 4)
    slope_col = find_column(headers, ["slope"]) or (anchor_col - 3)
    min_col = find_column(headers, ["min"]) or (anchor_col + 1)

    history = collect_history_points(snapshot, snapshot.start_row, anchor_row - 1, x_col, y_col)
    if len(history) < 2:
        history = collect_history_points(snapshot, anchor_row + 1, snapshot.end_row, x_col, y_col)
    if len(history) < 2:
        print(f"  skipped {sheet.name}: no numeric history for regression extraction")
        return []

    max_quarters = min(10, len(history))
    helper_row = anchor_row
    helper_col = snapshot.end_col + 8
    helper_intercept = sheet.cells(helper_row, helper_col)
    helper_slope = sheet.cells(helper_row, helper_col + 1)
    helper_forecast = sheet.cells(helper_row, helper_col + 2)

    forecast_x = first_numeric(snapshot, [anchor_row + 1, anchor_row], x_col) or history[-1].x

    rows: list[dict[str, Any]] = []
    previous_signature: tuple[Any, ...] | None = None

    for n_quarters in range(2, max_quarters + 1):
        subset = history[-n_quarters:]
        first_row = subset[0].row
        last_row = subset[-1].row

        intercept_formula = (
            f'=IFERROR(INTERCEPT(R{first_row}C{y_col}:R{last_row}C{y_col},'
            f'R{first_row}C{x_col}:R{last_row}C{x_col}), "")'
        )
        slope_formula = (
            f'=IFERROR(SLOPE(R{first_row}C{y_col}:R{last_row}C{y_col},'
            f'R{first_row}C{x_col}:R{last_row}C{x_col}), "")'
        )
        forecast_formula = (
            f'=IFERROR(R{helper_row}C{helper_col}+'
            f'R{helper_row}C{helper_col + 1}*{forecast_x}, "")'
        )
        set_formula2(helper_intercept, intercept_formula)
        set_formula2(helper_slope, slope_formula)
        set_formula2(helper_forecast, forecast_formula)
        workbook.app.calculate()

        table_row = anchor_row + n_quarters
        intercept = safe_float(snapshot.get(table_row, intercept_col))
        if intercept is None:
            intercept = safe_float(helper_intercept.value)

        slope = safe_float(snapshot.get(table_row, slope_col))
        if slope is None:
            slope = safe_float(helper_slope.value)

        forecast_value = safe_float(snapshot.get(table_row, forecast_col))
        if forecast_value is None:
            forecast_value = safe_float(helper_forecast.value)

        forecast_max = safe_float(snapshot.get(table_row, anchor_col))
        forecast_min = safe_float(snapshot.get(table_row, min_col))
        if (forecast_max is None or forecast_min is None) and forecast_value is not None:
            spread = abs(forecast_value) * 0.05
            if forecast_max is None:
                forecast_max = forecast_value + spread
            if forecast_min is None:
                forecast_min = forecast_value - spread

        range_width = None
        if forecast_max is not None and forecast_min is not None:
            range_width = forecast_max - forecast_min

        table_num_quarters = safe_int(snapshot.get(table_row, num_quarters_col))
        if table_num_quarters is not None:
            n_quarters = table_num_quarters

        actual_value = snapshot.get(table_row, y_col)
        if actual_value in (None, ""):
            actual_value = ""

        signature = (
            round_or_none(intercept, 10),
            round_or_none(slope, 10),
            round_or_none(forecast_value, 6),
            round_or_none(forecast_max, 6),
            round_or_none(forecast_min, 6),
        )
        if signature == previous_signature:
            continue
        previous_signature = signature

        row = {
            "model": metadata["model"],
            "ticker": metadata["ticker"],
            "model_period": metadata["model_period"],
            "model_date": metadata["model_date"],
            "method": "regression",
            "parameter_name": "num_quarters_used",
            "parameter_value": n_quarters,
            "num_quarters_used": n_quarters,
            "forecast_value": forecast_value,
            "actual_value": actual_value,
            "forecast_max": forecast_max,
            "forecast_min": forecast_min,
            "range_width": range_width,
            "intercept": intercept,
            "slope": slope,
            "source_file": source_file,
        }
        rows.append(row)

    return rows


def write_sheet(ws: Any, columns: list[str], rows: list[dict[str, Any]]) -> None:
    ws.append(columns)
    for row in rows:
        ws.append([row.get(column, "") for column in columns])

    for cell in ws[1]:
        cell.font = Font(bold=True)

    ws.freeze_panes = "A2"
    last_col_letter = get_column_letter(len(columns))
    ws.auto_filter.ref = f"A1:{last_col_letter}{max(1, ws.max_row)}"

    for col_idx, column in enumerate(columns, start=1):
        max_len = len(column)
        for row in rows:
            value = row.get(column)
            value_len = len(str(value)) if value not in (None, "") else 0
            if value_len > max_len:
                max_len = value_len
        ws.column_dimensions[get_column_letter(col_idx)].width = min(max(max_len + 2, 12), 48)


def write_output_workbook(
    output_path: Path,
    empirical_rows: list[dict[str, Any]],
    regression_rows: list[dict[str, Any]],
) -> None:
    workbook = Workbook()
    default_sheet = workbook.active
    workbook.remove(default_sheet)

    empirical_ws = workbook.create_sheet("empirical_candidates")
    regression_ws = workbook.create_sheet("regression_candidates")

    write_sheet(empirical_ws, EMPIRICAL_COLUMNS, empirical_rows)
    write_sheet(regression_ws, REGRESSION_COLUMNS, regression_rows)

    workbook.save(output_path)


def iter_input_files(in_dir: Path) -> list[Path]:
    selected: list[Path] = []
    for path in sorted(in_dir.iterdir()):
        if not path.is_file():
            continue
        name = path.name
        if name.startswith("~"):
            print(f"skipped {name}: temp file")
            continue
        if path.suffix.lower() != ".xlsx":
            print(f"skipped {name}: not .xlsx")
            continue
        if re.search(r"(?i)_PARAM(\.\d+)?\.xlsx$", name):
            print(f"skipped {name}: appears to be an output artifact")
            continue
        selected.append(path)
    return selected


def main() -> None:
    in_dir = Path(input_dir).expanduser().resolve()
    out_dir = Path(output_dir).expanduser().resolve()

    if not in_dir.exists():
        raise FileNotFoundError(f"input_dir does not exist: {in_dir}")
    out_dir.mkdir(parents=True, exist_ok=True)

    source_files = iter_input_files(in_dir)
    output_path = ensure_unique_output_path(in_dir, out_dir)

    empirical_rows: list[dict[str, Any]] = []
    regression_rows: list[dict[str, Any]] = []
    processed_files = 0

    app = xw.App(visible=False, add_book=False)
    try:
        app.display_alerts = False
    except Exception:
        pass
    try:
        app.screen_updating = False
    except Exception:
        pass
    try:
        app.calculation = "manual"
    except Exception:
        pass

    try:
        for file_path in source_files:
            print(f"processing {file_path.name}")
            metadata = parse_model_metadata(file_path)
            workbook: xw.Book | None = None
            try:
                workbook = app.books.open(str(file_path), update_links=False)
                empirical_sheet = get_sheet(workbook, "Empirical Model")
                regression_sheet = get_sheet(workbook, "Regression Model")

                if empirical_sheet is None:
                    print("  skipped Empirical Model: sheet not found")
                else:
                    empirical_rows.extend(
                        extract_empirical_candidates(empirical_sheet, workbook, metadata, file_path.name)
                    )

                if regression_sheet is None:
                    print("  skipped Regression Model: sheet not found")
                else:
                    regression_rows.extend(
                        extract_regression_candidates(regression_sheet, workbook, metadata, file_path.name)
                    )

                processed_files += 1
                print(f"processed {file_path.name}")
            except Exception as exc:
                print(f"skipped {file_path.name}: {exc}")
            finally:
                if workbook is not None:
                    close_source_workbook(workbook)
    finally:
        try:
            app.quit()
        except Exception:
            pass

    write_output_workbook(output_path, empirical_rows, regression_rows)

    print(f"output path: {output_path}")
    print(f"number of files processed: {processed_files}")
    print(f"number of empirical rows: {len(empirical_rows)}")
    print(f"number of regression rows: {len(regression_rows)}")


if __name__ == "__main__":
    main()
