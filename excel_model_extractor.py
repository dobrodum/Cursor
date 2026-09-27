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


# -------- inputs --------
input_dir = Path("./input")
output_dir = Path("./output")


EMPIRICAL_SHEET = "Empirical Model"
REGRESSION_SHEET = "Regression Model"
MAX_QUARTERS = 10


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


# Offsets are anchored from the cell containing "max".
EMPIRICAL_OFFSETS = {
    "avg_penetration_input": (-2, 1),
    "forecast_value": (-3, 1),
    "actual_value": (-1, 1),
    "forecast_max": (0, 1),
    "forecast_min": (1, 1),
    "quarterly_sales": (-5, 1),
    "reported_sales": (-1, 1),
    "growth_rate_pct": (-6, 1),
    "sales_captured_in_db_pct": (-7, 1),
    "penetration_row": (-1, 0),
    "quarter_label_row": (-2, 0),
    "temp_formula_cell": (10, 3),
}


REGRESSION_OFFSETS = {
    "forecast_total_without_sa": (-3, 1),
    "actual_value": (-2, 1),
    "forecast_max": (0, 1),
    "forecast_min": (1, 1),
    "temp_intercept_cell": (10, 3),
    "temp_slope_cell": (11, 3),
}


@dataclass(frozen=True)
class FileMetadata:
    model: str
    ticker: str
    model_period: str
    model_date: str


@dataclass
class SheetSnapshot:
    sheet: Any
    used_row: int
    used_col: int
    values: List[List[Any]]
    label_positions: Dict[str, List[Tuple[int, int]]]


def normalize_label(value: Any) -> str:
    if not isinstance(value, str):
        return ""
    return re.sub(r"\s+", " ", value.strip().lower())


def to_2d(values: Any) -> List[List[Any]]:
    if isinstance(values, list):
        if not values:
            return []
        if isinstance(values[0], list):
            return values
        return [values]
    return [[values]]


def to_number(value: Any) -> Optional[float]:
    if value is None:
        return None
    if isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        return float(value)
    if isinstance(value, str):
        raw = value.strip().replace(",", "")
        if not raw:
            return None
        if raw.endswith("%"):
            raw = raw[:-1].strip()
            try:
                return float(raw) / 100.0
            except ValueError:
                return None
        try:
            return float(raw)
        except ValueError:
            return None
    return None


def get_month_number(month_abbrev: str) -> int:
    month_map = {
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
    return month_map[month_abbrev.lower()]


def parse_file_metadata(file_path: Path) -> FileMetadata:
    stem = file_path.stem
    pattern = re.compile(
        r" - (?P<ticker>[A-Za-z0-9]+) - (?P<period>(?P<bucket>Early|Mid|Late)(?P<month>[A-Za-z]{3})(?P<year>\d{4}))",
        re.IGNORECASE,
    )
    match = pattern.search(stem)
    if not match:
        ticker = stem.split(" - ")[1].strip() if " - " in stem else stem
        model_period = "Unknown"
        model_date = ""
        model = f"{ticker}_{model_period}"
        return FileMetadata(model=model, ticker=ticker, model_period=model_period, model_date=model_date)

    ticker = match.group("ticker").upper()
    bucket = match.group("bucket")
    month = match.group("month").title()
    year = int(match.group("year"))

    day_map = {"early": 5, "mid": 15, "late": 25}
    day = day_map[bucket.lower()]
    month_num = get_month_number(month)
    model_date = date(year, month_num, day).isoformat()
    model_period = f"{bucket.title()}{month}_{year}"
    model = f"{ticker}_{model_period}"
    return FileMetadata(model=model, ticker=ticker, model_period=model_period, model_date=model_date)


def build_output_path(in_dir: Path, out_dir: Path) -> Path:
    out_dir.mkdir(parents=True, exist_ok=True)
    base_name = f"{in_dir.name}_PARAM.xlsx"
    candidate = out_dir / base_name
    if not candidate.exists():
        return candidate

    idx = 1
    while True:
        candidate = out_dir / f"{in_dir.name}_PARAM.{idx}.xlsx"
        if not candidate.exists():
            return candidate
        idx += 1


def close_workbook_without_save(wb: Any) -> None:
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
        return
    except Exception:
        pass

    try:
        wb.api.Close(False)
    except Exception:
        pass


def build_sheet_snapshot(sheet: Any) -> SheetSnapshot:
    used = sheet.used_range
    values = to_2d(used.value)
    row0, col0 = used.row, used.column
    positions: Dict[str, List[Tuple[int, int]]] = {}
    for r_idx, row_values in enumerate(values):
        for c_idx, cell_value in enumerate(row_values):
            label = normalize_label(cell_value)
            if not label:
                continue
            abs_row = row0 + r_idx
            abs_col = col0 + c_idx
            positions.setdefault(label, []).append((abs_row, abs_col))
    return SheetSnapshot(sheet=sheet, used_row=row0, used_col=col0, values=values, label_positions=positions)


def find_anchor(snapshot: SheetSnapshot, label: str = "max") -> Optional[Tuple[int, int]]:
    points = snapshot.label_positions.get(normalize_label(label), [])
    if not points:
        return None
    return points[0]


def nearest_label_position(
    snapshot: SheetSnapshot,
    label: str,
    anchor_row: int,
    anchor_col: int,
) -> Optional[Tuple[int, int]]:
    points = snapshot.label_positions.get(normalize_label(label), [])
    if not points:
        return None
    return min(points, key=lambda rc: abs(rc[0] - anchor_row) + abs(rc[1] - anchor_col))


def read_offset(sheet: Any, anchor_row: int, anchor_col: int, offset: Tuple[int, int]) -> Any:
    row = anchor_row + offset[0]
    col = anchor_col + offset[1]
    if row < 1 or col < 1:
        return None
    return sheet.cells(row, col).value


def set_formula2_r1c1(cell: Any, formula: str) -> None:
    try:
        cell.formula2 = formula
    except Exception:
        cell.api.Formula2R1C1 = formula


def maybe_sheet(book: Any, name: str) -> Optional[Any]:
    try:
        return book.sheets[name]
    except Exception:
        return None


def count_contiguous_numeric_pairs(sheet: Any, end_row: int, x_col: int, y_col: int, max_points: int = 40) -> int:
    count = 0
    row = end_row
    while row >= 1 and count < max_points:
        x_val = to_number(sheet.cells(row, x_col).value)
        y_val = to_number(sheet.cells(row, y_col).value)
        if x_val is None or y_val is None:
            break
        count += 1
        row -= 1
    return count


def find_nearest_data_end_row(sheet: Any, anchor_row: int, x_col: int, y_col: int, search_span: int = 50) -> Optional[int]:
    nearest_row = None
    nearest_distance = None
    start_row = max(1, anchor_row - search_span)
    end_row = anchor_row + search_span
    for row in range(start_row, end_row + 1):
        x_val = to_number(sheet.cells(row, x_col).value)
        y_val = to_number(sheet.cells(row, y_col).value)
        if x_val is None or y_val is None:
            continue
        distance = abs(row - anchor_row)
        if nearest_distance is None or distance < nearest_distance:
            nearest_distance = distance
            nearest_row = row
    return nearest_row


def extract_empirical_rows(wb: Any, metadata: FileMetadata, source_file: str) -> List[Dict[str, Any]]:
    sheet = maybe_sheet(wb, EMPIRICAL_SHEET)
    if sheet is None:
        print(f"skipped: {source_file} (missing sheet: {EMPIRICAL_SHEET})")
        return []

    snapshot = build_sheet_snapshot(sheet)
    anchor = find_anchor(snapshot, "max")
    if anchor is None:
        print(f"skipped: {source_file} (no 'max' anchor in {EMPIRICAL_SHEET})")
        return []

    anchor_row, anchor_col = anchor
    avg_pen_input = sheet.cells(
        anchor_row + EMPIRICAL_OFFSETS["avg_penetration_input"][0],
        anchor_col + EMPIRICAL_OFFSETS["avg_penetration_input"][1],
    )
    temp_formula_cell = sheet.cells(
        anchor_row + EMPIRICAL_OFFSETS["temp_formula_cell"][0],
        anchor_col + EMPIRICAL_OFFSETS["temp_formula_cell"][1],
    )
    original_avg_pen_input = avg_pen_input.value

    forecast_max_label = nearest_label_position(snapshot, "max", anchor_row, anchor_col)
    forecast_min_label = nearest_label_position(snapshot, "min", anchor_row, anchor_col)

    rows: List[Dict[str, Any]] = []
    penetration_row = anchor_row + EMPIRICAL_OFFSETS["penetration_row"][0]
    quarter_label_row = anchor_row + EMPIRICAL_OFFSETS["quarter_label_row"][0]
    end_col = anchor_col - 1

    try:
        for num_quarters_used in range(1, MAX_QUARTERS + 1):
            start_col = end_col - (num_quarters_used - 1)
            if start_col < 1:
                break

            avg_formula = f"=AVERAGE(R{penetration_row}C{start_col}:R{penetration_row}C{end_col})"
            set_formula2_r1c1(temp_formula_cell, avg_formula)
            wb.app.calculate()
            avg_penetration_pct = to_number(temp_formula_cell.value)
            if avg_penetration_pct is None:
                continue

            avg_pen_input.value = avg_penetration_pct
            wb.app.calculate()

            forecast_value = to_number(read_offset(sheet, anchor_row, anchor_col, EMPIRICAL_OFFSETS["forecast_value"]))
            actual_value = to_number(read_offset(sheet, anchor_row, anchor_col, EMPIRICAL_OFFSETS["actual_value"]))
            forecast_max = to_number(read_offset(sheet, anchor_row, anchor_col, EMPIRICAL_OFFSETS["forecast_max"]))
            forecast_min = to_number(read_offset(sheet, anchor_row, anchor_col, EMPIRICAL_OFFSETS["forecast_min"]))

            if forecast_max is None and forecast_max_label:
                forecast_max = to_number(sheet.cells(forecast_max_label[0], forecast_max_label[1] + 1).value)
            if forecast_min is None and forecast_min_label:
                forecast_min = to_number(sheet.cells(forecast_min_label[0], forecast_min_label[1] + 1).value)

            quarterly_sales = to_number(read_offset(sheet, anchor_row, anchor_col, EMPIRICAL_OFFSETS["quarterly_sales"]))
            reported_sales = to_number(read_offset(sheet, anchor_row, anchor_col, EMPIRICAL_OFFSETS["reported_sales"]))
            growth_rate_pct = to_number(read_offset(sheet, anchor_row, anchor_col, EMPIRICAL_OFFSETS["growth_rate_pct"]))
            sales_captured_in_db_pct = to_number(
                read_offset(sheet, anchor_row, anchor_col, EMPIRICAL_OFFSETS["sales_captured_in_db_pct"])
            )

            last_quarter_used = sheet.cells(quarter_label_row, end_col).value if quarter_label_row >= 1 else None
            range_width = None
            if forecast_max is not None and forecast_min is not None:
                range_width = forecast_max - forecast_min

            rows.append(
                {
                    "model": metadata.model,
                    "ticker": metadata.ticker,
                    "model_period": metadata.model_period,
                    "model_date": metadata.model_date,
                    "method": "empirical",
                    "parameter_name": "avg_penetration_pct",
                    "parameter_value": avg_penetration_pct,
                    "num_quarters_used": num_quarters_used,
                    "last_quarter_used": last_quarter_used,
                    "forecast_value": forecast_value,
                    "actual_value": actual_value if actual_value is not None else reported_sales,
                    "forecast_max": forecast_max,
                    "forecast_min": forecast_min,
                    "range_width": range_width,
                    "avg_penetration_pct": avg_penetration_pct,
                    "quarterly_sales": quarterly_sales,
                    "reported_sales": reported_sales,
                    "growth_rate_pct": growth_rate_pct,
                    "sales_captured_in_db_pct": sales_captured_in_db_pct,
                    "source_file": source_file,
                }
            )
    finally:
        avg_pen_input.value = original_avg_pen_input
        temp_formula_cell.value = None
        wb.app.calculate()

    return rows


def extract_regression_rows(wb: Any, metadata: FileMetadata, source_file: str) -> List[Dict[str, Any]]:
    sheet = maybe_sheet(wb, REGRESSION_SHEET)
    if sheet is None:
        print(f"skipped: {source_file} (missing sheet: {REGRESSION_SHEET})")
        return []

    snapshot = build_sheet_snapshot(sheet)
    anchor = find_anchor(snapshot, "max")
    if anchor is None:
        print(f"skipped: {source_file} (no 'max' anchor in {REGRESSION_SHEET})")
        return []

    anchor_row, anchor_col = anchor
    y_col = anchor_col - 7
    x_col = anchor_col - 11
    if x_col < 1 or y_col < 1:
        print(f"skipped: {source_file} (invalid regression anchor offsets)")
        return []

    data_end_row = find_nearest_data_end_row(sheet, anchor_row, x_col, y_col)
    if data_end_row is None:
        print(f"skipped: {source_file} (no regression x/y data near anchor)")
        return []

    contiguous_points = count_contiguous_numeric_pairs(sheet, data_end_row, x_col, y_col, max_points=MAX_QUARTERS)
    if contiguous_points < 2:
        print(f"skipped: {source_file} (insufficient regression datapoints)")
        return []

    temp_intercept_cell = sheet.cells(
        anchor_row + REGRESSION_OFFSETS["temp_intercept_cell"][0],
        anchor_col + REGRESSION_OFFSETS["temp_intercept_cell"][1],
    )
    temp_slope_cell = sheet.cells(
        anchor_row + REGRESSION_OFFSETS["temp_slope_cell"][0],
        anchor_col + REGRESSION_OFFSETS["temp_slope_cell"][1],
    )

    forecast_max_label = nearest_label_position(snapshot, "max", anchor_row, anchor_col)
    forecast_min_label = nearest_label_position(snapshot, "min", anchor_row, anchor_col)

    rows: List[Dict[str, Any]] = []
    previous_signature: Optional[Tuple[Optional[float], ...]] = None
    max_q = min(MAX_QUARTERS, contiguous_points)

    try:
        for num_quarters_used in range(2, max_q + 1):
            start_row = data_end_row - num_quarters_used + 1
            intercept_formula = (
                f"=INTERCEPT(R{start_row}C{y_col}:R{data_end_row}C{y_col},"
                f"R{start_row}C{x_col}:R{data_end_row}C{x_col})"
            )
            slope_formula = (
                f"=SLOPE(R{start_row}C{y_col}:R{data_end_row}C{y_col},"
                f"R{start_row}C{x_col}:R{data_end_row}C{x_col})"
            )

            set_formula2_r1c1(temp_intercept_cell, intercept_formula)
            set_formula2_r1c1(temp_slope_cell, slope_formula)
            wb.app.calculate()

            intercept = to_number(temp_intercept_cell.value)
            slope = to_number(temp_slope_cell.value)

            forecast_value = to_number(
                read_offset(sheet, anchor_row, anchor_col, REGRESSION_OFFSETS["forecast_total_without_sa"])
            )
            actual_value = to_number(read_offset(sheet, anchor_row, anchor_col, REGRESSION_OFFSETS["actual_value"]))
            forecast_max = to_number(read_offset(sheet, anchor_row, anchor_col, REGRESSION_OFFSETS["forecast_max"]))
            forecast_min = to_number(read_offset(sheet, anchor_row, anchor_col, REGRESSION_OFFSETS["forecast_min"]))

            if forecast_max is None and forecast_max_label:
                forecast_max = to_number(sheet.cells(forecast_max_label[0], forecast_max_label[1] + 1).value)
            if forecast_min is None and forecast_min_label:
                forecast_min = to_number(sheet.cells(forecast_min_label[0], forecast_min_label[1] + 1).value)

            if forecast_value is None and intercept is not None and slope is not None:
                next_x = to_number(sheet.cells(data_end_row + 1, x_col).value)
                if next_x is None:
                    next_x = to_number(sheet.cells(data_end_row, x_col).value)
                if next_x is not None:
                    forecast_value = intercept + (slope * next_x)

            range_width = None
            if forecast_max is not None and forecast_min is not None:
                range_width = forecast_max - forecast_min

            signature = (
                round(forecast_value, 10) if forecast_value is not None else None,
                round(forecast_max, 10) if forecast_max is not None else None,
                round(forecast_min, 10) if forecast_min is not None else None,
                round(intercept, 10) if intercept is not None else None,
                round(slope, 10) if slope is not None else None,
            )
            if previous_signature == signature:
                continue
            previous_signature = signature

            rows.append(
                {
                    "model": metadata.model,
                    "ticker": metadata.ticker,
                    "model_period": metadata.model_period,
                    "model_date": metadata.model_date,
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
    finally:
        temp_intercept_cell.value = None
        temp_slope_cell.value = None
        wb.app.calculate()

    return rows


def write_rows_to_sheet(ws: Any, columns: Sequence[str], rows: Iterable[Dict[str, Any]]) -> None:
    ws.append(list(columns))
    for header_cell in ws[1]:
        header_cell.font = Font(bold=True)

    for row in rows:
        ws.append([row.get(column) for column in columns])

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(columns))}{ws.max_row}"

    for idx, column in enumerate(columns, start=1):
        max_len = len(column)
        for row_idx in range(2, ws.max_row + 1):
            value = ws.cell(row=row_idx, column=idx).value
            if value is None:
                continue
            max_len = max(max_len, len(str(value)))
        ws.column_dimensions[get_column_letter(idx)].width = min(max_len + 2, 48)


def write_output_workbook(output_path: Path, empirical_rows: List[Dict[str, Any]], regression_rows: List[Dict[str, Any]]) -> None:
    wb = Workbook()
    default_sheet = wb.active
    wb.remove(default_sheet)

    empirical_ws = wb.create_sheet("empirical_candidates")
    write_rows_to_sheet(empirical_ws, EMPIRICAL_COLUMNS, empirical_rows)

    regression_ws = wb.create_sheet("regression_candidates")
    write_rows_to_sheet(regression_ws, REGRESSION_COLUMNS, regression_rows)

    wb.save(output_path)


def iter_source_files(input_path: Path) -> Iterable[Path]:
    for file_path in sorted(input_path.iterdir()):
        if not file_path.is_file():
            print(f"skipped: {file_path.name} (not a file)")
            continue
        if file_path.name.startswith("~"):
            print(f"skipped: {file_path.name} (temporary file)")
            continue
        if file_path.suffix.lower() != ".xlsx":
            print(f"skipped: {file_path.name} (not .xlsx)")
            continue
        yield file_path


def main() -> None:
    input_path = input_dir.expanduser().resolve()
    output_path_root = output_dir.expanduser().resolve()

    if not input_path.exists():
        raise FileNotFoundError(f"input_dir does not exist: {input_path}")
    if not input_path.is_dir():
        raise NotADirectoryError(f"input_dir is not a directory: {input_path}")

    output_path = build_output_path(input_path, output_path_root)

    empirical_rows: List[Dict[str, Any]] = []
    regression_rows: List[Dict[str, Any]] = []
    processed_files = 0

    app = xw.App(visible=False, add_book=False)
    app.display_alerts = False
    app.screen_updating = False

    try:
        for file_path in iter_source_files(input_path):
            print(f"processing: {file_path.name}")
            wb = None
            try:
                wb = app.books.open(str(file_path), update_links=False)
            except Exception as exc:
                print(f"skipped: {file_path.name} (open failed: {exc})")
                continue

            try:
                metadata = parse_file_metadata(file_path)
                empirical_rows.extend(extract_empirical_rows(wb, metadata, file_path.name))
                regression_rows.extend(extract_regression_rows(wb, metadata, file_path.name))
                processed_files += 1
            except Exception as exc:
                print(f"skipped: {file_path.name} (processing failed: {exc})")
            finally:
                if wb is not None:
                    close_workbook_without_save(wb)
    finally:
        try:
            app.quit()
        except Exception:
            pass

    write_output_workbook(output_path, empirical_rows, regression_rows)

    print(f"output_path: {output_path}")
    print(f"files_processed: {processed_files}")
    print(f"empirical_rows: {len(empirical_rows)}")
    print(f"regression_rows: {len(regression_rows)}")


if __name__ == "__main__":
    main()
