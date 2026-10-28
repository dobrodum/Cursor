#!/usr/bin/env python3
from __future__ import annotations

import math
import re
from dataclasses import dataclass
from datetime import date
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Sequence, Tuple

import xlwings as xw
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

# ---------------------------------------------------------------------------
# Inputs
# ---------------------------------------------------------------------------
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

DAY_BY_PERIOD_PREFIX = {"early": 5, "mid": 15, "late": 25}
MONTH_BY_ABBR = {
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
PERIOD_TOKEN_RE = re.compile(
    r"^(Early|Mid|Late)(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)(\d{4})$",
    re.IGNORECASE,
)
PARAM_OUTPUT_RE = re.compile(r"_PARAM(?:\.\d+)?$", re.IGNORECASE)


@dataclass
class FileLabels:
    model: str
    ticker: str
    model_period: str
    model_date: str


@dataclass
class UsedRangeSnapshot:
    first_row: int
    first_col: int
    last_row: int
    last_col: int
    values: List[List[Any]]

    def value_at(self, row: int, col: int) -> Any:
        if row < self.first_row or row > self.last_row:
            return None
        if col < self.first_col or col > self.last_col:
            return None
        row_values = self.values[row - self.first_row]
        col_index = col - self.first_col
        if col_index >= len(row_values):
            return None
        return row_values[col_index]


def parse_file_labels(file_path: Path) -> FileLabels:
    stem = file_path.stem
    parts = [part.strip() for part in stem.split(" - ")]
    ticker = parts[1].upper() if len(parts) >= 2 and parts[1].strip() else "UNKNOWN"

    period_token = ""
    if len(parts) >= 3:
        period_token = parts[2].split("_")[0].strip()

    model_period = "UnknownPeriod"
    model_date = ""
    period_match = PERIOD_TOKEN_RE.match(period_token)
    if period_match:
        prefix = period_match.group(1).title()
        month_abbr = period_match.group(2).title()
        year = int(period_match.group(3))
        month = MONTH_BY_ABBR[month_abbr.lower()]
        day = DAY_BY_PERIOD_PREFIX[prefix.lower()]
        model_period = f"{prefix}{month_abbr}_{year}"
        model_date = date(year, month, day).isoformat()

    model = f"{ticker}_{model_period}"
    return FileLabels(
        model=model,
        ticker=ticker,
        model_period=model_period,
        model_date=model_date,
    )


def next_output_path(input_folder: Path, destination_dir: Path) -> Path:
    destination_dir.mkdir(parents=True, exist_ok=True)
    base_name = f"{input_folder.name}_PARAM"
    first_candidate = destination_dir / f"{base_name}.xlsx"
    if not first_candidate.exists():
        return first_candidate

    counter = 1
    while True:
        candidate = destination_dir / f"{base_name}.{counter}.xlsx"
        if not candidate.exists():
            return candidate
        counter += 1


def normalize_text(value: Any) -> str:
    if value is None:
        return ""
    return re.sub(r"[^a-z0-9]+", "", str(value).strip().lower())


def coerce_number(value: Any) -> Optional[float]:
    if value is None or value == "":
        return None
    if isinstance(value, bool):
        return None
    try:
        number = float(value)
    except (TypeError, ValueError):
        return None
    if math.isnan(number):
        return None
    return number


def as_2d(raw: Any, expected_rows: int, expected_cols: int) -> List[List[Any]]:
    if expected_rows <= 0 or expected_cols <= 0:
        return []

    if expected_rows == 1 and expected_cols == 1:
        return [[raw]]

    if expected_rows == 1:
        if isinstance(raw, Sequence) and not isinstance(raw, (str, bytes)):
            row = list(raw)
        else:
            row = [raw]
        if len(row) < expected_cols:
            row.extend([None] * (expected_cols - len(row)))
        return [row[:expected_cols]]

    if expected_cols == 1:
        if isinstance(raw, Sequence) and not isinstance(raw, (str, bytes)):
            values = list(raw)
        else:
            values = [raw]
        matrix: List[List[Any]] = []
        for idx in range(expected_rows):
            matrix.append([values[idx] if idx < len(values) else None])
        return matrix

    if not isinstance(raw, Sequence) or isinstance(raw, (str, bytes)):
        return [[raw if r == 0 and c == 0 else None for c in range(expected_cols)] for r in range(expected_rows)]

    matrix = []
    rows = list(raw)
    for r in range(expected_rows):
        row_value = rows[r] if r < len(rows) else []
        if isinstance(row_value, Sequence) and not isinstance(row_value, (str, bytes)):
            row = list(row_value)
        else:
            row = [row_value]
        if len(row) < expected_cols:
            row.extend([None] * (expected_cols - len(row)))
        matrix.append(row[:expected_cols])
    return matrix


def snapshot_used_range(sheet: xw.Sheet) -> UsedRangeSnapshot:
    used = sheet.used_range
    first_row = used.row
    first_col = used.column
    row_count = used.rows.count
    col_count = used.columns.count
    values = as_2d(used.value, row_count, col_count)
    return UsedRangeSnapshot(
        first_row=first_row,
        first_col=first_col,
        last_row=first_row + row_count - 1,
        last_col=first_col + col_count - 1,
        values=values,
    )


def find_max_anchor(snapshot: UsedRangeSnapshot) -> Optional[Tuple[int, int]]:
    candidates: List[Tuple[int, int, int]] = []
    for row_index, row in enumerate(snapshot.values, start=snapshot.first_row):
        for col_offset, value in enumerate(row):
            if normalize_text(value) != "max":
                continue
            col_index = snapshot.first_col + col_offset
            has_min_neighbor = False
            for delta in range(-6, 7):
                if delta == 0:
                    continue
                maybe_col = col_index + delta
                if normalize_text(snapshot.value_at(row_index, maybe_col)) == "min":
                    has_min_neighbor = True
                    break
            priority = 0 if has_min_neighbor else 1
            candidates.append((priority, row_index, col_index))

    if not candidates:
        return None

    candidates.sort(key=lambda item: (item[0], item[1], item[2]))
    _, row, col = candidates[0]
    return row, col


def header_text_by_column(snapshot: UsedRangeSnapshot, header_row: int) -> Dict[int, str]:
    headers: Dict[int, str] = {}
    for col in range(snapshot.first_col, snapshot.last_col + 1):
        text = normalize_text(snapshot.value_at(header_row, col))
        if text:
            headers[col] = text
    return headers


def resolve_column(
    headers: Dict[int, str],
    exact_aliases: Iterable[str],
    contains_aliases: Iterable[Tuple[str, ...]],
    fallback_col: int,
) -> int:
    exact_set = {alias for alias in exact_aliases if alias}
    for col, header in sorted(headers.items()):
        if header in exact_set:
            return col

    for col, header in sorted(headers.items()):
        for keywords in contains_aliases:
            if keywords and all(keyword in header for keyword in keywords):
                return col

    return max(1, fallback_col)


def read_cell(sheet: xw.Sheet, snapshot: UsedRangeSnapshot, row: int, col: int) -> Any:
    col = max(1, col)
    value = snapshot.value_at(row, col)
    if row < snapshot.first_row or row > snapshot.last_row:
        return sheet.range((row, col)).value
    if col < snapshot.first_col or col > snapshot.last_col:
        return sheet.range((row, col)).value
    return value


def set_formula2(cell: xw.Range, formula: str) -> None:
    try:
        cell.formula2 = formula
    except Exception:
        cell.formula = formula


def safe_subtract(left: Optional[float], right: Optional[float]) -> Optional[float]:
    if left is None or right is None:
        return None
    return left - right


def close_source_workbook(workbook: xw.Book) -> None:
    try:
        workbook.close(save=False)
        return
    except TypeError:
        pass
    except Exception:
        pass

    try:
        workbook.api.Close(SaveChanges=False)
        return
    except Exception:
        pass

    try:
        workbook.close()
    except Exception:
        pass


def process_empirical_sheet(
    workbook: xw.Book,
    source_file: Path,
    labels: FileLabels,
) -> List[Dict[str, Any]]:
    try:
        sheet = workbook.sheets["Empirical Model"]
    except Exception:
        print(f"Skipped sheet in {source_file.name}: 'Empirical Model' not found")
        return []

    snapshot = snapshot_used_range(sheet)
    anchor = find_max_anchor(snapshot)
    if not anchor:
        print(f"Skipped sheet in {source_file.name}: 'max' anchor not found in 'Empirical Model'")
        return []

    anchor_row, anchor_col = anchor
    headers = header_text_by_column(snapshot, anchor_row)

    forecast_max_col = resolve_column(headers, ["max"], [("forecast", "max")], anchor_col)
    forecast_min_col = resolve_column(headers, ["min"], [("forecast", "min")], anchor_col + 1)
    forecast_value_col = resolve_column(
        headers,
        ["estimatedtotalsold", "totfcst", "totalforecast", "forecastvalue"],
        [("estimated", "total", "sold"), ("tot", "fcst"), ("forecast", "value")],
        anchor_col - 1,
    )
    actual_value_col = resolve_column(
        headers,
        ["reportedsales", "actualsales", "actualvalue", "actual"],
        [("reported", "sales"), ("actual", "sales"), ("actual", "value")],
        anchor_col - 2,
    )
    num_quarters_col = resolve_column(
        headers,
        ["numquartersused", "nquarters", "quartersused", "numqtrs"],
        [("num", "quarter"), ("quarters", "used"), ("n", "q")],
        anchor_col - 8,
    )
    last_quarter_col = resolve_column(
        headers,
        ["lastquarterused", "lastquarter"],
        [("last", "quarter", "used"), ("last", "quarter")],
        anchor_col - 9,
    )
    avg_penetration_col = resolve_column(
        headers,
        ["avgpenetrationpct", "averagepenetration", "avgpenetration"],
        [("avg", "penetration"), ("average", "penetration")],
        anchor_col - 7,
    )
    quarterly_sales_col = resolve_column(
        headers,
        ["quarterlysales"],
        [("quarterly", "sales")],
        anchor_col - 5,
    )
    reported_sales_col = resolve_column(
        headers,
        ["reportedsales"],
        [("reported", "sales")],
        anchor_col - 4,
    )
    growth_rate_col = resolve_column(
        headers,
        ["growthratepct", "growthrate"],
        [("growth", "rate")],
        anchor_col - 3,
    )
    sales_captured_col = resolve_column(
        headers,
        ["salescapturedindbpct", "salescapturedindb"],
        [("sales", "captured", "db"), ("captured", "db")],
        anchor_col - 6,
    )

    # Penetration history source for R1C1 average formula.
    penetration_history_col = resolve_column(
        headers,
        ["penetrationpct", "penetration"],
        [("penetration",)],
        anchor_col - 10,
    )

    helper_row = snapshot.last_row + 2
    helper_col = snapshot.last_col + 2
    helper_cell = sheet.range((helper_row, helper_col))

    output_rows: List[Dict[str, Any]] = []
    data_end_row = anchor_row - 1
    empty_streak = 0

    for n_quarters in range(1, 11):
        result_row = anchor_row + n_quarters
        num_quarters = coerce_number(read_cell(sheet, snapshot, result_row, num_quarters_col))
        num_quarters_used = int(num_quarters) if num_quarters is not None and num_quarters > 0 else n_quarters

        avg_penetration_pct = coerce_number(read_cell(sheet, snapshot, result_row, avg_penetration_col))
        if avg_penetration_pct is None and data_end_row >= snapshot.first_row:
            start_row = max(snapshot.first_row, data_end_row - num_quarters_used + 1)
            if start_row <= data_end_row:
                average_formula = (
                    f"=AVERAGE(R{start_row}C{max(1, penetration_history_col)}:"
                    f"R{data_end_row}C{max(1, penetration_history_col)})"
                )
                set_formula2(helper_cell, average_formula)
                workbook.app.calculate()
                avg_penetration_pct = coerce_number(helper_cell.value)

        forecast_value = coerce_number(read_cell(sheet, snapshot, result_row, forecast_value_col))
        actual_value = coerce_number(read_cell(sheet, snapshot, result_row, actual_value_col))
        forecast_max = coerce_number(read_cell(sheet, snapshot, result_row, forecast_max_col))
        forecast_min = coerce_number(read_cell(sheet, snapshot, result_row, forecast_min_col))
        quarterly_sales = coerce_number(read_cell(sheet, snapshot, result_row, quarterly_sales_col))
        reported_sales = coerce_number(read_cell(sheet, snapshot, result_row, reported_sales_col))
        growth_rate_pct = coerce_number(read_cell(sheet, snapshot, result_row, growth_rate_col))
        sales_captured_in_db_pct = coerce_number(read_cell(sheet, snapshot, result_row, sales_captured_col))
        last_quarter_used = read_cell(sheet, snapshot, result_row, last_quarter_col)

        if last_quarter_used in ("", None) and data_end_row >= snapshot.first_row:
            history_row = max(snapshot.first_row, data_end_row - num_quarters_used + 1)
            last_quarter_used = read_cell(sheet, snapshot, history_row, max(1, last_quarter_col))

        if all(
            value is None
            for value in [forecast_value, forecast_max, forecast_min, actual_value, avg_penetration_pct]
        ):
            empty_streak += 1
            if empty_streak >= 2:
                break
            continue
        empty_streak = 0

        output_rows.append(
            {
                "model": labels.model,
                "ticker": labels.ticker,
                "model_period": labels.model_period,
                "model_date": labels.model_date,
                "method": "empirical",
                "parameter_name": "avg_penetration_pct",
                "parameter_value": avg_penetration_pct,
                "num_quarters_used": num_quarters_used,
                "last_quarter_used": last_quarter_used,
                "forecast_value": forecast_value,
                "actual_value": actual_value,
                "forecast_max": forecast_max,
                "forecast_min": forecast_min,
                "range_width": safe_subtract(forecast_max, forecast_min),
                "avg_penetration_pct": avg_penetration_pct,
                "quarterly_sales": quarterly_sales,
                "reported_sales": reported_sales,
                "growth_rate_pct": growth_rate_pct,
                "sales_captured_in_db_pct": sales_captured_in_db_pct,
                "source_file": source_file.name,
            }
        )

    return output_rows


def values_equal(left: Any, right: Any) -> bool:
    left_num = coerce_number(left)
    right_num = coerce_number(right)
    if left_num is not None or right_num is not None:
        if left_num is None or right_num is None:
            return False
        return math.isclose(left_num, right_num, rel_tol=1e-9, abs_tol=1e-9)
    return left == right


def is_duplicate_regression_row(previous_row: Dict[str, Any], current_row: Dict[str, Any]) -> bool:
    keys = [
        "num_quarters_used",
        "forecast_value",
        "forecast_max",
        "forecast_min",
        "intercept",
        "slope",
    ]
    return all(values_equal(previous_row.get(key), current_row.get(key)) for key in keys)


def process_regression_sheet(
    workbook: xw.Book,
    source_file: Path,
    labels: FileLabels,
) -> List[Dict[str, Any]]:
    try:
        sheet = workbook.sheets["Regression Model"]
    except Exception:
        print(f"Skipped sheet in {source_file.name}: 'Regression Model' not found")
        return []

    snapshot = snapshot_used_range(sheet)
    anchor = find_max_anchor(snapshot)
    if not anchor:
        print(f"Skipped sheet in {source_file.name}: 'max' anchor not found in 'Regression Model'")
        return []

    anchor_row, anchor_col = anchor
    headers = header_text_by_column(snapshot, anchor_row)

    num_quarters_col = resolve_column(
        headers,
        ["numquartersused", "nquarters", "quartersused", "numqtrs"],
        [("num", "quarter"), ("quarters", "used"), ("n", "q")],
        anchor_col - 6,
    )
    forecast_value_col = resolve_column(
        headers,
        ["totfcstwosa", "forecastvalue", "totalforecast", "totfcst"],
        [("tot", "fcst", "wosa"), ("tot", "fcst"), ("forecast", "value")],
        anchor_col - 1,
    )
    actual_value_col = resolve_column(
        headers,
        ["actualvalue", "actualsales", "reportedsales", "actual"],
        [("actual",), ("reported", "sales")],
        anchor_col - 2,
    )
    forecast_max_col = resolve_column(headers, ["max"], [("forecast", "max")], anchor_col)
    forecast_min_col = resolve_column(headers, ["min"], [("forecast", "min")], anchor_col + 1)

    y_col = max(1, anchor_col - 7)
    x_col = max(1, anchor_col - 11)

    helper_row = snapshot.last_row + 2
    helper_col = snapshot.last_col + 2
    intercept_cell = sheet.range((helper_row, helper_col))
    slope_cell = sheet.range((helper_row, helper_col + 1))

    data_end_row = anchor_row - 1
    if data_end_row < snapshot.first_row:
        return []

    forecast_x = coerce_number(read_cell(sheet, snapshot, anchor_row, x_col))
    if forecast_x is None:
        forecast_x = coerce_number(read_cell(sheet, snapshot, data_end_row, x_col))

    output_rows: List[Dict[str, Any]] = []
    empty_streak = 0

    for n_quarters in range(1, 11):
        result_row = anchor_row + n_quarters
        num_quarters_value = coerce_number(read_cell(sheet, snapshot, result_row, num_quarters_col))
        num_quarters_used = int(num_quarters_value) if num_quarters_value is not None and num_quarters_value > 0 else n_quarters
        start_row = max(snapshot.first_row, data_end_row - num_quarters_used + 1)
        if start_row > data_end_row:
            continue

        intercept_formula = (
            f"=INTERCEPT(R{start_row}C{y_col}:R{data_end_row}C{y_col},"
            f"R{start_row}C{x_col}:R{data_end_row}C{x_col})"
        )
        slope_formula = (
            f"=SLOPE(R{start_row}C{y_col}:R{data_end_row}C{y_col},"
            f"R{start_row}C{x_col}:R{data_end_row}C{x_col})"
        )
        set_formula2(intercept_cell, intercept_formula)
        set_formula2(slope_cell, slope_formula)
        workbook.app.calculate()

        intercept = coerce_number(intercept_cell.value)
        slope = coerce_number(slope_cell.value)
        forecast_value = coerce_number(read_cell(sheet, snapshot, result_row, forecast_value_col))
        if forecast_value is None and intercept is not None and slope is not None and forecast_x is not None:
            forecast_value = intercept + slope * forecast_x

        forecast_max = coerce_number(read_cell(sheet, snapshot, result_row, forecast_max_col))
        forecast_min = coerce_number(read_cell(sheet, snapshot, result_row, forecast_min_col))
        actual_value = coerce_number(read_cell(sheet, snapshot, result_row, actual_value_col))

        if all(value is None for value in [forecast_value, forecast_max, forecast_min, intercept, slope]):
            empty_streak += 1
            if empty_streak >= 2:
                break
            continue
        empty_streak = 0

        row_data = {
            "model": labels.model,
            "ticker": labels.ticker,
            "model_period": labels.model_period,
            "model_date": labels.model_date,
            "method": "regression",
            "parameter_name": "num_quarters_used",
            "parameter_value": num_quarters_used,
            "num_quarters_used": num_quarters_used,
            "forecast_value": forecast_value,
            "actual_value": actual_value,
            "forecast_max": forecast_max,
            "forecast_min": forecast_min,
            "range_width": safe_subtract(forecast_max, forecast_min),
            "intercept": intercept,
            "slope": slope,
            "source_file": source_file.name,
        }

        if output_rows and is_duplicate_regression_row(output_rows[-1], row_data):
            continue

        output_rows.append(row_data)

    return output_rows


def write_rows_to_sheet(sheet: Any, columns: List[str], rows: List[Dict[str, Any]]) -> None:
    sheet.append(columns)
    for row in rows:
        sheet.append([row.get(column) for column in columns])

    for header_cell in sheet[1]:
        header_cell.font = Font(bold=True)

    sheet.freeze_panes = "A2"
    sheet.auto_filter.ref = sheet.dimensions

    for column_index, column_name in enumerate(columns, start=1):
        max_length = len(column_name)
        for row_index in range(2, sheet.max_row + 1):
            value = sheet.cell(row=row_index, column=column_index).value
            if value is None:
                continue
            text = str(value)
            if len(text) > max_length:
                max_length = len(text)
        sheet.column_dimensions[get_column_letter(column_index)].width = min(max(max_length + 2, 12), 48)


def write_output_workbook(
    output_path: Path,
    empirical_rows: List[Dict[str, Any]],
    regression_rows: List[Dict[str, Any]],
) -> None:
    workbook = Workbook()
    empirical_sheet = workbook.active
    empirical_sheet.title = "empirical_candidates"
    regression_sheet = workbook.create_sheet("regression_candidates")

    write_rows_to_sheet(empirical_sheet, EMPIRICAL_COLUMNS, empirical_rows)
    write_rows_to_sheet(regression_sheet, REGRESSION_COLUMNS, regression_rows)

    workbook.save(output_path)
    workbook.close()


def list_input_files(folder: Path) -> List[Path]:
    input_files: List[Path] = []
    if not folder.exists():
        print(f"Input directory does not exist: {folder}")
        return input_files

    for file_path in sorted(folder.iterdir(), key=lambda path: path.name.lower()):
        if not file_path.is_file():
            continue
        if file_path.name.startswith("~"):
            print(f"Skipped {file_path.name}: temporary file")
            continue
        if file_path.suffix.lower() != ".xlsx":
            print(f"Skipped {file_path.name}: not an .xlsx file")
            continue
        if PARAM_OUTPUT_RE.search(file_path.stem):
            print(f"Skipped {file_path.name}: appears to be a previous PARAM output")
            continue
        input_files.append(file_path)
    return input_files


def process_workbooks(files: List[Path]) -> Tuple[int, List[Dict[str, Any]], List[Dict[str, Any]]]:
    processed_files = 0
    empirical_rows: List[Dict[str, Any]] = []
    regression_rows: List[Dict[str, Any]] = []

    if not files:
        return processed_files, empirical_rows, regression_rows

    app = xw.App(visible=False, add_book=False)
    app.display_alerts = False
    app.screen_updating = False

    previous_calculation_mode = None
    try:
        try:
            previous_calculation_mode = app.calculation
            app.calculation = "manual"
        except Exception:
            previous_calculation_mode = None

        for file_path in files:
            print(f"Processing {file_path.name}")
            workbook = None
            try:
                workbook = app.books.open(str(file_path), update_links=False)
                labels = parse_file_labels(file_path)
                empirical_rows.extend(process_empirical_sheet(workbook, file_path, labels))
                regression_rows.extend(process_regression_sheet(workbook, file_path, labels))
                processed_files += 1
            except Exception as error:
                print(f"Skipped {file_path.name}: failed to process ({error})")
            finally:
                if workbook is not None:
                    close_source_workbook(workbook)
    finally:
        if previous_calculation_mode is not None:
            try:
                app.calculation = previous_calculation_mode
            except Exception:
                pass
        app.quit()

    return processed_files, empirical_rows, regression_rows


def main() -> None:
    files = list_input_files(input_dir)
    output_path = next_output_path(input_dir, output_dir)

    processed_files, empirical_rows, regression_rows = process_workbooks(files)
    write_output_workbook(output_path, empirical_rows, regression_rows)

    print(f"Output written to: {output_path}")
    print(f"Files processed: {processed_files}")
    print(f"Empirical rows: {len(empirical_rows)}")
    print(f"Regression rows: {len(regression_rows)}")


if __name__ == "__main__":
    main()
