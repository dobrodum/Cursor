#!/usr/bin/env python3
"""Extract empirical/regression candidates from .xlsx model workbooks in one pass."""

from __future__ import annotations

import re
from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Sequence, Tuple

import xlwings as xw
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

# ---------------------------------------------------------------------------
# User-configurable paths
# ---------------------------------------------------------------------------
input_dir = "/path/to/input"
output_dir = "/path/to/output"

# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------
EMPIRICAL_SHEET_NAME = "Empirical Model"
REGRESSION_SHEET_NAME = "Regression Model"
N_QUARTERS = 10

# Anchor-relative fallback offsets from the "max" column.
EMPIRICAL_OFFSETS = {
    "num_quarters_used": -7,
    "last_quarter_used": -6,
    "quarterly_sales": -5,
    "reported_sales": -4,
    "growth_rate_pct": -3,
    "sales_captured_in_db_pct": -2,
    "forecast_value": -1,  # estimated total sold
    "forecast_max": 0,
    "forecast_min": 1,
    "penetration_pct": -5,
}

REGRESSION_OFFSETS = {
    "num_quarters_used": -3,
    "actual_value": -2,
    "forecast_value": -1,  # TOT FCST w/o SA
    "forecast_max": 0,
    "forecast_min": 1,
}

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

HEADER_ALIASES = {
    "num_quarters_used": [
        "num quarters used",
        "quarters used",
        "num qtrs used",
        "num qtrs",
        "n quarters",
        "n qtrs",
    ],
    "last_quarter_used": ["last quarter used", "last qtr used", "last quarter", "quarter"],
    "quarterly_sales": ["quarterly sales", "quarter sales", "qtr sales", "sales"],
    "reported_sales": ["reported sales", "reported", "actual sales", "actual"],
    "growth_rate_pct": ["growth rate pct", "growth rate", "growth %", "growth pct"],
    "sales_captured_in_db_pct": [
        "sales captured in db pct",
        "sales captured in db",
        "captured in db pct",
        "captured in db",
        "db capture pct",
    ],
    "forecast_value": [
        "estimated total sold",
        "est total sold",
        "tot fcst w o sa",
        "tot fcst wo sa",
        "tot fcst w/o sa",
        "total forecast",
        "forecast value",
        "forecast",
    ],
    "actual_value": ["actual value", "actual", "reported sales"],
    "forecast_max": ["max"],
    "forecast_min": ["min"],
    "penetration_pct": ["penetration pct", "penetration %", "penetration"],
}

DAY_BY_PREFIX = {"early": 5, "mid": 15, "late": 25}
PERIOD_RE = re.compile(r"(Early|Mid|Late)([A-Za-z]{3})(\d{4})", re.IGNORECASE)


@dataclass
class FileLabels:
    model: str
    ticker: str
    model_period: str
    model_date: Optional[str]


@dataclass
class SheetGrid:
    values: List[List[Any]]
    row_start: int
    col_start: int
    row_end: int
    col_end: int


def normalize_text(value: Any) -> str:
    if value is None:
        return ""
    text = str(value).strip().lower().replace("\n", " ")
    text = re.sub(r"[^a-z0-9%]+", " ", text)
    return re.sub(r"\s+", " ", text).strip()


def to_2d(values: Any) -> List[List[Any]]:
    if values is None:
        return []
    if isinstance(values, tuple):
        values = list(values)
    if not isinstance(values, list):
        return [[values]]
    if not values:
        return []
    if isinstance(values[0], tuple):
        values = [list(row) if isinstance(row, tuple) else row for row in values]
    if isinstance(values[0], list):
        rows = [list(row) if isinstance(row, list) else [row] for row in values]
        width = max((len(row) for row in rows), default=0)
        return [row + [None] * (width - len(row)) for row in rows]
    return [list(values)]


def get_sheet_grid(sheet: xw.Sheet) -> Optional[SheetGrid]:
    used = sheet.used_range
    values = to_2d(used.value)
    if not values:
        return None
    row_start = used.row
    col_start = used.column
    row_end = row_start + len(values) - 1
    col_end = col_start + len(values[0]) - 1
    return SheetGrid(values=values, row_start=row_start, col_start=col_start, row_end=row_end, col_end=col_end)


def find_anchor_cell(grid: SheetGrid, anchor_text: str = "max") -> Optional[Tuple[int, int]]:
    target = normalize_text(anchor_text)
    for r_idx, row in enumerate(grid.values):
        for c_idx, value in enumerate(row):
            if normalize_text(value) == target:
                return grid.row_start + r_idx, grid.col_start + c_idx
    return None


def build_header_map(grid: SheetGrid, row_candidates: Sequence[int]) -> Dict[str, int]:
    header_map: Dict[str, int] = {}
    for sheet_row in row_candidates:
        row_idx = sheet_row - grid.row_start
        if row_idx < 0 or row_idx >= len(grid.values):
            continue
        for c_idx, value in enumerate(grid.values[row_idx]):
            key = normalize_text(value)
            if key:
                header_map.setdefault(key, grid.col_start + c_idx)
    return header_map


def resolve_column(
    header_map: Dict[str, int],
    aliases: Iterable[str],
    anchor_col: int,
    fallback_offset: int,
) -> int:
    normalized_aliases = [normalize_text(alias) for alias in aliases]
    for alias in normalized_aliases:
        if alias in header_map:
            return header_map[alias]
    for existing_key, col in header_map.items():
        for alias in normalized_aliases:
            if alias and alias in existing_key:
                return col
    return anchor_col + fallback_offset


def parse_file_labels(file_path: Path) -> FileLabels:
    stem = file_path.stem
    parts = [part.strip() for part in stem.split(" - ") if part.strip()]
    ticker = parts[1].upper() if len(parts) >= 2 else "UNKNOWN"
    period_source = parts[2] if len(parts) >= 3 else stem

    model_period = period_source
    model_date: Optional[str] = None

    period_match = PERIOD_RE.search(period_source.replace("_", ""))
    if period_match:
        prefix_raw, month_raw, year_raw = period_match.groups()
        prefix = prefix_raw.title()
        month_abbrev = month_raw.title()
        year = int(year_raw)
        day = DAY_BY_PREFIX[prefix_raw.lower()]
        try:
            month = datetime.strptime(month_abbrev, "%b").month
            model_date = date(year, month, day).isoformat()
            model_period = f"{prefix}{month_abbrev}_{year}"
        except ValueError:
            model_date = None

    model = f"{ticker}_{model_period}"
    return FileLabels(model=model, ticker=ticker, model_period=model_period, model_date=model_date)


def as_float(value: Any) -> Optional[float]:
    if value in (None, ""):
        return None
    if isinstance(value, bool):
        return float(value)
    if isinstance(value, (int, float)):
        return float(value)
    text = str(value).strip().replace(",", "")
    if not text:
        return None
    if text.endswith("%"):
        text = text[:-1]
        try:
            return float(text) / 100.0
        except ValueError:
            return None
    try:
        return float(text)
    except ValueError:
        return None


def as_int(value: Any, fallback: Optional[int] = None) -> Optional[int]:
    number = as_float(value)
    if number is None:
        return fallback
    return int(round(number))


def subtract_values(left: Optional[float], right: Optional[float]) -> Optional[float]:
    if left is None or right is None:
        return None
    return left - right


def get_sheet_case_insensitive(workbook: xw.Book, target_name: str) -> Optional[xw.Sheet]:
    target = target_name.strip().lower()
    for sheet in workbook.sheets:
        if sheet.name.strip().lower() == target:
            return sheet
    return None


def safe_cell_value(sheet: xw.Sheet, row: int, col: int) -> Any:
    if row < 1 or col < 1:
        return None
    try:
        return sheet.cells(row, col).value
    except Exception:
        return None


def safe_set_formula2_r1c1(cell: Any, formula_r1c1: str) -> None:
    # Primary write path requested by prompt.
    try:
        cell.formula2 = formula_r1c1
        return
    except Exception:
        pass
    # Safe fallback for Excel variants where Formula2R1C1 is exposed only via API.
    try:
        cell.api.Formula2R1C1 = formula_r1c1
        return
    except Exception:
        pass
    cell.formula = formula_r1c1


def close_workbook_safely(workbook: xw.Book) -> None:
    try:
        workbook.close(save=False)
        return
    except TypeError:
        pass
    except Exception:
        pass

    try:
        workbook.close(False)
        return
    except Exception:
        pass

    try:
        workbook.api.Close(SaveChanges=False)
    except Exception:
        pass


def approx_equal(a: Any, b: Any, tol: float = 1e-10) -> bool:
    a_float = as_float(a)
    b_float = as_float(b)
    if a_float is not None and b_float is not None:
        return abs(a_float - b_float) <= tol
    return a == b


def rows_are_duplicate(prev_row: Dict[str, Any], curr_row: Dict[str, Any], fields: Sequence[str]) -> bool:
    return all(approx_equal(prev_row.get(field), curr_row.get(field)) for field in fields)


def extract_empirical_candidates(workbook: xw.Book, labels: FileLabels, source_file: str) -> List[Dict[str, Any]]:
    sheet = get_sheet_case_insensitive(workbook, EMPIRICAL_SHEET_NAME)
    if sheet is None:
        return []

    grid = get_sheet_grid(sheet)
    if grid is None:
        return []

    anchor = find_anchor_cell(grid, "max")
    if anchor is None:
        return []
    anchor_row, anchor_col = anchor

    header_map = build_header_map(grid, [anchor_row - 1, anchor_row, anchor_row + 1])

    col_num_q = resolve_column(
        header_map,
        HEADER_ALIASES["num_quarters_used"],
        anchor_col,
        EMPIRICAL_OFFSETS["num_quarters_used"],
    )
    col_last_q = resolve_column(
        header_map,
        HEADER_ALIASES["last_quarter_used"],
        anchor_col,
        EMPIRICAL_OFFSETS["last_quarter_used"],
    )
    col_quarterly_sales = resolve_column(
        header_map,
        HEADER_ALIASES["quarterly_sales"],
        anchor_col,
        EMPIRICAL_OFFSETS["quarterly_sales"],
    )
    col_reported_sales = resolve_column(
        header_map,
        HEADER_ALIASES["reported_sales"],
        anchor_col,
        EMPIRICAL_OFFSETS["reported_sales"],
    )
    col_growth_rate = resolve_column(
        header_map,
        HEADER_ALIASES["growth_rate_pct"],
        anchor_col,
        EMPIRICAL_OFFSETS["growth_rate_pct"],
    )
    col_sales_captured = resolve_column(
        header_map,
        HEADER_ALIASES["sales_captured_in_db_pct"],
        anchor_col,
        EMPIRICAL_OFFSETS["sales_captured_in_db_pct"],
    )
    col_forecast_value = resolve_column(
        header_map,
        HEADER_ALIASES["forecast_value"],
        anchor_col,
        EMPIRICAL_OFFSETS["forecast_value"],
    )
    col_forecast_max = resolve_column(
        header_map,
        HEADER_ALIASES["forecast_max"],
        anchor_col,
        EMPIRICAL_OFFSETS["forecast_max"],
    )
    col_forecast_min = resolve_column(
        header_map,
        HEADER_ALIASES["forecast_min"],
        anchor_col,
        EMPIRICAL_OFFSETS["forecast_min"],
    )
    col_penetration = resolve_column(
        header_map,
        HEADER_ALIASES["penetration_pct"],
        anchor_col,
        EMPIRICAL_OFFSETS["penetration_pct"],
    )

    start_row = anchor_row + 1
    end_row = anchor_row + N_QUARTERS
    row_numbers = list(range(start_row, end_row + 1))

    scratch_col = max(grid.col_end, col_forecast_min, col_forecast_max) + 2
    avg_penetration_values: Dict[int, Optional[float]] = {}

    for idx, row in enumerate(row_numbers):
        num_quarters = as_int(safe_cell_value(sheet, row, col_num_q), fallback=idx + 1) or (idx + 1)
        num_quarters = max(1, min(num_quarters, N_QUARTERS))

        rel_col = col_penetration - scratch_col
        if num_quarters == 1:
            formula = f"=RC[{rel_col}]"
        else:
            formula = f"=AVERAGE(R[-{num_quarters - 1}]C[{rel_col}]:RC[{rel_col}])"

        safe_set_formula2_r1c1(sheet.cells(row, scratch_col), formula)

    workbook.app.calculate()

    avg_pen_values = sheet.range((start_row, scratch_col), (end_row, scratch_col)).value
    if not isinstance(avg_pen_values, list):
        avg_pen_values = [avg_pen_values]
    for idx, row in enumerate(row_numbers):
        value = avg_pen_values[idx] if idx < len(avg_pen_values) else None
        avg_penetration_values[row] = as_float(value)

    sheet.range((start_row, scratch_col), (end_row, scratch_col)).value = None

    rows: List[Dict[str, Any]] = []
    for idx, row in enumerate(row_numbers):
        num_quarters = as_int(safe_cell_value(sheet, row, col_num_q), fallback=idx + 1) or (idx + 1)
        last_quarter_used = safe_cell_value(sheet, row, col_last_q)
        quarterly_sales = as_float(safe_cell_value(sheet, row, col_quarterly_sales))
        reported_sales = as_float(safe_cell_value(sheet, row, col_reported_sales))
        growth_rate_pct = as_float(safe_cell_value(sheet, row, col_growth_rate))
        sales_captured_in_db_pct = as_float(safe_cell_value(sheet, row, col_sales_captured))
        forecast_value = as_float(safe_cell_value(sheet, row, col_forecast_value))
        forecast_max = as_float(safe_cell_value(sheet, row, col_forecast_max))
        forecast_min = as_float(safe_cell_value(sheet, row, col_forecast_min))
        avg_penetration_pct = avg_penetration_values.get(row)
        if avg_penetration_pct is None:
            avg_penetration_pct = as_float(safe_cell_value(sheet, row, col_penetration))

        if all(
            value is None
            for value in (
                num_quarters,
                forecast_value,
                forecast_max,
                forecast_min,
                reported_sales,
                quarterly_sales,
            )
        ):
            continue

        rows.append(
            {
                "model": labels.model,
                "ticker": labels.ticker,
                "model_period": labels.model_period,
                "model_date": labels.model_date,
                "method": "empirical",
                "parameter_name": "avg_penetration_pct",
                "parameter_value": avg_penetration_pct,
                "num_quarters_used": num_quarters,
                "last_quarter_used": last_quarter_used,
                "forecast_value": forecast_value,
                "actual_value": reported_sales,
                "forecast_max": forecast_max,
                "forecast_min": forecast_min,
                "range_width": subtract_values(forecast_max, forecast_min),
                "avg_penetration_pct": avg_penetration_pct,
                "quarterly_sales": quarterly_sales,
                "reported_sales": reported_sales,
                "growth_rate_pct": growth_rate_pct,
                "sales_captured_in_db_pct": sales_captured_in_db_pct,
                "source_file": source_file,
            }
        )

    return rows


def extract_regression_candidates(workbook: xw.Book, labels: FileLabels, source_file: str) -> List[Dict[str, Any]]:
    sheet = get_sheet_case_insensitive(workbook, REGRESSION_SHEET_NAME)
    if sheet is None:
        return []

    grid = get_sheet_grid(sheet)
    if grid is None:
        return []

    anchor = find_anchor_cell(grid, "max")
    if anchor is None:
        return []
    anchor_row, anchor_col = anchor

    header_map = build_header_map(grid, [anchor_row - 1, anchor_row, anchor_row + 1])

    col_num_q = resolve_column(
        header_map,
        HEADER_ALIASES["num_quarters_used"],
        anchor_col,
        REGRESSION_OFFSETS["num_quarters_used"],
    )
    col_actual = resolve_column(
        header_map,
        HEADER_ALIASES["actual_value"],
        anchor_col,
        REGRESSION_OFFSETS["actual_value"],
    )
    col_forecast_value = resolve_column(
        header_map,
        HEADER_ALIASES["forecast_value"],
        anchor_col,
        REGRESSION_OFFSETS["forecast_value"],
    )
    col_forecast_max = resolve_column(
        header_map,
        HEADER_ALIASES["forecast_max"],
        anchor_col,
        REGRESSION_OFFSETS["forecast_max"],
    )
    col_forecast_min = resolve_column(
        header_map,
        HEADER_ALIASES["forecast_min"],
        anchor_col,
        REGRESSION_OFFSETS["forecast_min"],
    )

    y_col = anchor_col - 7
    x_col = anchor_col - 11
    if y_col < 1 or x_col < 1:
        return []

    start_row = anchor_row + 1
    end_row = anchor_row + N_QUARTERS
    row_numbers = list(range(start_row, end_row + 1))

    intercept_col = max(grid.col_end, col_forecast_min, col_forecast_max) + 2
    slope_col = intercept_col + 1

    for idx, row in enumerate(row_numbers):
        num_quarters = as_int(safe_cell_value(sheet, row, col_num_q), fallback=idx + 1) or (idx + 1)
        num_quarters = max(1, min(num_quarters, N_QUARTERS))

        y_rel = y_col - intercept_col
        x_rel = x_col - intercept_col

        if num_quarters == 1:
            intercept_formula = f"=RC[{y_rel}]"
            slope_formula = "=0"
        else:
            intercept_formula = (
                f"=INTERCEPT(R[-{num_quarters - 1}]C[{y_rel}]:RC[{y_rel}],"
                f"R[-{num_quarters - 1}]C[{x_rel}]:RC[{x_rel}])"
            )
            slope_formula = (
                f"=SLOPE(R[-{num_quarters - 1}]C[{y_rel}]:RC[{y_rel}],"
                f"R[-{num_quarters - 1}]C[{x_rel}]:RC[{x_rel}])"
            )

        safe_set_formula2_r1c1(sheet.cells(row, intercept_col), intercept_formula)
        safe_set_formula2_r1c1(sheet.cells(row, slope_col), slope_formula)

    workbook.app.calculate()

    intercept_values = sheet.range((start_row, intercept_col), (end_row, intercept_col)).value
    slope_values = sheet.range((start_row, slope_col), (end_row, slope_col)).value
    if not isinstance(intercept_values, list):
        intercept_values = [intercept_values]
    if not isinstance(slope_values, list):
        slope_values = [slope_values]
    sheet.range((start_row, intercept_col), (end_row, slope_col)).value = None

    rows: List[Dict[str, Any]] = []
    for idx, row in enumerate(row_numbers):
        num_quarters = as_int(safe_cell_value(sheet, row, col_num_q), fallback=idx + 1) or (idx + 1)
        actual_value = as_float(safe_cell_value(sheet, row, col_actual))
        forecast_value = as_float(safe_cell_value(sheet, row, col_forecast_value))
        forecast_max = as_float(safe_cell_value(sheet, row, col_forecast_max))
        forecast_min = as_float(safe_cell_value(sheet, row, col_forecast_min))
        intercept = as_float(intercept_values[idx] if idx < len(intercept_values) else None)
        slope = as_float(slope_values[idx] if idx < len(slope_values) else None)

        if all(
            value is None
            for value in (num_quarters, forecast_value, forecast_max, forecast_min, intercept, slope)
        ):
            continue

        row_data = {
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
            "range_width": subtract_values(forecast_max, forecast_min),
            "intercept": intercept,
            "slope": slope,
            "source_file": source_file,
        }

        if rows and rows_are_duplicate(
            rows[-1],
            row_data,
            fields=("forecast_value", "forecast_max", "forecast_min", "intercept", "slope"),
        ):
            continue

        rows.append(row_data)

    return rows


def write_sheet(worksheet, columns: Sequence[str], rows: Sequence[Dict[str, Any]]) -> None:
    worksheet.append(list(columns))
    for row in rows:
        worksheet.append([row.get(column) for column in columns])

    for cell in worksheet[1]:
        cell.font = Font(bold=True)

    worksheet.freeze_panes = "A2"
    worksheet.auto_filter.ref = f"A1:{get_column_letter(len(columns))}{worksheet.max_row}"

    for col_idx, column_name in enumerate(columns, start=1):
        values = [column_name]
        for row_idx in range(2, worksheet.max_row + 1):
            value = worksheet.cell(row=row_idx, column=col_idx).value
            if value is not None:
                values.append(str(value))
        width = min(max((len(value) for value in values), default=10) + 2, 48)
        worksheet.column_dimensions[get_column_letter(col_idx)].width = width


def build_output_path(input_path: Path, output_path: Path) -> Path:
    base_name = f"{input_path.name}_PARAM"
    candidate = output_path / f"{base_name}.xlsx"
    if not candidate.exists():
        return candidate

    index = 1
    while True:
        candidate = output_path / f"{base_name}.{index}.xlsx"
        if not candidate.exists():
            return candidate
        index += 1


def write_output_workbook(
    output_file: Path,
    empirical_rows: Sequence[Dict[str, Any]],
    regression_rows: Sequence[Dict[str, Any]],
) -> None:
    workbook = Workbook()
    empirical_sheet = workbook.active
    empirical_sheet.title = "empirical_candidates"
    write_sheet(empirical_sheet, EMPIRICAL_COLUMNS, empirical_rows)

    regression_sheet = workbook.create_sheet("regression_candidates")
    write_sheet(regression_sheet, REGRESSION_COLUMNS, regression_rows)

    workbook.save(output_file)
    workbook.close()


def iter_input_files(input_path: Path) -> Iterable[Path]:
    for path in sorted(input_path.iterdir(), key=lambda p: p.name.lower()):
        yield path


def skip_reason(file_path: Path) -> Optional[str]:
    if not file_path.is_file():
        return "not a file"
    if file_path.name.startswith("~"):
        return "temp file"
    if file_path.suffix.lower() != ".xlsx":
        return "not an .xlsx file"
    return None


def main() -> None:
    input_path = Path(input_dir).expanduser()
    output_path = Path(output_dir).expanduser()

    if not input_path.exists() or not input_path.is_dir():
        raise SystemExit(f"input_dir does not exist or is not a directory: {input_path}")

    output_path.mkdir(parents=True, exist_ok=True)
    output_file = build_output_path(input_path, output_path)

    processed_files = 0
    empirical_rows: List[Dict[str, Any]] = []
    regression_rows: List[Dict[str, Any]] = []

    app: Optional[xw.App] = None
    original_calculation_mode: Optional[str] = None

    try:
        app = xw.App(visible=False, add_book=False)
        app.display_alerts = False
        app.screen_updating = False
        try:
            app.enable_events = False
        except Exception:
            pass
        try:
            original_calculation_mode = app.calculation
            app.calculation = "manual"
        except Exception:
            original_calculation_mode = None

        for file_path in iter_input_files(input_path):
            reason = skip_reason(file_path)
            if reason:
                print(f"skipped {file_path.name}: {reason}")
                continue

            print(f"processing {file_path.name}")
            workbook: Optional[xw.Book] = None
            try:
                workbook = app.books.open(str(file_path), update_links=False)
                labels = parse_file_labels(file_path)
                empirical_rows.extend(extract_empirical_candidates(workbook, labels, file_path.name))
                regression_rows.extend(extract_regression_candidates(workbook, labels, file_path.name))
                processed_files += 1
                print(f"processed {file_path.name}")
            except Exception as exc:
                print(f"skipped {file_path.name}: {exc}")
            finally:
                if workbook is not None:
                    close_workbook_safely(workbook)

    finally:
        if app is not None:
            if original_calculation_mode is not None:
                try:
                    app.calculation = original_calculation_mode
                except Exception:
                    pass
            try:
                app.quit()
            except Exception:
                pass

    write_output_workbook(output_file, empirical_rows, regression_rows)

    print(f"output path: {output_file}")
    print(f"number of files processed: {processed_files}")
    print(f"number of empirical rows: {len(empirical_rows)}")
    print(f"number of regression rows: {len(regression_rows)}")


if __name__ == "__main__":
    main()
