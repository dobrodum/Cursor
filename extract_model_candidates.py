#!/usr/bin/env python3
from __future__ import annotations

from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path
import math
import re
from typing import Any, Dict, List, Optional, Sequence, Tuple

import xlwings as xw


# Update these before running.
input_dir = Path("input")
output_dir = Path("output")


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

DAY_BY_PERIOD = {
    "early": 5,
    "mid": 15,
    "late": 25,
}

FILE_PATTERN = re.compile(
    r"-\s*(?P<ticker>[A-Za-z0-9]+)\s*-\s*(?P<period>Early|Mid|Late)"
    r"(?P<month>[A-Za-z]{3,9})(?P<year>\d{4})",
    re.IGNORECASE,
)


@dataclass
class ModelLabel:
    model: str
    ticker: str
    model_period: str
    model_date: str


def normalize_text(value: Any) -> str:
    if value is None:
        return ""
    cleaned = re.sub(r"[^a-z0-9]+", " ", str(value).strip().lower())
    return cleaned.strip()


def to_2d(values: Any) -> List[List[Any]]:
    if values is None:
        return []
    if not isinstance(values, list):
        return [[values]]
    if not values:
        return []
    if isinstance(values[0], list):
        max_len = max((len(row) for row in values), default=0)
        normalized: List[List[Any]] = []
        for row in values:
            normalized.append(row + [None] * (max_len - len(row)))
        return normalized
    return [values]


def to_float(value: Any) -> Optional[float]:
    if value is None or isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        parsed = float(value)
        if math.isnan(parsed) or math.isinf(parsed):
            return None
        return parsed
    if isinstance(value, str):
        text = value.strip()
        if not text:
            return None
        is_pct = text.endswith("%")
        text = text.replace(",", "").replace("%", "")
        try:
            parsed = float(text)
        except ValueError:
            return None
        return parsed / 100.0 if is_pct else parsed
    return None


def maybe_int(value: Optional[float]) -> Optional[int]:
    if value is None:
        return None
    rounded = int(round(value))
    if abs(value - rounded) < 1e-9:
        return rounded
    return None


def safe_subtract(a: Any, b: Any) -> Optional[float]:
    left = to_float(a)
    right = to_float(b)
    if left is None or right is None:
        return None
    return left - right


def parse_model_label(file_name: str) -> ModelLabel:
    stem = Path(file_name).stem
    match = FILE_PATTERN.search(stem)
    if not match:
        chunks = [chunk.strip() for chunk in stem.split("-") if chunk.strip()]
        fallback_ticker = chunks[1].upper() if len(chunks) > 1 else "UNKNOWN"
        fallback_period = "unknown_period"
        return ModelLabel(
            model=f"{fallback_ticker}_{fallback_period}",
            ticker=fallback_ticker,
            model_period=fallback_period,
            model_date="",
        )

    ticker = match.group("ticker").upper()
    period_token = match.group("period").title()
    month_token = match.group("month")[:3].title()
    year_num = int(match.group("year"))

    try:
        month_num = datetime.strptime(month_token, "%b").month
        day = DAY_BY_PERIOD[period_token.lower()]
        model_date = date(year_num, month_num, day).isoformat()
    except Exception:
        model_date = ""

    model_period = f"{period_token}{month_token}_{year_num}"
    model = f"{ticker}_{model_period}"
    return ModelLabel(
        model=model,
        ticker=ticker,
        model_period=model_period,
        model_date=model_date,
    )


def output_path_for_run(source_input_dir: Path, source_output_dir: Path) -> Path:
    source_output_dir.mkdir(parents=True, exist_ok=True)
    base_name = f"{source_input_dir.name}_PARAM"
    output_path = source_output_dir / f"{base_name}.xlsx"
    suffix = 1
    while output_path.exists():
        output_path = source_output_dir / f"{base_name}.{suffix}.xlsx"
        suffix += 1
    return output_path


def safe_close_workbook(workbook: xw.Book) -> None:
    close_attempts = (
        lambda: workbook.close(save=False),
        lambda: workbook.close(False),
        lambda: workbook.api.Close(False),
        lambda: workbook.api.Close(SaveChanges=False),
        lambda: workbook.close(),
    )
    for close_fn in close_attempts:
        try:
            close_fn()
            return
        except Exception:
            continue


def set_formula2(cell: xw.Range, formula_r1c1: str) -> None:
    try:
        cell.formula2 = formula_r1c1
    except Exception:
        cell.formula = formula_r1c1


def load_sheet_grid(sheet: xw.Sheet) -> Tuple[List[List[Any]], int, int, int, int]:
    used = sheet.used_range
    grid = to_2d(used.value)
    if not grid:
        return [], used.row, used.column, used.row, used.column
    start_row = used.row
    start_col = used.column
    end_row = start_row + len(grid) - 1
    end_col = start_col + len(grid[0]) - 1
    return grid, start_row, start_col, end_row, end_col


def grid_cell(
    grid: Sequence[Sequence[Any]],
    start_row: int,
    start_col: int,
    row: int,
    col: Optional[int],
) -> Any:
    if col is None or col < 1:
        return None
    row_idx = row - start_row
    col_idx = col - start_col
    if row_idx < 0 or row_idx >= len(grid):
        return None
    if col_idx < 0 or col_idx >= len(grid[row_idx]):
        return None
    return grid[row_idx][col_idx]


def find_anchor_max(
    grid: Sequence[Sequence[Any]],
    start_row: int,
    start_col: int,
) -> Optional[Tuple[int, int]]:
    for row_offset, row_values in enumerate(grid):
        for col_offset, value in enumerate(row_values):
            if normalize_text(value) == "max":
                return start_row + row_offset, start_col + col_offset
    return None


def build_header_map(
    grid: Sequence[Sequence[Any]],
    start_row: int,
    start_col: int,
    header_row: int,
) -> Dict[str, int]:
    row_idx = header_row - start_row
    if row_idx < 0 or row_idx >= len(grid):
        return {}
    mapping: Dict[str, int] = {}
    for col_idx, value in enumerate(grid[row_idx]):
        key = normalize_text(value)
        if key:
            mapping[key] = start_col + col_idx
    return mapping


def resolve_column(
    header_map: Dict[str, int],
    phrases: Sequence[str],
    default: Optional[int],
) -> Optional[int]:
    targets = [normalize_text(phrase) for phrase in phrases]
    for target in targets:
        if target in header_map:
            return header_map[target]
    for key, col in header_map.items():
        for target in targets:
            if target and (target in key or key in target):
                return col
    return default


def collect_table_rows(
    grid: Sequence[Sequence[Any]],
    start_row: int,
    start_col: int,
    first_data_row: int,
    last_data_row: int,
    columns: Dict[str, Optional[int]],
) -> List[Dict[str, Any]]:
    records: List[Dict[str, Any]] = []
    for row in range(first_data_row, last_data_row + 1):
        record: Dict[str, Any] = {"_row": row}
        for key, col in columns.items():
            record[key] = grid_cell(grid, start_row, start_col, row, col)
        record["num_quarters_used"] = maybe_int(to_float(record.get("num_quarters_used")))

        signal_values = [
            to_float(record.get("forecast_value")),
            to_float(record.get("forecast_max")),
            to_float(record.get("forecast_min")),
            to_float(record.get("avg_penetration_pct")),
            to_float(record.get("reported_sales")),
        ]
        has_signal = any(item is not None for item in signal_values)
        if has_signal:
            records.append(record)
        elif records and len(records) >= MAX_QUARTERS:
            break
    return records


def empirical_avg_pen_by_quarter(
    workbook: xw.Book,
    sheet: xw.Sheet,
    grid: Sequence[Sequence[Any]],
    start_row: int,
    start_col: int,
    end_row: int,
    end_col: int,
    anchor_row: int,
    penetration_col: Optional[int],
) -> Dict[int, float]:
    if penetration_col is None:
        return {}

    history_rows: List[int] = []
    for row in range(start_row, anchor_row):
        pen_value = to_float(grid_cell(grid, start_row, start_col, row, penetration_col))
        if pen_value is not None:
            history_rows.append(row)

    if not history_rows:
        return {}

    max_quarters = min(MAX_QUARTERS, len(history_rows))
    scratch_start_row = end_row + 2
    scratch_col = max(end_col + 4, 250)

    for idx, n_quarters in enumerate(range(1, max_quarters + 1)):
        first_row = history_rows[-n_quarters]
        last_row = history_rows[-1]
        avg_formula = (
            f"=AVERAGE(R{first_row}C{penetration_col}:R{last_row}C{penetration_col})"
        )
        set_formula2(sheet.range((scratch_start_row + idx, scratch_col)), avg_formula)

    workbook.app.calculate()
    values = to_2d(
        sheet.range(
            (scratch_start_row, scratch_col),
            (scratch_start_row + max_quarters - 1, scratch_col),
        ).value
    )
    output: Dict[int, float] = {}
    for idx, n_quarters in enumerate(range(1, max_quarters + 1)):
        parsed = to_float(values[idx][0]) if idx < len(values) else None
        if parsed is not None:
            output[n_quarters] = parsed
    return output


def extract_empirical_rows(
    workbook: xw.Book,
    label: ModelLabel,
    source_file: str,
) -> List[Dict[str, Any]]:
    try:
        sheet = workbook.sheets["Empirical Model"]
    except Exception:
        return []

    grid, start_row, start_col, end_row, end_col = load_sheet_grid(sheet)
    if not grid:
        return []

    anchor = find_anchor_max(grid, start_row, start_col)
    if anchor is None:
        return []
    anchor_row, anchor_col = anchor

    header_map = build_header_map(grid, start_row, start_col, anchor_row)
    columns = {
        "num_quarters_used": resolve_column(
            header_map,
            ["num_quarters_used", "num quarters used", "quarters used", "num quarters"],
            default=anchor_col - 8,
        ),
        "last_quarter_used": resolve_column(
            header_map,
            ["last_quarter_used", "last quarter used", "last quarter"],
            default=anchor_col - 7,
        ),
        "forecast_value": resolve_column(
            header_map,
            ["estimated total sold", "est total sold", "forecast value", "tot fcst w/o sa"],
            default=anchor_col - 1,
        ),
        "forecast_max": resolve_column(header_map, ["max"], default=anchor_col),
        "forecast_min": resolve_column(header_map, ["min"], default=anchor_col + 1),
        "avg_penetration_pct": resolve_column(
            header_map,
            ["avg_penetration_pct", "avg penetration", "average penetration", "penetration pct"],
            default=anchor_col - 6,
        ),
        "quarterly_sales": resolve_column(
            header_map,
            ["quarterly_sales", "quarterly sales", "qtr sales"],
            default=anchor_col - 5,
        ),
        "reported_sales": resolve_column(
            header_map,
            ["reported_sales", "reported sales", "actual sales"],
            default=anchor_col - 4,
        ),
        "growth_rate_pct": resolve_column(
            header_map,
            ["growth_rate_pct", "growth rate pct", "growth rate"],
            default=anchor_col - 3,
        ),
        "sales_captured_in_db_pct": resolve_column(
            header_map,
            ["sales_captured_in_db_pct", "sales captured in db", "captured in db"],
            default=anchor_col - 2,
        ),
    }

    penetration_col = resolve_column(
        header_map,
        ["penetration", "penetration pct", "quarterly penetration"],
        default=columns["avg_penetration_pct"],
    )
    avg_pen_by_quarter = empirical_avg_pen_by_quarter(
        workbook=workbook,
        sheet=sheet,
        grid=grid,
        start_row=start_row,
        start_col=start_col,
        end_row=end_row,
        end_col=end_col,
        anchor_row=anchor_row,
        penetration_col=penetration_col,
    )

    table_rows = collect_table_rows(
        grid=grid,
        start_row=start_row,
        start_col=start_col,
        first_data_row=anchor_row + 1,
        last_data_row=min(end_row, anchor_row + 60),
        columns=columns,
    )
    if not table_rows:
        return []

    rows_by_quarter: Dict[int, Dict[str, Any]] = {}
    fallback_rows: List[Dict[str, Any]] = []
    for row_data in table_rows:
        fallback_rows.append(row_data)
        n_quarters = row_data.get("num_quarters_used")
        if isinstance(n_quarters, int) and 1 <= n_quarters <= MAX_QUARTERS:
            rows_by_quarter.setdefault(n_quarters, row_data)

    output: List[Dict[str, Any]] = []
    for n_quarters in range(1, MAX_QUARTERS + 1):
        source_row = rows_by_quarter.get(n_quarters)
        if source_row is None and n_quarters - 1 < len(fallback_rows):
            source_row = fallback_rows[n_quarters - 1]
        if source_row is None:
            continue

        avg_penetration = avg_pen_by_quarter.get(
            n_quarters, to_float(source_row.get("avg_penetration_pct"))
        )
        forecast_max = to_float(source_row.get("forecast_max"))
        forecast_min = to_float(source_row.get("forecast_min"))
        reported_sales = to_float(source_row.get("reported_sales"))
        forecast_value = to_float(source_row.get("forecast_value"))
        if forecast_value is None and reported_sales is not None and avg_penetration not in (None, 0.0):
            forecast_value = reported_sales / avg_penetration

        output.append(
            {
                "model": label.model,
                "ticker": label.ticker,
                "model_period": label.model_period,
                "model_date": label.model_date,
                "method": "empirical",
                "parameter_name": "avg_penetration_pct",
                "parameter_value": avg_penetration,
                "num_quarters_used": n_quarters,
                "last_quarter_used": source_row.get("last_quarter_used"),
                "forecast_value": forecast_value,
                "actual_value": reported_sales,
                "forecast_max": forecast_max,
                "forecast_min": forecast_min,
                "range_width": safe_subtract(forecast_max, forecast_min),
                "avg_penetration_pct": avg_penetration,
                "quarterly_sales": to_float(source_row.get("quarterly_sales")),
                "reported_sales": reported_sales,
                "growth_rate_pct": to_float(source_row.get("growth_rate_pct")),
                "sales_captured_in_db_pct": to_float(source_row.get("sales_captured_in_db_pct")),
                "source_file": source_file,
            }
        )
    return output


def regression_intercept_slope_by_quarter(
    workbook: xw.Book,
    sheet: xw.Sheet,
    point_rows: Sequence[int],
    x_col: int,
    y_col: int,
    end_row: int,
    end_col: int,
) -> Dict[int, Tuple[Optional[float], Optional[float]]]:
    if len(point_rows) < 2:
        return {}

    max_quarters = min(MAX_QUARTERS, len(point_rows))
    scratch_start_row = end_row + 2
    scratch_col = max(end_col + 4, 260)
    quarter_counts = list(range(2, max_quarters + 1))

    for idx, n_quarters in enumerate(quarter_counts):
        first_row = point_rows[-n_quarters]
        last_row = point_rows[-1]
        intercept_formula = (
            f"=INTERCEPT(R{first_row}C{y_col}:R{last_row}C{y_col},"
            f"R{first_row}C{x_col}:R{last_row}C{x_col})"
        )
        slope_formula = (
            f"=SLOPE(R{first_row}C{y_col}:R{last_row}C{y_col},"
            f"R{first_row}C{x_col}:R{last_row}C{x_col})"
        )
        set_formula2(sheet.range((scratch_start_row + idx, scratch_col)), intercept_formula)
        set_formula2(sheet.range((scratch_start_row + idx, scratch_col + 1)), slope_formula)

    workbook.app.calculate()
    values = to_2d(
        sheet.range(
            (scratch_start_row, scratch_col),
            (scratch_start_row + len(quarter_counts) - 1, scratch_col + 1),
        ).value
    )
    output: Dict[int, Tuple[Optional[float], Optional[float]]] = {}
    for idx, n_quarters in enumerate(quarter_counts):
        intercept = to_float(values[idx][0]) if idx < len(values) else None
        slope = to_float(values[idx][1]) if idx < len(values) else None
        output[n_quarters] = (intercept, slope)
    return output


def extract_regression_rows(
    workbook: xw.Book,
    label: ModelLabel,
    source_file: str,
) -> List[Dict[str, Any]]:
    try:
        sheet = workbook.sheets["Regression Model"]
    except Exception:
        return []

    grid, start_row, start_col, end_row, end_col = load_sheet_grid(sheet)
    if not grid:
        return []

    anchor = find_anchor_max(grid, start_row, start_col)
    if anchor is None:
        return []
    anchor_row, anchor_col = anchor

    header_map = build_header_map(grid, start_row, start_col, anchor_row)
    x_col = anchor_col - 11
    y_col = anchor_col - 7

    columns = {
        "num_quarters_used": resolve_column(
            header_map,
            ["num_quarters_used", "num quarters used", "quarters used", "num quarters"],
            default=anchor_col - 8,
        ),
        "forecast_value": resolve_column(
            header_map,
            ["tot fcst w/o sa", "tot fcst wo sa", "forecast value", "forecast_total_without_sa"],
            default=anchor_col - 1,
        ),
        "actual_value": resolve_column(
            header_map,
            ["actual value", "reported sales", "actual sales"],
            default=None,
        ),
        "forecast_max": resolve_column(header_map, ["max"], default=anchor_col),
        "forecast_min": resolve_column(header_map, ["min"], default=anchor_col + 1),
    }

    point_rows: List[int] = []
    for row in range(start_row, end_row + 1):
        x_val = to_float(grid_cell(grid, start_row, start_col, row, x_col))
        y_val = to_float(grid_cell(grid, start_row, start_col, row, y_col))
        if x_val is not None and y_val is not None:
            point_rows.append(row)

    intercept_slope = regression_intercept_slope_by_quarter(
        workbook=workbook,
        sheet=sheet,
        point_rows=point_rows,
        x_col=x_col,
        y_col=y_col,
        end_row=end_row,
        end_col=end_col,
    )

    table_rows = collect_table_rows(
        grid=grid,
        start_row=start_row,
        start_col=start_col,
        first_data_row=anchor_row + 1,
        last_data_row=min(end_row, anchor_row + 60),
        columns=columns,
    )
    if not table_rows:
        return []

    rows_by_quarter: Dict[int, Dict[str, Any]] = {}
    fallback_rows: List[Dict[str, Any]] = []
    for row_data in table_rows:
        fallback_rows.append(row_data)
        n_quarters = row_data.get("num_quarters_used")
        if isinstance(n_quarters, int) and 1 <= n_quarters <= MAX_QUARTERS:
            rows_by_quarter.setdefault(n_quarters, row_data)

    output: List[Dict[str, Any]] = []
    previous_signature: Optional[Tuple[Any, ...]] = None
    max_quarters = min(MAX_QUARTERS, max(intercept_slope.keys(), default=0))
    for n_quarters in range(2, max_quarters + 1):
        source_row = rows_by_quarter.get(n_quarters)
        if source_row is None and n_quarters - 1 < len(fallback_rows):
            source_row = fallback_rows[n_quarters - 1]
        if source_row is None:
            continue

        intercept, slope = intercept_slope.get(n_quarters, (None, None))
        forecast_value = to_float(source_row.get("forecast_value"))
        forecast_max = to_float(source_row.get("forecast_max"))
        forecast_min = to_float(source_row.get("forecast_min"))
        actual_value_float = to_float(source_row.get("actual_value"))
        actual_value: Any = actual_value_float if actual_value_float is not None else ""

        signature = (
            n_quarters,
            round(forecast_value, 10) if forecast_value is not None else None,
            round(forecast_max, 10) if forecast_max is not None else None,
            round(forecast_min, 10) if forecast_min is not None else None,
            round(intercept, 10) if intercept is not None else None,
            round(slope, 10) if slope is not None else None,
        )
        if signature == previous_signature:
            continue
        previous_signature = signature

        output.append(
            {
                "model": label.model,
                "ticker": label.ticker,
                "model_period": label.model_period,
                "model_date": label.model_date,
                "method": "regression",
                "parameter_name": "num_quarters_used",
                "parameter_value": n_quarters,
                "num_quarters_used": n_quarters,
                "forecast_value": forecast_value,
                "actual_value": actual_value,
                "forecast_max": forecast_max,
                "forecast_min": forecast_min,
                "range_width": safe_subtract(forecast_max, forecast_min),
                "intercept": intercept,
                "slope": slope,
                "source_file": source_file,
            }
        )
    return output


def write_rows_to_sheet(
    workbook: xw.Book,
    sheet: xw.Sheet,
    columns: Sequence[str],
    rows: Sequence[Dict[str, Any]],
) -> None:
    matrix: List[List[Any]] = [list(columns)]
    for row in rows:
        matrix.append([row.get(col, "") for col in columns])

    sheet.range("A1").value = matrix
    last_row = max(1, len(matrix))
    last_col = len(columns)

    header_range = sheet.range((1, 1), (1, last_col))
    table_range = sheet.range((1, 1), (last_row, last_col))

    header_range.api.Font.Bold = True
    table_range.api.AutoFilter()

    try:
        sheet.autofit("c")
    except Exception:
        pass

    for col_idx, column_name in enumerate(columns, start=1):
        cell = sheet.range((1, col_idx))
        width = cell.column_width
        if width is None:
            width = 10
        width = max(width, min(max(12, len(column_name) + 2), 48))
        cell.column_width = min(width, 56)

    sheet.activate()
    active_window = workbook.app.api.ActiveWindow
    active_window.SplitColumn = 0
    active_window.SplitRow = 1
    active_window.FreezePanes = True


def write_output_workbook(
    app: xw.App,
    output_path: Path,
    empirical_rows: Sequence[Dict[str, Any]],
    regression_rows: Sequence[Dict[str, Any]],
) -> None:
    output_book = app.books.add()
    try:
        empirical_sheet = output_book.sheets[0]
        empirical_sheet.name = "empirical_candidates"

        if len(output_book.sheets) >= 2:
            regression_sheet = output_book.sheets[1]
            regression_sheet.name = "regression_candidates"
        else:
            regression_sheet = output_book.sheets.add(
                "regression_candidates", after=empirical_sheet
            )

        while len(output_book.sheets) > 2:
            output_book.sheets[-1].delete()

        write_rows_to_sheet(output_book, empirical_sheet, EMPIRICAL_COLUMNS, empirical_rows)
        write_rows_to_sheet(output_book, regression_sheet, REGRESSION_COLUMNS, regression_rows)
        output_book.save(str(output_path))
    finally:
        try:
            output_book.close()
        except Exception:
            pass


def iter_source_files(scan_input_dir: Path) -> List[Path]:
    if not scan_input_dir.exists():
        return []
    return sorted(path for path in scan_input_dir.iterdir() if path.is_file())


def should_skip_input_file(file_path: Path, input_folder_name: str) -> Optional[str]:
    if file_path.name.startswith("~"):
        return "temporary file"
    if file_path.suffix.lower() != ".xlsx":
        return "not an .xlsx file"
    if file_path.name.startswith(f"{input_folder_name}_PARAM"):
        return "previous output file"
    return None


def main() -> None:
    resolved_input_dir = input_dir.expanduser().resolve()
    resolved_output_dir = output_dir.expanduser().resolve()

    if not resolved_input_dir.exists():
        raise SystemExit(f"input_dir does not exist: {resolved_input_dir}")

    output_path = output_path_for_run(resolved_input_dir, resolved_output_dir)
    empirical_rows: List[Dict[str, Any]] = []
    regression_rows: List[Dict[str, Any]] = []
    files_processed = 0

    app = xw.App(visible=False, add_book=False)
    app.display_alerts = False
    app.screen_updating = False

    try:
        try:
            app.calculation = "manual"
        except Exception:
            pass

        for file_path in iter_source_files(resolved_input_dir):
            skip_reason = should_skip_input_file(file_path, resolved_input_dir.name)
            if skip_reason is not None:
                print(f"Skipped {file_path.name}: {skip_reason}")
                continue

            workbook: Optional[xw.Book] = None
            try:
                workbook = app.books.open(str(file_path), update_links=False)
                label = parse_model_label(file_path.name)

                empirical_rows.extend(
                    extract_empirical_rows(
                        workbook=workbook,
                        label=label,
                        source_file=file_path.name,
                    )
                )
                regression_rows.extend(
                    extract_regression_rows(
                        workbook=workbook,
                        label=label,
                        source_file=file_path.name,
                    )
                )
                files_processed += 1
                print(f"Processed {file_path.name}")
            except Exception as exc:
                print(f"Skipped {file_path.name}: processing error: {exc}")
            finally:
                if workbook is not None:
                    safe_close_workbook(workbook)

        write_output_workbook(
            app=app,
            output_path=output_path,
            empirical_rows=empirical_rows,
            regression_rows=regression_rows,
        )
    finally:
        try:
            app.quit()
        except Exception:
            pass

    print(f"Output path: {output_path}")
    print(f"Number of files processed: {files_processed}")
    print(f"Number of empirical rows: {len(empirical_rows)}")
    print(f"Number of regression rows: {len(regression_rows)}")


if __name__ == "__main__":
    main()
