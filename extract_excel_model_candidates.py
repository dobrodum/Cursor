#!/usr/bin/env python3
"""Extract empirical/regression candidate rows from model workbooks.

This script scans an input directory for `.xlsx` model files and writes a
single output workbook containing:
  - empirical_candidates
  - regression_candidates

Performance and safety rules implemented:
  - Uses one hidden Excel app session for the full run.
  - Opens each source workbook exactly once.
  - Processes both target sheets while workbook is open.
  - Never saves source files.
  - Uses anchor-based offsets from the "max" cell.
  - Uses R1C1 `.formula2` for computed fields.
"""

from __future__ import annotations

import re
from datetime import date
from itertools import count
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Sequence, Tuple

try:
    import xlwings as xw
except ModuleNotFoundError:  # pragma: no cover - depends on host environment
    xw = None
try:
    from openpyxl import Workbook
    from openpyxl.styles import Font
    from openpyxl.utils import get_column_letter
except ModuleNotFoundError:  # pragma: no cover - depends on host environment
    Workbook = None
    Font = None
    get_column_letter = None

# --------------------------
# User-configurable paths
# --------------------------
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


MONTH_NUMBER = {
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

PERIOD_DAY = {"early": 5, "mid": 15, "late": 25}


def normalize_text(value: Any) -> str:
    if value is None:
        return ""
    text = str(value).strip().lower()
    text = re.sub(r"[^a-z0-9]+", " ", text)
    return re.sub(r"\s+", " ", text).strip()


def is_blank(value: Any) -> bool:
    return value is None or (isinstance(value, str) and value.strip() == "")


def to_float(value: Any) -> Optional[float]:
    if isinstance(value, bool) or is_blank(value):
        return None
    try:
        return float(value)
    except (TypeError, ValueError):
        return None


def to_int(value: Any, fallback: Optional[int] = None) -> Optional[int]:
    as_float = to_float(value)
    if as_float is None:
        return fallback
    return int(round(as_float))


def format_output_value(value: Any) -> Any:
    if is_blank(value):
        return ""
    as_float = to_float(value)
    if as_float is not None:
        return as_float
    return str(value).strip()


def subtract_or_blank(left: Any, right: Any) -> Any:
    left_float = to_float(left)
    right_float = to_float(right)
    if left_float is None or right_float is None:
        return ""
    return left_float - right_float


def as_matrix(values: Any) -> List[List[Any]]:
    if values is None:
        return []
    if isinstance(values, Sequence) and not isinstance(values, (str, bytes)):
        if not values:
            return []
        first = values[0]
        if isinstance(first, Sequence) and not isinstance(first, (str, bytes)):
            return [list(row) for row in values]
        return [list(values)]
    return [[values]]


def get_unique_output_path(src_dir: Path, dst_dir: Path) -> Path:
    base_name = f"{src_dir.name}_PARAM"
    default_output = dst_dir / f"{base_name}.xlsx"
    if not default_output.exists():
        return default_output

    for i in count(1):
        candidate = dst_dir / f"{base_name}.{i}.xlsx"
        if not candidate.exists():
            return candidate

    raise RuntimeError("Unable to create a unique output filename.")


def parse_file_labels(file_name: str) -> Dict[str, str]:
    stem = Path(file_name).stem
    parts = [piece.strip() for piece in stem.split(" - ")]

    ticker = "UNKNOWN"
    period_token = ""
    if len(parts) >= 2 and parts[1]:
        ticker = parts[1].upper()
    if len(parts) >= 3 and parts[2]:
        period_token = parts[2].split("_")[0]

    period_match = re.search(
        r"(?P<bucket>early|mid|late)(?P<month>[a-z]{3})(?P<year>\d{4})",
        period_token,
        flags=re.IGNORECASE,
    )

    model_period = ""
    model_date = ""

    if period_match:
        bucket = period_match.group("bucket").lower()
        bucket_title = bucket.capitalize()
        month_abbrev = period_match.group("month").lower()
        month_title = month_abbrev.capitalize()
        year = int(period_match.group("year"))
        month_num = MONTH_NUMBER.get(month_abbrev)
        day = PERIOD_DAY.get(bucket)

        if month_num and day:
            model_period = f"{bucket_title}{month_title}_{year}"
            model_date = date(year, month_num, day).isoformat()

    if not model_period and period_token:
        model_period = period_token
    model = f"{ticker}_{model_period}" if model_period else ticker

    return {
        "ticker": ticker,
        "model_period": model_period,
        "model_date": model_date,
        "model": model,
    }


def get_sheet(workbook: xw.Book, sheet_name: str) -> Optional[xw.Sheet]:
    try:
        return workbook.sheets[sheet_name]
    except Exception:
        return None


def find_max_anchor(sheet: xw.Sheet) -> Tuple[int, int]:
    used = sheet.used_range
    values = as_matrix(used.value)
    if not values:
        raise ValueError('Sheet has no used cells to search for "max".')

    start_row = used.row
    start_col = used.column
    for row_offset, row_values in enumerate(values):
        for col_offset, value in enumerate(row_values):
            if normalize_text(value) == "max":
                return start_row + row_offset, start_col + col_offset

    raise ValueError('Could not find "max" anchor cell.')


def header_map_around_anchor(
    sheet: xw.Sheet,
    anchor_row: int,
    anchor_col: int,
    window: int = 25,
) -> Dict[str, int]:
    left_col = max(1, anchor_col - window)
    right_col = anchor_col + window
    row_values = sheet.range((anchor_row, left_col), (anchor_row, right_col)).value
    row_values_list = list(row_values) if isinstance(row_values, list) else [row_values]

    mapping: Dict[str, int] = {}
    for idx, value in enumerate(row_values_list):
        key = normalize_text(value)
        if key and key not in mapping:
            mapping[key] = left_col + idx
    return mapping


def choose_column(
    header_map: Dict[str, int],
    candidate_labels: Iterable[str],
    anchor_col: int,
    fallback_offset: int,
) -> int:
    candidates = [normalize_text(label) for label in candidate_labels]

    for candidate in candidates:
        if candidate in header_map:
            return header_map[candidate]

    for candidate in candidates:
        for header, col_idx in header_map.items():
            if candidate and candidate in header:
                return col_idx

    return max(1, anchor_col + fallback_offset)


def safe_close_workbook(workbook: xw.Book) -> None:
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
        try:
            workbook.close()
        except Exception:
            pass


def _get_block_value(block: List[List[Any]], row_idx: int, col_idx: int, min_col: int) -> Any:
    cell_idx = col_idx - min_col
    if row_idx < 0 or row_idx >= len(block):
        return None
    row = block[row_idx]
    if cell_idx < 0 or cell_idx >= len(row):
        return None
    return row[cell_idx]


def process_empirical_sheet(
    workbook: xw.Book,
    model_info: Dict[str, str],
    source_file: str,
) -> List[Dict[str, Any]]:
    sheet = get_sheet(workbook, "Empirical Model")
    if sheet is None:
        print(f"Skipped empirical extraction for {source_file}: missing 'Empirical Model' sheet.")
        return []

    anchor_row, anchor_col = find_max_anchor(sheet)
    headers = header_map_around_anchor(sheet, anchor_row, anchor_col)

    num_quarters_col = choose_column(
        headers,
        ["num_quarters_used", "num quarters used", "quarters used", "num quarters", "quarters"],
        anchor_col,
        -12,
    )
    last_quarter_col = choose_column(
        headers,
        ["last_quarter_used", "last quarter used", "last quarter"],
        anchor_col,
        -10,
    )
    forecast_value_col = choose_column(
        headers,
        ["estimated total sold", "forecast value", "tot fcst", "total forecast", "forecast"],
        anchor_col,
        -2,
    )
    actual_value_col = choose_column(
        headers,
        ["actual value", "actual", "reported sales"],
        anchor_col,
        -1,
    )
    forecast_min_col = choose_column(headers, ["forecast_min", "forecast min", "min"], anchor_col, 1)
    quarterly_sales_col = choose_column(
        headers,
        ["quarterly_sales", "quarterly sales", "quarter sales"],
        anchor_col,
        -8,
    )
    reported_sales_col = choose_column(
        headers,
        ["reported_sales", "reported sales"],
        anchor_col,
        -7,
    )
    growth_rate_col = choose_column(
        headers,
        ["growth_rate_pct", "growth rate pct", "growth rate", "growth"],
        anchor_col,
        -5,
    )
    sales_captured_col = choose_column(
        headers,
        ["sales_captured_in_db_pct", "sales captured in db pct", "sales captured in db", "captured in db"],
        anchor_col,
        -4,
    )
    avg_penetration_col = choose_column(
        headers,
        ["avg_penetration_pct", "avg penetration pct", "avg penetration", "average penetration"],
        anchor_col,
        -6,
    )

    start_row = anchor_row + 1
    end_row = start_row + N_QUARTERS - 1

    # Formula2 in R1C1 for avg penetration to avoid column-letter conversion.
    temp_avg_col = anchor_col + 30
    avg_formula_range = sheet.range((start_row, temp_avg_col), (end_row, temp_avg_col))
    avg_formulas = [
        [
            f'=IFERROR(AVERAGE(R{start_row}C{sales_captured_col}:R{start_row + idx}C{sales_captured_col}),"")'
        ]
        for idx in range(N_QUARTERS)
    ]
    avg_formula_range.formula2 = avg_formulas
    workbook.app.calculate()
    avg_values = as_matrix(avg_formula_range.value)
    avg_formula_range.clear_contents()

    data_cols = [
        num_quarters_col,
        last_quarter_col,
        forecast_value_col,
        actual_value_col,
        anchor_col,
        forecast_min_col,
        quarterly_sales_col,
        reported_sales_col,
        growth_rate_col,
        sales_captured_col,
        avg_penetration_col,
    ]
    min_col = min(data_cols)
    max_col = max(data_cols)

    data_block = as_matrix(sheet.range((start_row, min_col), (end_row, max_col)).value)

    rows: List[Dict[str, Any]] = []
    for idx in range(N_QUARTERS):
        row_num = start_row + idx
        num_quarters_used = to_int(_get_block_value(data_block, idx, num_quarters_col, min_col), fallback=idx + 1)
        last_quarter_used = _get_block_value(data_block, idx, last_quarter_col, min_col)
        forecast_value = _get_block_value(data_block, idx, forecast_value_col, min_col)
        actual_value = _get_block_value(data_block, idx, actual_value_col, min_col)
        forecast_max = _get_block_value(data_block, idx, anchor_col, min_col)
        forecast_min = _get_block_value(data_block, idx, forecast_min_col, min_col)
        quarterly_sales = _get_block_value(data_block, idx, quarterly_sales_col, min_col)
        reported_sales = _get_block_value(data_block, idx, reported_sales_col, min_col)
        growth_rate_pct = _get_block_value(data_block, idx, growth_rate_col, min_col)
        sales_captured_in_db_pct = _get_block_value(data_block, idx, sales_captured_col, min_col)
        avg_penetration_from_sheet = _get_block_value(data_block, idx, avg_penetration_col, min_col)

        avg_penetration_calc = avg_values[idx][0] if idx < len(avg_values) and avg_values[idx] else None
        avg_penetration_pct = avg_penetration_calc
        if is_blank(avg_penetration_pct):
            avg_penetration_pct = avg_penetration_from_sheet

        if all(
            is_blank(value)
            for value in (
                forecast_value,
                actual_value,
                forecast_max,
                forecast_min,
                avg_penetration_pct,
                quarterly_sales,
                reported_sales,
            )
        ):
            continue

        rows.append(
            {
                "model": model_info["model"],
                "ticker": model_info["ticker"],
                "model_period": model_info["model_period"],
                "model_date": model_info["model_date"],
                "method": "empirical",
                "parameter_name": "avg_penetration_pct",
                "parameter_value": format_output_value(avg_penetration_pct),
                "num_quarters_used": num_quarters_used if num_quarters_used is not None else "",
                "last_quarter_used": format_output_value(last_quarter_used),
                "forecast_value": format_output_value(forecast_value),
                "actual_value": format_output_value(actual_value),
                "forecast_max": format_output_value(forecast_max),
                "forecast_min": format_output_value(forecast_min),
                "range_width": format_output_value(subtract_or_blank(forecast_max, forecast_min)),
                "avg_penetration_pct": format_output_value(avg_penetration_pct),
                "quarterly_sales": format_output_value(quarterly_sales),
                "reported_sales": format_output_value(reported_sales),
                "growth_rate_pct": format_output_value(growth_rate_pct),
                "sales_captured_in_db_pct": format_output_value(sales_captured_in_db_pct),
                "source_file": source_file,
            }
        )

    return rows


def _signature_value(value: Any) -> Any:
    as_float = to_float(value)
    if as_float is not None:
        return round(as_float, 8)
    if is_blank(value):
        return ""
    return str(value).strip()


def process_regression_sheet(
    workbook: xw.Book,
    model_info: Dict[str, str],
    source_file: str,
) -> List[Dict[str, Any]]:
    sheet = get_sheet(workbook, "Regression Model")
    if sheet is None:
        print(f"Skipped regression extraction for {source_file}: missing 'Regression Model' sheet.")
        return []

    anchor_row, anchor_col = find_max_anchor(sheet)
    headers = header_map_around_anchor(sheet, anchor_row, anchor_col)

    x_col = anchor_col - 11
    y_col = anchor_col - 7

    num_quarters_col = choose_column(
        headers,
        ["num_quarters_used", "num quarters used", "quarters used", "num quarters", "quarters"],
        anchor_col,
        -12,
    )
    forecast_value_col = choose_column(
        headers,
        [
            "tot fcst w o sa",
            "tot fcst w/o sa",
            "tot fcst without sa",
            "forecast total without sa",
            "forecast value",
        ],
        anchor_col,
        -2,
    )
    actual_value_col = choose_column(headers, ["actual value", "actual"], anchor_col, -1)
    forecast_min_col = choose_column(headers, ["forecast_min", "forecast min", "min"], anchor_col, 1)

    start_row = anchor_row + 1
    end_row = start_row + N_QUARTERS - 1

    # Formula2 in R1C1 for INTERCEPT and SLOPE.
    temp_intercept_col = anchor_col + 30
    temp_slope_col = anchor_col + 31

    intercept_range = sheet.range((start_row, temp_intercept_col), (end_row, temp_intercept_col))
    slope_range = sheet.range((start_row, temp_slope_col), (end_row, temp_slope_col))

    intercept_formulas = [
        [f'=IFERROR(INTERCEPT(R{start_row}C{y_col}:R{start_row + idx}C{y_col},R{start_row}C{x_col}:R{start_row + idx}C{x_col}),"")']
        for idx in range(N_QUARTERS)
    ]
    slope_formulas = [
        [f'=IFERROR(SLOPE(R{start_row}C{y_col}:R{start_row + idx}C{y_col},R{start_row}C{x_col}:R{start_row + idx}C{x_col}),"")']
        for idx in range(N_QUARTERS)
    ]

    intercept_range.formula2 = intercept_formulas
    slope_range.formula2 = slope_formulas
    workbook.app.calculate()
    intercept_values = as_matrix(intercept_range.value)
    slope_values = as_matrix(slope_range.value)
    intercept_range.clear_contents()
    slope_range.clear_contents()

    data_cols = [num_quarters_col, forecast_value_col, actual_value_col, anchor_col, forecast_min_col]
    min_col = min(data_cols)
    max_col = max(data_cols)
    data_block = as_matrix(sheet.range((start_row, min_col), (end_row, max_col)).value)

    rows: List[Dict[str, Any]] = []
    previous_signature: Optional[Tuple[Any, ...]] = None

    for idx in range(N_QUARTERS):
        num_quarters_used = to_int(_get_block_value(data_block, idx, num_quarters_col, min_col), fallback=idx + 1)
        forecast_value = _get_block_value(data_block, idx, forecast_value_col, min_col)
        actual_value = _get_block_value(data_block, idx, actual_value_col, min_col)
        forecast_max = _get_block_value(data_block, idx, anchor_col, min_col)
        forecast_min = _get_block_value(data_block, idx, forecast_min_col, min_col)
        intercept_value = intercept_values[idx][0] if idx < len(intercept_values) and intercept_values[idx] else None
        slope_value = slope_values[idx][0] if idx < len(slope_values) and slope_values[idx] else None

        if all(
            is_blank(value)
            for value in (forecast_value, forecast_max, forecast_min, intercept_value, slope_value)
        ):
            continue

        signature = (
            _signature_value(num_quarters_used),
            _signature_value(forecast_value),
            _signature_value(forecast_max),
            _signature_value(forecast_min),
            _signature_value(intercept_value),
            _signature_value(slope_value),
        )

        # Some source files emit a duplicate final row; skip only that case.
        if idx == N_QUARTERS - 1 and previous_signature == signature:
            continue
        previous_signature = signature

        rows.append(
            {
                "model": model_info["model"],
                "ticker": model_info["ticker"],
                "model_period": model_info["model_period"],
                "model_date": model_info["model_date"],
                "method": "regression",
                "parameter_name": "num_quarters_used",
                "parameter_value": num_quarters_used if num_quarters_used is not None else "",
                "num_quarters_used": num_quarters_used if num_quarters_used is not None else "",
                "forecast_value": format_output_value(forecast_value),
                "actual_value": "" if is_blank(actual_value) else format_output_value(actual_value),
                "forecast_max": format_output_value(forecast_max),
                "forecast_min": format_output_value(forecast_min),
                "range_width": format_output_value(subtract_or_blank(forecast_max, forecast_min)),
                "intercept": format_output_value(intercept_value),
                "slope": format_output_value(slope_value),
                "source_file": source_file,
            }
        )

    return rows


def write_sheet(
    workbook: Workbook,
    title: str,
    columns: List[str],
    rows: List[Dict[str, Any]],
) -> None:
    ws = workbook.create_sheet(title=title)
    ws.append(columns)

    for row in rows:
        ws.append([row.get(column, "") for column in columns])

    for cell in ws[1]:
        cell.font = Font(bold=True)

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions

    for col_idx, column_name in enumerate(columns, start=1):
        max_len = len(column_name)
        for row_idx in range(2, ws.max_row + 1):
            value = ws.cell(row=row_idx, column=col_idx).value
            if value is None:
                continue
            max_len = max(max_len, len(str(value)))
        ws.column_dimensions[get_column_letter(col_idx)].width = min(max(max_len + 2, 12), 48)


def write_output_workbook(
    output_path: Path,
    empirical_rows: List[Dict[str, Any]],
    regression_rows: List[Dict[str, Any]],
) -> None:
    wb = Workbook()
    default_sheet = wb.active
    wb.remove(default_sheet)

    write_sheet(wb, "empirical_candidates", EMPIRICAL_COLUMNS, empirical_rows)
    write_sheet(wb, "regression_candidates", REGRESSION_COLUMNS, regression_rows)
    wb.save(output_path)


def collect_source_files(src_dir: Path) -> List[Path]:
    candidates: List[Path] = []
    for path in sorted(src_dir.iterdir(), key=lambda p: p.name.lower()):
        if not path.is_file():
            print(f"Skipped {path.name}: not a file.")
            continue
        if path.name.startswith("~"):
            print(f"Skipped {path.name}: temporary Excel file.")
            continue
        if path.suffix.lower() != ".xlsx":
            print(f"Skipped {path.name}: not an .xlsx file.")
            continue
        candidates.append(path)
    return candidates


def run() -> int:
    missing_deps: List[str] = []
    if xw is None:
        missing_deps.append("xlwings")
    if Workbook is None or Font is None or get_column_letter is None:
        missing_deps.append("openpyxl")
    if missing_deps:
        print(
            "Missing dependency packages: "
            + ", ".join(missing_deps)
            + ". Install with `pip install "
            + " ".join(missing_deps)
            + "`."
        )
        return 1

    src_dir = input_dir.expanduser().resolve()
    dst_dir = output_dir.expanduser().resolve()
    dst_dir.mkdir(parents=True, exist_ok=True)

    if not src_dir.exists():
        print(f"Input directory does not exist: {src_dir}")
        return 1

    source_files = collect_source_files(src_dir)
    output_path = get_unique_output_path(src_dir, dst_dir)

    empirical_rows: List[Dict[str, Any]] = []
    regression_rows: List[Dict[str, Any]] = []
    files_processed = 0

    app = xw.App(visible=False, add_book=False)
    app.display_alerts = False
    app.screen_updating = False
    app.enable_events = False
    app.calculation = "manual"

    try:
        for file_path in source_files:
            print(f"Processing {file_path.name}")
            workbook: Optional[xw.Book] = None

            try:
                workbook = app.books.open(str(file_path), update_links=False)
                model_info = parse_file_labels(file_path.name)

                empirical_rows.extend(
                    process_empirical_sheet(
                        workbook=workbook,
                        model_info=model_info,
                        source_file=file_path.name,
                    )
                )
                regression_rows.extend(
                    process_regression_sheet(
                        workbook=workbook,
                        model_info=model_info,
                        source_file=file_path.name,
                    )
                )
                files_processed += 1
            except Exception as exc:
                print(f"Skipped {file_path.name}: {exc}")
            finally:
                if workbook is not None:
                    safe_close_workbook(workbook)
    finally:
        app.quit()

    write_output_workbook(output_path, empirical_rows, regression_rows)

    print(f"Output path: {output_path}")
    print(f"Files processed: {files_processed}")
    print(f"Empirical rows: {len(empirical_rows)}")
    print(f"Regression rows: {len(regression_rows)}")
    return 0


if __name__ == "__main__":
    raise SystemExit(run())
