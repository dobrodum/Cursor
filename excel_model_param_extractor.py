#!/usr/bin/env python3
"""
Extract empirical and regression candidate rows from all .xlsx files in a folder.

Requirements implemented:
- One hidden Excel app instance for the full run.
- Each source workbook is opened once, both model sheets are processed, then closed.
- Source workbooks are never saved or modified permanently.
- Output workbook has two sheets:
    empirical_candidates
    regression_candidates
"""

from __future__ import annotations

import math
import re
from datetime import date
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Sequence, Tuple

import xlwings as xw
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

# ---------------------------------------------------------------------------
# Configure these two paths before running.
# ---------------------------------------------------------------------------
input_dir = "/path/to/input"
output_dir = "/path/to/output"

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


MONTH_TO_NUMBER = {
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

# Header aliases are normalized by stripping non-alphanumeric characters.
HEADER_ALIASES = {
    "num_quarters_used": {
        "numquartersused",
        "numquarters",
        "nquarters",
        "quartersused",
    },
    "last_quarter_used": {"lastquarterused", "lastquarter", "quarter"},
    "forecast_value_empirical": {
        "estimatedtotalsold",
        "forecastvalue",
        "forecast",
        "totfcstwosa",
    },
    "forecast_value_regression": {
        "totfcstwosa",
        "totfcstwithoutsa",
        "forecasttotalwithoutsa",
        "forecastvalue",
        "forecast",
    },
    "actual_value": {"actualvalue", "actual", "reportedsales"},
    "forecast_min": {"min", "forecastmin"},
    "quarterly_sales": {"quarterlysales", "quartersales"},
    "reported_sales": {"reportedsales", "salesreported"},
    "growth_rate_pct": {"growthratepct", "growthrate"},
    "sales_captured_in_db_pct": {
        "salescapturedindbpct",
        "salescaptureddbpct",
        "capturedindbpct",
    },
    "avg_penetration_pct": {"avgpenetrationpct", "penetrationpct", "avgpenetration"},
    "quarter": {"quarter", "period", "fiscalquarter"},
}


def normalize_header(value: Any) -> str:
    if value is None:
        return ""
    return re.sub(r"[^a-z0-9]+", "", str(value).strip().lower())


def as_2d(values: Any) -> List[List[Any]]:
    if values is None:
        return []
    if not isinstance(values, list):
        return [[values]]
    if values and not isinstance(values[0], list):
        return [values]
    return values


def is_number(value: Any) -> bool:
    if isinstance(value, bool) or value is None:
        return False
    if isinstance(value, (int, float)):
        return not (isinstance(value, float) and math.isnan(value))
    if isinstance(value, str):
        text = value.strip().replace(",", "")
        if not text:
            return False
        if text.endswith("%"):
            text = text[:-1]
        try:
            float(text)
            return True
        except ValueError:
            return False
    return False


def to_float(value: Any) -> Optional[float]:
    if not is_number(value):
        return None
    if isinstance(value, str):
        text = value.strip().replace(",", "")
        if text.endswith("%"):
            text = text[:-1]
        return float(text)
    return float(value)


def to_int(value: Any) -> Optional[int]:
    number = to_float(value)
    if number is None:
        return None
    if abs(number - round(number)) <= 1e-9:
        return int(round(number))
    return None


def safe_subtract(a: Any, b: Any) -> Optional[float]:
    fa = to_float(a)
    fb = to_float(b)
    if fa is None or fb is None:
        return None
    return fa - fb


def parse_file_label(file_name: str) -> Dict[str, str]:
    """
    Example:
      MedMiner_Model - AORT - MidJan2026_Send.xlsx
    ->
      ticker: AORT
      model_period: MidJan_2026
      model_date: 2026-01-15
      model: AORT_MidJan_2026
    """
    stem = Path(file_name).stem
    parts = [p.strip() for p in stem.split("-")]

    ticker = "UNKNOWN"
    if len(parts) >= 2 and parts[1]:
        ticker = re.sub(r"\s+", "", parts[1]).upper()

    # Match both "MidJan2026" and "Mid Jan 2026" styles.
    period_match = re.search(
        r"(Early|Mid|Late)\s*([A-Za-z]{3,9})\s*(20\d{2})", stem, flags=re.IGNORECASE
    )

    model_period = "unknown_period"
    model_date = ""

    if period_match:
        phase_raw, month_raw, year_raw = period_match.groups()
        phase = phase_raw.capitalize()
        month_abbrev = month_raw[:3].capitalize()
        month_key = month_abbrev.lower()
        year = int(year_raw)

        month_num = MONTH_TO_NUMBER.get(month_key)
        if month_num is not None:
            day = PERIOD_DAY[phase.lower()]
            model_period = f"{phase}{month_abbrev}_{year}"
            model_date = date(year, month_num, day).isoformat()

    model = f"{ticker}_{model_period}"
    return {
        "model": model,
        "ticker": ticker,
        "model_period": model_period,
        "model_date": model_date,
    }


def next_output_path(input_folder: Path, out_folder: Path) -> Path:
    base = f"{input_folder.name}_PARAM"
    first = out_folder / f"{base}.xlsx"
    if not first.exists():
        return first

    index = 1
    while True:
        candidate = out_folder / f"{base}.{index}.xlsx"
        if not candidate.exists():
            return candidate
        index += 1


def safe_close_workbook(workbook: xw.Book) -> None:
    """
    Close workbook without saving, with fallbacks across Excel backends.
    """
    try:
        workbook.close(save=False)
        return
    except TypeError:
        # Some backends may not accept "save" keyword.
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


def find_anchor_max(sheet: xw.Sheet) -> Optional[Tuple[int, int]]:
    used = sheet.used_range
    values = as_2d(used.value)
    if not values:
        return None

    base_row = used.row
    base_col = used.column
    for r_offset, row_values in enumerate(values):
        for c_offset, value in enumerate(row_values):
            if isinstance(value, str) and value.strip().lower() == "max":
                return base_row + r_offset, base_col + c_offset
    return None


def build_header_map(
    sheet: xw.Sheet, header_row: int, start_col: int, end_col: int
) -> Dict[str, int]:
    values = sheet.range((header_row, start_col), (header_row, end_col)).value
    if not isinstance(values, list):
        values = [values]

    mapping: Dict[str, int] = {}
    for idx, value in enumerate(values):
        key = normalize_header(value)
        if key and key not in mapping:
            mapping[key] = start_col + idx
    return mapping


def pick_column(
    header_map: Dict[str, int], alias_keys: Iterable[str], fallback: int
) -> int:
    for key in alias_keys:
        col = header_map.get(key)
        if col is not None:
            return col
    return max(1, fallback)


def read_cell(sheet: xw.Sheet, row: int, col: int) -> Any:
    if row < 1 or col < 1:
        return None
    return sheet.cells(row, col).value


def last_numeric_row(sheet: xw.Sheet, col: int, upper_row: int) -> Optional[int]:
    if col < 1 or upper_row < 1:
        return None

    values = sheet.range((1, col), (upper_row, col)).value
    if not isinstance(values, list):
        values = [values]

    for idx in range(len(values) - 1, -1, -1):
        if is_number(values[idx]):
            return idx + 1
    return None


def set_formula_r1c1(cell: xw.Range, formula: str) -> None:
    try:
        cell.formula2 = formula
    except Exception:
        # Fallback for environments where formula2 is unsupported.
        cell.formula = formula


def build_result_lookup(
    sheet: xw.Sheet, anchor_row: int, cols: Sequence[int], num_col: int, n_quarters: int
) -> Dict[int, Dict[int, Any]]:
    active_cols = sorted({col for col in cols if col >= 1})
    if not active_cols or num_col < 1:
        return {}

    start_col = min(active_cols)
    end_col = max(active_cols)
    start_row = anchor_row + 1
    end_row = anchor_row + n_quarters + 12

    block = as_2d(sheet.range((start_row, start_col), (end_row, end_col)).value)
    lookup: Dict[int, Dict[int, Any]] = {}

    for row_values in block:
        by_col: Dict[int, Any] = {}
        for col in active_cols:
            rel = col - start_col
            by_col[col] = row_values[rel] if rel < len(row_values) else None

        n_value = to_int(by_col.get(num_col))
        if n_value is None or not (1 <= n_value <= n_quarters):
            continue
        if n_value not in lookup:
            lookup[n_value] = by_col

    return lookup


def compute_empirical_avg_penetration(
    sheet: xw.Sheet, penetration_col: int, data_end_row: Optional[int], n_quarters: int
) -> List[Any]:
    if data_end_row is None or data_end_row < 1 or penetration_col < 1:
        return [None] * n_quarters

    helper_col = 16384  # XFD
    max_n = min(n_quarters, data_end_row)

    for n in range(1, n_quarters + 1):
        if n <= max_n:
            start_row = data_end_row - n + 1
            formula = (
                f'=IFERROR(AVERAGE(R{start_row}C{penetration_col}:'
                f"R{data_end_row}C{penetration_col}),\"\")"
            )
        else:
            formula = '=""'
        set_formula_r1c1(sheet.cells(n, helper_col), formula)

    sheet.book.app.calculate()

    values = sheet.range((1, helper_col), (n_quarters, helper_col)).value
    if not isinstance(values, list):
        values = [values]

    sheet.range((1, helper_col), (n_quarters, helper_col)).clear_contents()
    return values


def compute_regression_params(
    sheet: xw.Sheet,
    y_col: int,
    x_col: int,
    data_end_row: Optional[int],
    n_quarters: int,
) -> Tuple[List[Any], List[Any]]:
    if (
        data_end_row is None
        or data_end_row < 1
        or y_col < 1
        or x_col < 1
        or x_col == y_col
    ):
        return [None] * n_quarters, [None] * n_quarters

    helper_intercept_col = 16384  # XFD
    helper_slope_col = 16383  # XFC
    max_n = min(n_quarters, data_end_row)

    for n in range(1, n_quarters + 1):
        if n <= max_n:
            start_row = data_end_row - n + 1
            intercept_formula = (
                f'=IFERROR(INTERCEPT(R{start_row}C{y_col}:R{data_end_row}C{y_col},'
                f"R{start_row}C{x_col}:R{data_end_row}C{x_col}),\"\")"
            )
            slope_formula = (
                f'=IFERROR(SLOPE(R{start_row}C{y_col}:R{data_end_row}C{y_col},'
                f"R{start_row}C{x_col}:R{data_end_row}C{x_col}),\"\")"
            )
        else:
            intercept_formula = '=""'
            slope_formula = '=""'

        set_formula_r1c1(sheet.cells(n, helper_intercept_col), intercept_formula)
        set_formula_r1c1(sheet.cells(n, helper_slope_col), slope_formula)

    sheet.book.app.calculate()

    intercept_values = sheet.range(
        (1, helper_intercept_col), (n_quarters, helper_intercept_col)
    ).value
    slope_values = sheet.range((1, helper_slope_col), (n_quarters, helper_slope_col)).value

    if not isinstance(intercept_values, list):
        intercept_values = [intercept_values]
    if not isinstance(slope_values, list):
        slope_values = [slope_values]

    sheet.range((1, helper_intercept_col), (n_quarters, helper_intercept_col)).clear_contents()
    sheet.range((1, helper_slope_col), (n_quarters, helper_slope_col)).clear_contents()
    return intercept_values, slope_values


def extract_empirical_rows(
    workbook: xw.Book, labels: Dict[str, str], source_file: str
) -> List[Dict[str, Any]]:
    try:
        sheet = workbook.sheets["Empirical Model"]
    except Exception:
        print(f"Skipped sheet in {source_file}: missing 'Empirical Model'")
        return []

    anchor = find_anchor_max(sheet)
    if anchor is None:
        print(f"Skipped sheet in {source_file}: could not find 'max' anchor in Empirical Model")
        return []

    anchor_row, anchor_col = anchor
    header_map = build_header_map(
        sheet,
        header_row=anchor_row,
        start_col=max(1, anchor_col - 24),
        end_col=anchor_col + 12,
    )

    num_col = pick_column(
        header_map, HEADER_ALIASES["num_quarters_used"], fallback=anchor_col - 12
    )
    last_quarter_col = pick_column(
        header_map, HEADER_ALIASES["last_quarter_used"], fallback=anchor_col - 11
    )
    forecast_col = pick_column(
        header_map, HEADER_ALIASES["forecast_value_empirical"], fallback=anchor_col - 3
    )
    actual_col = pick_column(header_map, HEADER_ALIASES["actual_value"], fallback=anchor_col - 2)
    min_col = pick_column(header_map, HEADER_ALIASES["forecast_min"], fallback=anchor_col + 1)
    quarterly_sales_col = pick_column(
        header_map, HEADER_ALIASES["quarterly_sales"], fallback=anchor_col - 10
    )
    reported_sales_col = pick_column(
        header_map, HEADER_ALIASES["reported_sales"], fallback=actual_col
    )
    growth_rate_col = pick_column(
        header_map, HEADER_ALIASES["growth_rate_pct"], fallback=anchor_col - 8
    )
    captured_col = pick_column(
        header_map, HEADER_ALIASES["sales_captured_in_db_pct"], fallback=anchor_col - 7
    )
    penetration_col = pick_column(
        header_map, HEADER_ALIASES["avg_penetration_pct"], fallback=anchor_col - 6
    )
    quarter_col = pick_column(header_map, HEADER_ALIASES["quarter"], fallback=last_quarter_col)

    data_end_row = last_numeric_row(sheet, penetration_col, anchor_row - 1)
    if data_end_row is None:
        data_end_row = last_numeric_row(sheet, quarterly_sales_col, anchor_row - 1)

    avg_penetration_values = compute_empirical_avg_penetration(
        sheet, penetration_col, data_end_row, N_QUARTERS
    )

    result_lookup = build_result_lookup(
        sheet=sheet,
        anchor_row=anchor_row,
        cols=[num_col, forecast_col, actual_col, anchor_col, min_col],
        num_col=num_col,
        n_quarters=N_QUARTERS,
    )

    last_quarter_used = read_cell(sheet, data_end_row or 0, quarter_col)
    quarterly_sales = read_cell(sheet, data_end_row or 0, quarterly_sales_col)
    reported_sales = read_cell(sheet, data_end_row or 0, reported_sales_col)
    growth_rate_pct = read_cell(sheet, data_end_row or 0, growth_rate_col)
    captured_pct = read_cell(sheet, data_end_row or 0, captured_col)

    rows: List[Dict[str, Any]] = []
    for n in range(1, N_QUARTERS + 1):
        row_map = result_lookup.get(n, {})
        fallback_row = anchor_row + n

        forecast_value = row_map.get(forecast_col, read_cell(sheet, fallback_row, forecast_col))
        actual_value = row_map.get(actual_col, read_cell(sheet, fallback_row, actual_col))
        forecast_max = row_map.get(anchor_col, read_cell(sheet, fallback_row, anchor_col))
        forecast_min = row_map.get(min_col, read_cell(sheet, fallback_row, min_col))
        avg_penetration = (
            avg_penetration_values[n - 1] if n - 1 < len(avg_penetration_values) else None
        )

        if actual_value in (None, ""):
            actual_value = reported_sales

        if all(
            v in (None, "")
            for v in (avg_penetration, forecast_value, forecast_max, forecast_min, actual_value)
        ):
            continue

        rows.append(
            {
                "model": labels["model"],
                "ticker": labels["ticker"],
                "model_period": labels["model_period"],
                "model_date": labels["model_date"],
                "method": "empirical",
                "parameter_name": "avg_penetration_pct",
                "parameter_value": avg_penetration,
                "num_quarters_used": n,
                "last_quarter_used": last_quarter_used,
                "forecast_value": forecast_value,
                "actual_value": actual_value,
                "forecast_max": forecast_max,
                "forecast_min": forecast_min,
                "range_width": safe_subtract(forecast_max, forecast_min),
                "avg_penetration_pct": avg_penetration,
                "quarterly_sales": quarterly_sales,
                "reported_sales": reported_sales,
                "growth_rate_pct": growth_rate_pct,
                "sales_captured_in_db_pct": captured_pct,
                "source_file": source_file,
            }
        )

    return rows


def extract_regression_rows(
    workbook: xw.Book, labels: Dict[str, str], source_file: str
) -> List[Dict[str, Any]]:
    try:
        sheet = workbook.sheets["Regression Model"]
    except Exception:
        print(f"Skipped sheet in {source_file}: missing 'Regression Model'")
        return []

    anchor = find_anchor_max(sheet)
    if anchor is None:
        print(f"Skipped sheet in {source_file}: could not find 'max' anchor in Regression Model")
        return []

    anchor_row, anchor_col = anchor
    header_map = build_header_map(
        sheet,
        header_row=anchor_row,
        start_col=max(1, anchor_col - 24),
        end_col=anchor_col + 12,
    )

    num_col = pick_column(
        header_map, HEADER_ALIASES["num_quarters_used"], fallback=anchor_col - 9
    )
    forecast_col = pick_column(
        header_map, HEADER_ALIASES["forecast_value_regression"], fallback=anchor_col - 2
    )
    min_col = pick_column(header_map, HEADER_ALIASES["forecast_min"], fallback=anchor_col + 1)

    y_col = anchor_col - 7
    x_col = anchor_col - 11
    data_end_row = last_numeric_row(sheet, y_col, anchor_row - 1)

    intercept_values, slope_values = compute_regression_params(
        sheet=sheet,
        y_col=y_col,
        x_col=x_col,
        data_end_row=data_end_row,
        n_quarters=N_QUARTERS,
    )

    result_lookup = build_result_lookup(
        sheet=sheet,
        anchor_row=anchor_row,
        cols=[num_col, forecast_col, anchor_col, min_col],
        num_col=num_col,
        n_quarters=N_QUARTERS,
    )

    rows: List[Dict[str, Any]] = []
    previous_signature: Optional[Tuple[Any, ...]] = None

    for n in range(1, N_QUARTERS + 1):
        row_map = result_lookup.get(n, {})
        fallback_row = anchor_row + n

        forecast_value = row_map.get(forecast_col, read_cell(sheet, fallback_row, forecast_col))
        forecast_max = row_map.get(anchor_col, read_cell(sheet, fallback_row, anchor_col))
        forecast_min = row_map.get(min_col, read_cell(sheet, fallback_row, min_col))
        intercept = intercept_values[n - 1] if n - 1 < len(intercept_values) else None
        slope = slope_values[n - 1] if n - 1 < len(slope_values) else None

        if all(v in (None, "") for v in (forecast_value, forecast_max, forecast_min, intercept, slope)):
            continue

        signature = (
            round(to_float(intercept) or 0.0, 10),
            round(to_float(slope) or 0.0, 10),
            round(to_float(forecast_value) or 0.0, 10),
            round(to_float(forecast_max) or 0.0, 10),
            round(to_float(forecast_min) or 0.0, 10),
        )

        # Prevent duplicate terminal rows (common in some templates).
        if signature == previous_signature:
            continue
        previous_signature = signature

        rows.append(
            {
                "model": labels["model"],
                "ticker": labels["ticker"],
                "model_period": labels["model_period"],
                "model_date": labels["model_date"],
                "method": "regression",
                "parameter_name": "num_quarters_used",
                "parameter_value": n,
                "num_quarters_used": n,
                "forecast_value": forecast_value,
                "actual_value": "",
                "forecast_max": forecast_max,
                "forecast_min": forecast_min,
                "range_width": safe_subtract(forecast_max, forecast_min),
                "intercept": intercept,
                "slope": slope,
                "source_file": source_file,
            }
        )

    return rows


def write_sheet(ws: Any, columns: Sequence[str], rows: Sequence[Dict[str, Any]]) -> None:
    ws.append(list(columns))
    for record in rows:
        ws.append([record.get(col, "") for col in columns])

    header_font = Font(bold=True)
    for cell in ws[1]:
        cell.font = header_font

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions

    for col_idx, col_name in enumerate(columns, start=1):
        width = len(col_name)
        for row_idx in range(2, ws.max_row + 1):
            value = ws.cell(row=row_idx, column=col_idx).value
            if value is None:
                continue
            width = max(width, len(str(value)))
        ws.column_dimensions[get_column_letter(col_idx)].width = min(max(width + 2, 12), 48)


def write_output_workbook(
    output_path: Path,
    empirical_rows: Sequence[Dict[str, Any]],
    regression_rows: Sequence[Dict[str, Any]],
) -> None:
    wb = Workbook()
    ws_empirical = wb.active
    ws_empirical.title = "empirical_candidates"
    ws_regression = wb.create_sheet("regression_candidates")

    write_sheet(ws_empirical, EMPIRICAL_COLUMNS, empirical_rows)
    write_sheet(ws_regression, REGRESSION_COLUMNS, regression_rows)

    wb.save(output_path)


def should_skip_file(path: Path) -> Optional[str]:
    if not path.is_file():
        return "not a regular file"
    if path.name.startswith("~"):
        return "temporary file"
    if path.suffix.lower() != ".xlsx":
        return "not .xlsx"
    return None


def process_workbooks(in_dir: Path, out_dir: Path) -> Tuple[Path, int, int, int]:
    output_path = next_output_path(in_dir, out_dir)
    empirical_rows: List[Dict[str, Any]] = []
    regression_rows: List[Dict[str, Any]] = []
    files_processed = 0

    app: Optional[xw.App] = None
    try:
        app = xw.App(visible=False, add_book=False)
        app.display_alerts = False
        app.screen_updating = False

        for file_path in sorted(in_dir.iterdir()):
            skip_reason = should_skip_file(file_path)
            if skip_reason:
                print(f"Skipped {file_path.name}: {skip_reason}")
                continue

            workbook: Optional[xw.Book] = None
            try:
                workbook = app.books.open(str(file_path), update_links=False)
                labels = parse_file_label(file_path.name)

                empirical_chunk = extract_empirical_rows(workbook, labels, file_path.name)
                regression_chunk = extract_regression_rows(workbook, labels, file_path.name)
                empirical_rows.extend(empirical_chunk)
                regression_rows.extend(regression_chunk)

                files_processed += 1
                print(
                    f"Processed {file_path.name}: "
                    f"{len(empirical_chunk)} empirical rows, {len(regression_chunk)} regression rows"
                )
            except Exception as exc:
                print(f"Skipped {file_path.name}: processing error ({exc})")
            finally:
                if workbook is not None:
                    safe_close_workbook(workbook)
    finally:
        if app is not None:
            app.quit()

    write_output_workbook(output_path, empirical_rows, regression_rows)
    return output_path, files_processed, len(empirical_rows), len(regression_rows)


def main() -> None:
    in_dir = Path(input_dir).expanduser()
    out_dir = Path(output_dir).expanduser()

    if not in_dir.exists() or not in_dir.is_dir():
        raise SystemExit(f"Input directory does not exist or is not a directory: {in_dir}")
    out_dir.mkdir(parents=True, exist_ok=True)

    output_path, files_processed, empirical_count, regression_count = process_workbooks(
        in_dir, out_dir
    )

    print(f"Output path: {output_path}")
    print(f"Number of files processed: {files_processed}")
    print(f"Number of empirical rows: {empirical_count}")
    print(f"Number of regression rows: {regression_count}")


if __name__ == "__main__":
    main()
