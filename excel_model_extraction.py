#!/usr/bin/env python3
"""Extract empirical/regression model candidates from all .xlsx files in a folder.

This script opens each source workbook once, processes both model sheets while the
book is open, and writes one combined output workbook with:
  - empirical_candidates
  - regression_candidates
"""

from __future__ import annotations

from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path
import re
from typing import Any, Dict, List, Optional, Sequence, Tuple

import xlwings as xw
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter


# ---------------------------------------------------------------------------
# User-configurable inputs
# ---------------------------------------------------------------------------
input_dir = Path("/path/to/input")
output_dir = Path("/path/to/output")


# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------
N_QUARTERS = 10
EMPIRICAL_SHEET_NAME = "Empirical Model"
REGRESSION_SHEET_NAME = "Regression Model"

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

MONTH_TO_INT = {
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

PERIOD_DAY = {"Early": 5, "Mid": 15, "Late": 25}
PERIOD_RE = re.compile(
    r"(Early|Mid|Late)\s*[-_ ]?\s*(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\s*[-_ ]?\s*(\d{4})",
    flags=re.IGNORECASE,
)


@dataclass(frozen=True)
class ModelMetadata:
    model: str
    ticker: str
    model_period: str
    model_date: str


@dataclass(frozen=True)
class SheetAnchor:
    anchor_row: int
    anchor_col: int
    left_col: int
    right_col: int


def normalize_header(value: Any) -> str:
    if value is None:
        return ""
    return re.sub(r"[^a-z0-9]+", " ", str(value).strip().lower()).strip()


def is_blank(value: Any) -> bool:
    if value is None:
        return True
    if isinstance(value, str) and value.strip() == "":
        return True
    return False


def to_float(value: Any) -> Optional[float]:
    if value is None:
        return None
    if isinstance(value, bool):
        return float(value)
    if isinstance(value, (int, float)):
        return float(value)
    if isinstance(value, str):
        text = value.strip()
        if text == "":
            return None
        text = text.replace(",", "").replace("%", "")
        try:
            return float(text)
        except ValueError:
            return None
    return None


def to_int(value: Any) -> Optional[int]:
    number = to_float(value)
    if number is None:
        return None
    return int(round(number))


def range_width(max_value: Any, min_value: Any) -> Optional[float]:
    max_num = to_float(max_value)
    min_num = to_float(min_value)
    if max_num is None or min_num is None:
        return None
    return max_num - min_num


def normalize_output_value(value: Any) -> Any:
    if isinstance(value, datetime):
        return value.date().isoformat()
    if isinstance(value, date):
        return value.isoformat()
    return value


def signature_value(value: Any) -> Any:
    number = to_float(value)
    if number is not None:
        return round(number, 10)
    if value is None:
        return None
    if isinstance(value, (datetime, date)):
        return normalize_output_value(value)
    return str(value).strip()


def ensure_matrix(raw: Any, rows: int, cols: int) -> List[List[Any]]:
    if rows <= 0 or cols <= 0:
        return []

    if rows == 1 and cols == 1:
        return [[raw]]

    if rows == 1:
        if isinstance(raw, list):
            if raw and isinstance(raw[0], list):
                return [raw[0]]
            return [raw]
        return [[raw]]

    if cols == 1:
        if isinstance(raw, list):
            if raw and isinstance(raw[0], list):
                return raw
            return [[v] for v in raw]
        return [[raw]] + [[None] for _ in range(rows - 1)]

    if isinstance(raw, list) and raw and isinstance(raw[0], list):
        return raw

    return [[None for _ in range(cols)] for _ in range(rows)]


def read_range_matrix(
    sheet: xw.Sheet, row_start: int, col_start: int, row_end: int, col_end: int
) -> List[List[Any]]:
    row_count = row_end - row_start + 1
    col_count = col_end - col_start + 1
    raw = sheet.range((row_start, col_start), (row_end, col_end)).value
    return ensure_matrix(raw, row_count, col_count)


def read_column_values(sheet: xw.Sheet, row_start: int, row_end: int, col: int) -> List[Any]:
    row_count = row_end - row_start + 1
    raw = sheet.range((row_start, col), (row_end, col)).value
    if row_count == 1:
        return [raw]
    if isinstance(raw, list):
        if raw and isinstance(raw[0], list):
            return [r[0] for r in raw]
        return raw
    return [raw] + [None for _ in range(row_count - 1)]


def get_matrix_value(matrix: List[List[Any]], row_idx: int, col_idx: int) -> Any:
    if row_idx < 0 or row_idx >= len(matrix):
        return None
    row = matrix[row_idx]
    if col_idx < 0 or col_idx >= len(row):
        return None
    return row[col_idx]


def set_formula2(cell: xw.Range, formula: str) -> None:
    try:
        cell.formula2 = formula
    except Exception:
        cell.formula = formula


def parse_model_metadata(file_name: str) -> ModelMetadata:
    stem = Path(file_name).stem
    parts = [part.strip() for part in stem.split(" - ")]
    ticker = parts[1].upper() if len(parts) > 1 and parts[1] else "UNKNOWN"

    match = PERIOD_RE.search(stem)
    if not match:
        model_period = "UnknownPeriod"
        model_date = ""
        model = f"{ticker}_{model_period}"
        return ModelMetadata(model=model, ticker=ticker, model_period=model_period, model_date=model_date)

    stage = match.group(1).title()
    month_abbrev = match.group(2).title()[:3]
    year = int(match.group(3))
    month_num = MONTH_TO_INT[month_abbrev.lower()]
    day_num = PERIOD_DAY[stage]

    model_period = f"{stage}{month_abbrev}_{year}"
    model_date = date(year, month_num, day_num).isoformat()
    model = f"{ticker}_{model_period}"

    return ModelMetadata(model=model, ticker=ticker, model_period=model_period, model_date=model_date)


def next_output_path(in_dir: Path, out_dir: Path) -> Path:
    base_name = f"{in_dir.name}_PARAM.xlsx"
    base_path = out_dir / base_name
    if not base_path.exists():
        return base_path

    counter = 1
    while True:
        candidate = out_dir / f"{in_dir.name}_PARAM.{counter}.xlsx"
        if not candidate.exists():
            return candidate
        counter += 1


def safe_close_workbook(workbook: Optional[xw.Book]) -> None:
    if workbook is None:
        return
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
        workbook.api.Close(False)
    except Exception:
        try:
            workbook.api.Close(SaveChanges=False)
        except Exception:
            pass


def find_anchor(sheet: xw.Sheet, label: str) -> Optional[SheetAnchor]:
    used = sheet.used_range
    if used is None:
        return None

    top_row = used.row
    left_col = used.column
    row_count = used.rows.count
    col_count = used.columns.count
    right_col = left_col + col_count - 1
    bottom_row = top_row + row_count - 1

    matrix = read_range_matrix(sheet, top_row, left_col, bottom_row, right_col)
    target = normalize_header(label)

    for r_idx in range(row_count):
        for c_idx in range(col_count):
            cell_value = get_matrix_value(matrix, r_idx, c_idx)
            if isinstance(cell_value, str) and normalize_header(cell_value) == target:
                return SheetAnchor(
                    anchor_row=top_row + r_idx,
                    anchor_col=left_col + c_idx,
                    left_col=left_col,
                    right_col=right_col,
                )
    return None


def build_header_map(sheet: xw.Sheet, row: int, left_col: int, right_col: int) -> Dict[int, str]:
    raw = sheet.range((row, left_col), (row, right_col)).value
    headers = raw if isinstance(raw, list) else [raw]
    mapping: Dict[int, str] = {}
    for idx, value in enumerate(headers):
        normalized = normalize_header(value)
        if normalized:
            mapping[left_col + idx] = normalized
    return mapping


def find_offset(
    header_map: Dict[int, str],
    anchor_col: int,
    term_groups: Sequence[Sequence[str]],
    default: Optional[int],
) -> Optional[int]:
    candidates: List[int] = []
    for col, header_text in header_map.items():
        for terms in term_groups:
            if all(term in header_text for term in terms):
                candidates.append(col)
                break
    if not candidates:
        return default
    best_col = min(candidates, key=lambda col: abs(col - anchor_col))
    return best_col - anchor_col


def get_sheet_if_exists(workbook: xw.Book, sheet_name: str) -> Optional[xw.Sheet]:
    try:
        return workbook.sheets[sheet_name]
    except Exception:
        return None


def process_empirical_sheet(workbook: xw.Book, metadata: ModelMetadata, source_file: str) -> List[Dict[str, Any]]:
    rows: List[Dict[str, Any]] = []
    sheet = get_sheet_if_exists(workbook, EMPIRICAL_SHEET_NAME)
    if sheet is None:
        print(f"  - skipped empirical: missing sheet '{EMPIRICAL_SHEET_NAME}'")
        return rows

    anchor = find_anchor(sheet, "max")
    if anchor is None:
        print("  - skipped empirical: could not find 'max' anchor")
        return rows

    header_map = build_header_map(sheet, anchor.anchor_row, anchor.left_col, anchor.right_col)

    offsets = {
        "num_quarters_used": find_offset(
            header_map,
            anchor.anchor_col,
            (("num", "quarter"), ("quarters", "used")),
            -6,
        ),
        "last_quarter_used": find_offset(
            header_map,
            anchor.anchor_col,
            (("last", "quarter"),),
            -5,
        ),
        "forecast_value": find_offset(
            header_map,
            anchor.anchor_col,
            (
                ("estimated", "total", "sold"),
                ("estimate", "total", "sold"),
                ("forecast", "value"),
            ),
            -4,
        ),
        "actual_value": find_offset(
            header_map,
            anchor.anchor_col,
            (("reported", "sales"), ("actual", "sales")),
            -3,
        ),
        "forecast_min": find_offset(
            header_map,
            anchor.anchor_col,
            (("min",), ("minimum",)),
            1,
        ),
        "quarterly_sales": find_offset(
            header_map,
            anchor.anchor_col,
            (("quarterly", "sales"), ("quarter", "sales")),
            -10,
        ),
        "reported_sales": find_offset(
            header_map,
            anchor.anchor_col,
            (("reported", "sales"),),
            -9,
        ),
        "growth_rate_pct": find_offset(
            header_map,
            anchor.anchor_col,
            (("growth", "rate"), ("growth",)),
            -8,
        ),
        "sales_captured_in_db_pct": find_offset(
            header_map,
            anchor.anchor_col,
            (("captured", "db"), ("sales", "captured"), ("captured",)),
            -7,
        ),
    }

    row_start = anchor.anchor_row + 1
    row_end = anchor.anchor_row + N_QUARTERS

    required_cols = [anchor.anchor_col]
    for offset in offsets.values():
        if offset is not None:
            required_cols.append(anchor.anchor_col + offset)
    col_start = min(required_cols)
    col_end = max(required_cols)

    matrix = read_range_matrix(sheet, row_start, col_start, row_end, col_end)

    helper_avg_col = anchor.right_col + 1
    sales_capture_col = anchor.anchor_col + (offsets["sales_captured_in_db_pct"] or -7)

    for i in range(N_QUARTERS):
        row = row_start + i
        num_quarters_col = anchor.anchor_col + (offsets["num_quarters_used"] or -6)
        num_quarters_raw = get_matrix_value(matrix, i, num_quarters_col - col_start)
        num_quarters = to_int(num_quarters_raw) or (i + 1)
        num_quarters = max(1, min(num_quarters, N_QUARTERS))
        avg_start_row = max(row_start, row - num_quarters + 1)

        formula = (
            f'=IFERROR(AVERAGE(R{avg_start_row}C{sales_capture_col}:'
            f"R{row}C{sales_capture_col}),\"\")"
        )
        set_formula2(sheet.range((row, helper_avg_col)), formula)

    workbook.app.calculate()
    avg_penetration_values = read_column_values(sheet, row_start, row_end, helper_avg_col)

    for i in range(N_QUARTERS):
        def col_value(offset_key: str) -> Any:
            offset = offsets[offset_key]
            if offset is None:
                return None
            abs_col = anchor.anchor_col + offset
            return get_matrix_value(matrix, i, abs_col - col_start)

        forecast_max = get_matrix_value(matrix, i, anchor.anchor_col - col_start)
        forecast_min = col_value("forecast_min")
        forecast_value = col_value("forecast_value")
        actual_value = col_value("actual_value")
        num_quarters_used = col_value("num_quarters_used")

        if all(
            is_blank(v)
            for v in (num_quarters_used, forecast_value, actual_value, forecast_max, forecast_min)
        ):
            continue

        avg_penetration_pct = avg_penetration_values[i] if i < len(avg_penetration_values) else None

        rows.append(
            {
                "model": metadata.model,
                "ticker": metadata.ticker,
                "model_period": metadata.model_period,
                "model_date": metadata.model_date,
                "method": "empirical",
                "parameter_name": "avg_penetration_pct",
                "parameter_value": normalize_output_value(avg_penetration_pct),
                "num_quarters_used": normalize_output_value(num_quarters_used),
                "last_quarter_used": normalize_output_value(col_value("last_quarter_used")),
                "forecast_value": normalize_output_value(forecast_value),
                "actual_value": normalize_output_value(actual_value),
                "forecast_max": normalize_output_value(forecast_max),
                "forecast_min": normalize_output_value(forecast_min),
                "range_width": range_width(forecast_max, forecast_min),
                "avg_penetration_pct": normalize_output_value(avg_penetration_pct),
                "quarterly_sales": normalize_output_value(col_value("quarterly_sales")),
                "reported_sales": normalize_output_value(col_value("reported_sales")),
                "growth_rate_pct": normalize_output_value(col_value("growth_rate_pct")),
                "sales_captured_in_db_pct": normalize_output_value(
                    col_value("sales_captured_in_db_pct")
                ),
                "source_file": source_file,
            }
        )

    return rows


def process_regression_sheet(workbook: xw.Book, metadata: ModelMetadata, source_file: str) -> List[Dict[str, Any]]:
    rows: List[Dict[str, Any]] = []
    sheet = get_sheet_if_exists(workbook, REGRESSION_SHEET_NAME)
    if sheet is None:
        print(f"  - skipped regression: missing sheet '{REGRESSION_SHEET_NAME}'")
        return rows

    anchor = find_anchor(sheet, "max")
    if anchor is None:
        print("  - skipped regression: could not find 'max' anchor")
        return rows

    header_map = build_header_map(sheet, anchor.anchor_row, anchor.left_col, anchor.right_col)

    offsets = {
        "num_quarters_used": find_offset(
            header_map,
            anchor.anchor_col,
            (("num", "quarter"), ("quarters", "used")),
            -6,
        ),
        "forecast_value": find_offset(
            header_map,
            anchor.anchor_col,
            (
                ("tot", "fcst", "sa"),
                ("tot", "forecast", "sa"),
                ("forecast", "without", "sa"),
            ),
            -4,
        ),
        "actual_value": find_offset(
            header_map,
            anchor.anchor_col,
            (("actual",), ("reported", "sales")),
            None,
        ),
        "forecast_min": find_offset(
            header_map,
            anchor.anchor_col,
            (("min",), ("minimum",)),
            1,
        ),
    }

    y_col = anchor.anchor_col - 7
    x_col = anchor.anchor_col - 11

    row_start = anchor.anchor_row + 1
    row_end = anchor.anchor_row + N_QUARTERS

    required_cols = [anchor.anchor_col, x_col, y_col]
    for offset in offsets.values():
        if offset is not None:
            required_cols.append(anchor.anchor_col + offset)
    col_start = min(required_cols)
    col_end = max(required_cols)

    matrix = read_range_matrix(sheet, row_start, col_start, row_end, col_end)

    helper_intercept_col = anchor.right_col + 1
    helper_slope_col = anchor.right_col + 2

    for i in range(N_QUARTERS):
        row = row_start + i
        num_quarters_col = anchor.anchor_col + (offsets["num_quarters_used"] or -6)
        num_quarters_raw = get_matrix_value(matrix, i, num_quarters_col - col_start)
        num_quarters = to_int(num_quarters_raw) or (i + 1)
        num_quarters = max(2, min(num_quarters, N_QUARTERS))
        reg_start_row = max(row_start, row - num_quarters + 1)

        intercept_formula = (
            f"=IFERROR(INTERCEPT(R{reg_start_row}C{y_col}:R{row}C{y_col},"
            f"R{reg_start_row}C{x_col}:R{row}C{x_col}),\"\")"
        )
        slope_formula = (
            f"=IFERROR(SLOPE(R{reg_start_row}C{y_col}:R{row}C{y_col},"
            f"R{reg_start_row}C{x_col}:R{row}C{x_col}),\"\")"
        )
        set_formula2(sheet.range((row, helper_intercept_col)), intercept_formula)
        set_formula2(sheet.range((row, helper_slope_col)), slope_formula)

    workbook.app.calculate()
    intercept_values = read_column_values(sheet, row_start, row_end, helper_intercept_col)
    slope_values = read_column_values(sheet, row_start, row_end, helper_slope_col)

    previous_signature: Optional[Tuple[Any, ...]] = None

    for i in range(N_QUARTERS):
        def col_value(offset_key: str) -> Any:
            offset = offsets[offset_key]
            if offset is None:
                return None
            abs_col = anchor.anchor_col + offset
            return get_matrix_value(matrix, i, abs_col - col_start)

        num_quarters_used = col_value("num_quarters_used")
        forecast_value = col_value("forecast_value")
        actual_value = col_value("actual_value")
        forecast_max = get_matrix_value(matrix, i, anchor.anchor_col - col_start)
        forecast_min = col_value("forecast_min")

        intercept = intercept_values[i] if i < len(intercept_values) else None
        slope = slope_values[i] if i < len(slope_values) else None

        if all(
            is_blank(v)
            for v in (num_quarters_used, forecast_value, forecast_max, forecast_min, intercept, slope)
        ):
            continue

        current_signature = tuple(
            signature_value(value)
            for value in (num_quarters_used, forecast_value, forecast_max, forecast_min, intercept, slope)
        )
        if previous_signature is not None and current_signature == previous_signature:
            continue
        previous_signature = current_signature

        rows.append(
            {
                "model": metadata.model,
                "ticker": metadata.ticker,
                "model_period": metadata.model_period,
                "model_date": metadata.model_date,
                "method": "regression",
                "parameter_name": "num_quarters_used",
                "parameter_value": normalize_output_value(num_quarters_used),
                "num_quarters_used": normalize_output_value(num_quarters_used),
                "forecast_value": normalize_output_value(forecast_value),
                "actual_value": normalize_output_value(actual_value) if actual_value is not None else "",
                "forecast_max": normalize_output_value(forecast_max),
                "forecast_min": normalize_output_value(forecast_min),
                "range_width": range_width(forecast_max, forecast_min),
                "intercept": normalize_output_value(intercept),
                "slope": normalize_output_value(slope),
                "source_file": source_file,
            }
        )

    return rows


def write_sheet(
    workbook: Workbook, sheet_name: str, columns: Sequence[str], rows: Sequence[Dict[str, Any]]
) -> None:
    ws = workbook.create_sheet(sheet_name)
    ws.append(list(columns))

    for row in rows:
        ws.append([row.get(column) for column in columns])

    for header_cell in ws[1]:
        header_cell.font = Font(bold=True)

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(columns))}{ws.max_row}"

    for col_idx, col_name in enumerate(columns, start=1):
        max_len = len(col_name)
        for row_idx in range(2, ws.max_row + 1):
            value = ws.cell(row=row_idx, column=col_idx).value
            if value is None:
                continue
            max_len = max(max_len, len(str(value)))
        ws.column_dimensions[get_column_letter(col_idx)].width = min(max(max_len + 2, 12), 50)


def write_output_workbook(
    output_path: Path, empirical_rows: Sequence[Dict[str, Any]], regression_rows: Sequence[Dict[str, Any]]
) -> None:
    output_path.parent.mkdir(parents=True, exist_ok=True)
    output_book = Workbook()
    default_sheet = output_book.active
    output_book.remove(default_sheet)

    write_sheet(output_book, "empirical_candidates", EMPIRICAL_COLUMNS, empirical_rows)
    write_sheet(output_book, "regression_candidates", REGRESSION_COLUMNS, regression_rows)
    output_book.save(output_path)


def collect_source_files(in_dir: Path) -> List[Path]:
    sources: List[Path] = []
    for candidate in sorted(in_dir.iterdir(), key=lambda p: p.name.lower()):
        if not candidate.is_file():
            continue

        file_name = candidate.name
        if file_name.startswith("~"):
            print(f"SKIPPED {file_name}: temporary file")
            continue
        if candidate.suffix.lower() != ".xlsx":
            print(f"SKIPPED {file_name}: not an .xlsx file")
            continue
        if re.search(r"_PARAM(?:\.\d+)?\.xlsx$", file_name, flags=re.IGNORECASE):
            print(f"SKIPPED {file_name}: appears to be a prior output file")
            continue

        sources.append(candidate)

    return sources


def main() -> None:
    if not input_dir.exists() or not input_dir.is_dir():
        raise SystemExit(f"Input directory does not exist: {input_dir}")

    output_dir.mkdir(parents=True, exist_ok=True)
    source_files = collect_source_files(input_dir)

    empirical_rows: List[Dict[str, Any]] = []
    regression_rows: List[Dict[str, Any]] = []
    processed_files = 0

    excel_app: Optional[xw.App] = None
    try:
        excel_app = xw.App(visible=False, add_book=False)
        excel_app.display_alerts = False
        excel_app.screen_updating = False
        excel_app.enable_events = False
        excel_app.calculation = "manual"

        for file_path in source_files:
            print(f"PROCESSING {file_path.name}")
            workbook: Optional[xw.Book] = None
            try:
                workbook = excel_app.books.open(str(file_path), update_links=False)
            except Exception as exc:
                print(f"SKIPPED {file_path.name}: open failed ({exc})")
                continue

            try:
                metadata = parse_model_metadata(file_path.name)
                empirical_rows.extend(process_empirical_sheet(workbook, metadata, file_path.name))
                regression_rows.extend(process_regression_sheet(workbook, metadata, file_path.name))
                processed_files += 1
            except Exception as exc:
                print(f"SKIPPED {file_path.name}: processing failed ({exc})")
            finally:
                safe_close_workbook(workbook)
    finally:
        if excel_app is not None:
            try:
                excel_app.quit()
            except Exception:
                try:
                    excel_app.kill()
                except Exception:
                    pass

    out_path = next_output_path(input_dir, output_dir)
    write_output_workbook(out_path, empirical_rows, regression_rows)

    print(f"OUTPUT {out_path}")
    print(f"FILES_PROCESSED {processed_files}")
    print(f"EMPIRICAL_ROWS {len(empirical_rows)}")
    print(f"REGRESSION_ROWS {len(regression_rows)}")


if __name__ == "__main__":
    main()
