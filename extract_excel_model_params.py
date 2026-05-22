#!/usr/bin/env python3
from __future__ import annotations

import re
import sys
from datetime import date
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Sequence, Tuple

from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

try:
    import xlwings as xw
except ImportError as exc:  # pragma: no cover - runtime dependency guard
    raise SystemExit(
        "Missing dependency 'xlwings'. Install it before running this script."
    ) from exc


# === User-configurable paths ===
input_dir = Path("input")
output_dir = Path("output")


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

EMPIRICAL_HEADER_PATTERNS: Dict[str, Sequence[str]] = {
    "num_quarters_used": ("num quarter", "quarters used", "n quarter"),
    "last_quarter_used": ("last quarter", "quarter used"),
    "forecast_value": (
        "estimated total sold",
        "tot fcst",
        "forecast value",
        "forecast",
    ),
    "actual_value": ("reported sales", "actual value", "actual"),
    "forecast_min": ("min",),
    "avg_penetration_pct": ("avg penetration", "average penetration", "penetration"),
    "quarterly_sales": ("quarterly sales", "qtr sales"),
    "reported_sales": ("reported sales",),
    "growth_rate_pct": ("growth rate",),
    "sales_captured_in_db_pct": ("captured in db", "sales captured"),
}

REGRESSION_HEADER_PATTERNS: Dict[str, Sequence[str]] = {
    "num_quarters_used": ("num quarter", "quarters used", "n quarter"),
    "forecast_value": ("tot fcst w/o sa", "tot fcst wo sa", "forecast"),
    "actual_value": ("reported sales", "actual value", "actual"),
    "forecast_min": ("min",),
}

MONTH_MAP = {
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

PART_DAY_MAP = {"early": 5, "mid": 15, "late": 25}


def normalize_text(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, str):
        text = value
    else:
        text = str(value)
    return re.sub(r"\s+", " ", text).strip().lower()


def is_blank(value: Any) -> bool:
    if value is None:
        return True
    if isinstance(value, str) and not value.strip():
        return True
    return False


def to_float(value: Any) -> Optional[float]:
    if is_blank(value):
        return None
    if isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        return float(value)
    text = str(value).strip().replace(",", "")
    if text.endswith("%"):
        try:
            return float(text[:-1]) / 100.0
        except ValueError:
            return None
    try:
        return float(text)
    except ValueError:
        return None


def safe_subtract(a: Any, b: Any) -> Any:
    fa = to_float(a)
    fb = to_float(b)
    if fa is None or fb is None:
        return ""
    return fa - fb


def normalize_matrix(matrix_value: Any) -> List[List[Any]]:
    if matrix_value is None:
        return []
    if isinstance(matrix_value, tuple):
        matrix_value = list(matrix_value)
    if not isinstance(matrix_value, list):
        return [[matrix_value]]
    if not matrix_value:
        return []
    if isinstance(matrix_value[0], tuple):
        matrix_value = [list(row) if isinstance(row, tuple) else row for row in matrix_value]
    if isinstance(matrix_value[0], list):
        return matrix_value
    return [matrix_value]


def matrix_cell(
    matrix: List[List[Any]],
    matrix_start_row: int,
    matrix_start_col: int,
    row: int,
    col: int,
) -> Any:
    r_idx = row - matrix_start_row
    c_idx = col - matrix_start_col
    if r_idx < 0 or c_idx < 0:
        return None
    if r_idx >= len(matrix):
        return None
    row_values = matrix[r_idx]
    if c_idx >= len(row_values):
        return None
    return row_values[c_idx]


def parse_file_label(file_name: str) -> Dict[str, str]:
    stem = Path(file_name).stem
    parts = [segment.strip() for segment in stem.split(" - ")]

    ticker = parts[1].upper() if len(parts) >= 2 else "UNKNOWN"
    period_token_raw = parts[2] if len(parts) >= 3 else stem
    period_token = period_token_raw.split("_")[0].strip()

    period_match = re.search(
        r"(Early|Mid|Late)([A-Za-z]{3,9})(\d{4})", period_token, flags=re.IGNORECASE
    )
    if period_match:
        part = period_match.group(1).capitalize()
        month_text = period_match.group(2)[:3].lower()
        year = int(period_match.group(3))
        month = MONTH_MAP.get(month_text, 1)
        model_period = f"{part}{month_text.capitalize()}_{year}"
        day = PART_DAY_MAP[part.lower()]
        model_date = date(year, month, day).isoformat()
    else:
        model_period = "unknown_period"
        model_date = ""

    model = f"{ticker}_{model_period}"
    return {
        "ticker": ticker,
        "model_period": model_period,
        "model_date": model_date,
        "model": model,
    }


def build_output_path(input_folder: Path, out_dir: Path) -> Path:
    out_dir.mkdir(parents=True, exist_ok=True)
    base_name = f"{input_folder.name}_PARAM"
    candidate = out_dir / f"{base_name}.xlsx"
    if not candidate.exists():
        return candidate

    idx = 1
    while True:
        candidate = out_dir / f"{base_name}.{idx}.xlsx"
        if not candidate.exists():
            return candidate
        idx += 1


def safe_close_source_workbook(wb: Any) -> None:
    close_attempts = (
        lambda: wb.close(save=False),
        lambda: wb.close(False),
        lambda: wb.close(),
    )
    for close_fn in close_attempts:
        try:
            close_fn()
            return
        except Exception:
            continue


def safe_quit_excel_app(app: Any) -> None:
    try:
        app.quit()
    except Exception:
        pass


def find_anchor(
    matrix: List[List[Any]],
    start_row: int,
    start_col: int,
    text: str = "max",
) -> Optional[Tuple[int, int]]:
    target = normalize_text(text)
    for r_idx, row_values in enumerate(matrix):
        for c_idx, cell_value in enumerate(row_values):
            cell_text = normalize_text(cell_value)
            if cell_text == target:
                return start_row + r_idx, start_col + c_idx
    return None


def header_map_for_row(
    row_values: List[Any],
    start_col: int,
    pattern_map: Dict[str, Sequence[str]],
) -> Dict[str, int]:
    mapping: Dict[str, int] = {}
    for c_idx, raw_value in enumerate(row_values):
        text = normalize_text(raw_value)
        if not text:
            continue
        for key, variants in pattern_map.items():
            if key in mapping:
                continue
            if any(variant in text for variant in variants):
                mapping[key] = start_col + c_idx
    return mapping


def choose_header_row(
    matrix: List[List[Any]],
    matrix_start_row: int,
    anchor_row: int,
    matrix_start_col: int,
    pattern_map: Dict[str, Sequence[str]],
) -> Tuple[int, Dict[str, int]]:
    if not matrix:
        return anchor_row, {}

    min_idx = max(0, anchor_row - matrix_start_row - 3)
    max_idx = min(len(matrix) - 1, anchor_row - matrix_start_row + 3)

    best_row = anchor_row
    best_map: Dict[str, int] = {}
    best_score = -1

    for row_idx in range(min_idx, max_idx + 1):
        current_map = header_map_for_row(matrix[row_idx], matrix_start_col, pattern_map)
        score = len(current_map)
        if score > best_score:
            best_score = score
            best_row = matrix_start_row + row_idx
            best_map = current_map

    if best_score <= 0:
        anchor_idx = anchor_row - matrix_start_row
        if 0 <= anchor_idx < len(matrix):
            best_map = header_map_for_row(
                matrix[anchor_idx], matrix_start_col, pattern_map
            )
        best_row = anchor_row

    return best_row, best_map


def sheet_used_range_matrix(ws: Any) -> Tuple[List[List[Any]], int, int, int]:
    used = ws.used_range
    start_row = used.row
    start_col = used.column
    matrix = normalize_matrix(used.options(ndim=2).value)
    row_count = len(matrix)
    end_row = start_row + row_count - 1 if row_count else start_row
    return matrix, start_row, start_col, end_row


def has_data_signal(values: Iterable[Any]) -> bool:
    return any(not is_blank(value) for value in values)


def first_non_blank(*values: Any) -> Any:
    for value in values:
        if not is_blank(value):
            return value
    return ""


def compact_signature(values: Sequence[Any]) -> Tuple[Any, ...]:
    compacted: List[Any] = []
    for value in values:
        numeric = to_float(value)
        if numeric is None:
            compacted.append(normalize_text(value))
        else:
            compacted.append(round(numeric, 10))
    return tuple(compacted)


def get_sheet_case_insensitive(wb: Any, target_name: str) -> Optional[Any]:
    lower_target = target_name.strip().lower()
    for sheet in wb.sheets:
        if sheet.name.strip().lower() == lower_target:
            return sheet
    return None


def process_empirical_sheet(
    ws: Any, metadata: Dict[str, str], source_file: str
) -> List[Dict[str, Any]]:
    matrix, start_row, start_col, end_row = sheet_used_range_matrix(ws)
    anchor = find_anchor(matrix, start_row, start_col, text="max")
    if anchor is None:
        print(f"  - Empirical Model: skipped (could not find 'max' anchor)")
        return []

    anchor_row, anchor_col = anchor
    header_row, header_map = choose_header_row(
        matrix, start_row, anchor_row, start_col, EMPIRICAL_HEADER_PATTERNS
    )
    data_start_row = header_row + 1
    n_quarters = 10

    max_col = anchor_col
    min_col = header_map.get("forecast_min", anchor_col + 1)
    num_q_col = header_map.get("num_quarters_used", anchor_col - 10)
    last_q_col = header_map.get("last_quarter_used", anchor_col - 9)
    forecast_col = header_map.get("forecast_value", anchor_col - 1)
    quarterly_col = header_map.get("quarterly_sales", anchor_col - 4)
    reported_col = header_map.get("reported_sales", anchor_col - 3)
    growth_col = header_map.get("growth_rate_pct", anchor_col - 2)
    captured_col = header_map.get("sales_captured_in_db_pct", anchor_col - 1)
    penetration_col = header_map.get("avg_penetration_pct")

    helper_avg_col = anchor_col + 20
    helper_fcst_col = anchor_col + 21
    rows_to_read: List[Tuple[int, int]] = []

    for i in range(n_quarters):
        row = data_start_row + i
        if row > end_row:
            break

        num_q_raw = matrix_cell(matrix, start_row, start_col, row, num_q_col)
        num_q = int(to_float(num_q_raw) or (i + 1))
        if num_q < 1:
            num_q = i + 1
        rolling_start = max(data_start_row, row - num_q + 1)

        avg_cell = ws.cells(row, helper_avg_col)
        fcst_cell = ws.cells(row, helper_fcst_col)

        if penetration_col is not None:
            avg_cell.formula2 = (
                f'=IFERROR(AVERAGE(R{rolling_start}C{penetration_col}:'
                f'R{row}C{penetration_col}),"")'
            )
        elif quarterly_col is not None and reported_col is not None:
            avg_cell.formula2 = (
                f'=IFERROR(SUM(R{rolling_start}C{quarterly_col}:R{row}C{quarterly_col})/'
                f'SUM(R{rolling_start}C{reported_col}:R{row}C{reported_col}),"")'
            )
        else:
            avg_cell.formula2 = '=""'

        if quarterly_col is not None:
            fcst_cell.formula2 = f'=IFERROR(RC{quarterly_col}/RC{helper_avg_col},"")'
        else:
            fcst_cell.formula2 = '=""'

        rows_to_read.append((row, num_q))

    if rows_to_read:
        ws.book.app.calculate()

    extracted_rows: List[Dict[str, Any]] = []
    for row, num_q in rows_to_read:
        forecast_max = matrix_cell(matrix, start_row, start_col, row, max_col)
        forecast_min = matrix_cell(matrix, start_row, start_col, row, min_col)
        quarterly_sales = matrix_cell(matrix, start_row, start_col, row, quarterly_col)
        reported_sales = matrix_cell(matrix, start_row, start_col, row, reported_col)
        raw_forecast = matrix_cell(matrix, start_row, start_col, row, forecast_col)

        avg_penetration = ws.cells(row, helper_avg_col).value
        calc_forecast = ws.cells(row, helper_fcst_col).value
        forecast_value = first_non_blank(raw_forecast, calc_forecast)
        actual_value = first_non_blank(reported_sales, "")
        last_quarter_used = matrix_cell(matrix, start_row, start_col, row, last_q_col)
        growth_rate = matrix_cell(matrix, start_row, start_col, row, growth_col)
        captured_pct = matrix_cell(matrix, start_row, start_col, row, captured_col)

        if not has_data_signal(
            [
                num_q,
                forecast_value,
                forecast_max,
                forecast_min,
                avg_penetration,
                quarterly_sales,
                reported_sales,
            ]
        ):
            continue

        extracted_rows.append(
            {
                "model": metadata["model"],
                "ticker": metadata["ticker"],
                "model_period": metadata["model_period"],
                "model_date": metadata["model_date"],
                "method": "empirical",
                "parameter_name": "avg_penetration_pct",
                "parameter_value": avg_penetration,
                "num_quarters_used": num_q,
                "last_quarter_used": last_quarter_used,
                "forecast_value": forecast_value,
                "actual_value": actual_value,
                "forecast_max": forecast_max,
                "forecast_min": forecast_min,
                "range_width": safe_subtract(forecast_max, forecast_min),
                "avg_penetration_pct": avg_penetration,
                "quarterly_sales": quarterly_sales,
                "reported_sales": reported_sales,
                "growth_rate_pct": growth_rate,
                "sales_captured_in_db_pct": captured_pct,
                "source_file": source_file,
            }
        )

    return extracted_rows


def process_regression_sheet(
    ws: Any, metadata: Dict[str, str], source_file: str
) -> List[Dict[str, Any]]:
    matrix, start_row, start_col, end_row = sheet_used_range_matrix(ws)
    anchor = find_anchor(matrix, start_row, start_col, text="max")
    if anchor is None:
        print(f"  - Regression Model: skipped (could not find 'max' anchor)")
        return []

    anchor_row, anchor_col = anchor
    header_row, header_map = choose_header_row(
        matrix, start_row, anchor_row, start_col, REGRESSION_HEADER_PATTERNS
    )
    data_start_row = header_row + 1
    n_quarters = 10

    y_col = anchor_col - 7
    x_col = anchor_col - 11

    max_col = anchor_col
    min_col = header_map.get("forecast_min", anchor_col + 1)
    num_q_col = header_map.get("num_quarters_used", anchor_col - 12)
    forecast_col = header_map.get("forecast_value", anchor_col - 1)
    actual_col = header_map.get("actual_value")

    helper_intercept_col = anchor_col + 20
    helper_slope_col = anchor_col + 21

    rows_to_read: List[Tuple[int, int]] = []
    for i in range(n_quarters):
        row = data_start_row + i
        if row > end_row:
            break

        num_q_raw = matrix_cell(matrix, start_row, start_col, row, num_q_col)
        num_q = int(to_float(num_q_raw) or (i + 1))
        if num_q < 2:
            num_q = 2

        regression_end_row = data_start_row + num_q - 1
        if regression_end_row > end_row:
            regression_end_row = end_row

        intercept_cell = ws.cells(row, helper_intercept_col)
        slope_cell = ws.cells(row, helper_slope_col)

        intercept_cell.formula2 = (
            f'=IFERROR(INTERCEPT(R{data_start_row}C{y_col}:R{regression_end_row}C{y_col},'
            f'R{data_start_row}C{x_col}:R{regression_end_row}C{x_col}),"")'
        )
        slope_cell.formula2 = (
            f'=IFERROR(SLOPE(R{data_start_row}C{y_col}:R{regression_end_row}C{y_col},'
            f'R{data_start_row}C{x_col}:R{regression_end_row}C{x_col}),"")'
        )
        rows_to_read.append((row, num_q))

    if rows_to_read:
        ws.book.app.calculate()

    extracted_rows: List[Dict[str, Any]] = []
    previous_signature: Optional[Tuple[Any, ...]] = None

    for row, num_q in rows_to_read:
        intercept = ws.cells(row, helper_intercept_col).value
        slope = ws.cells(row, helper_slope_col).value
        forecast_value = matrix_cell(matrix, start_row, start_col, row, forecast_col)
        forecast_max = matrix_cell(matrix, start_row, start_col, row, max_col)
        forecast_min = matrix_cell(matrix, start_row, start_col, row, min_col)
        actual_value = (
            matrix_cell(matrix, start_row, start_col, row, actual_col)
            if actual_col is not None
            else ""
        )

        if not has_data_signal(
            [num_q, intercept, slope, forecast_value, forecast_max, forecast_min]
        ):
            continue

        signature = compact_signature(
            [num_q, intercept, slope, forecast_value, forecast_max, forecast_min]
        )
        if signature == previous_signature:
            continue
        previous_signature = signature

        extracted_rows.append(
            {
                "model": metadata["model"],
                "ticker": metadata["ticker"],
                "model_period": metadata["model_period"],
                "model_date": metadata["model_date"],
                "method": "regression",
                "parameter_name": "num_quarters_used",
                "parameter_value": num_q,
                "num_quarters_used": num_q,
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

    return extracted_rows


def process_workbook_once(
    app: Any, file_path: Path
) -> Tuple[List[Dict[str, Any]], List[Dict[str, Any]]]:
    metadata = parse_file_label(file_path.name)
    empirical_rows: List[Dict[str, Any]] = []
    regression_rows: List[Dict[str, Any]] = []

    wb = app.books.open(str(file_path), update_links=False)
    try:
        empirical_sheet = get_sheet_case_insensitive(wb, "Empirical Model")
        if empirical_sheet is None:
            print("  - Empirical Model: skipped (sheet not found)")
        else:
            empirical_rows = process_empirical_sheet(
                empirical_sheet, metadata, file_path.name
            )

        regression_sheet = get_sheet_case_insensitive(wb, "Regression Model")
        if regression_sheet is None:
            print("  - Regression Model: skipped (sheet not found)")
        else:
            regression_rows = process_regression_sheet(
                regression_sheet, metadata, file_path.name
            )
    finally:
        safe_close_source_workbook(wb)

    return empirical_rows, regression_rows


def apply_sheet_formatting(ws: Any, headers: Sequence[str]) -> None:
    if not headers:
        return
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(headers))}{ws.max_row}"

    bold_font = Font(bold=True)
    for idx in range(1, len(headers) + 1):
        ws.cell(row=1, column=idx).font = bold_font

    for idx, header in enumerate(headers, start=1):
        max_len = len(header)
        for row_idx in range(2, ws.max_row + 1):
            value = ws.cell(row=row_idx, column=idx).value
            if value is None:
                continue
            value_len = len(str(value))
            if value_len > max_len:
                max_len = value_len
        ws.column_dimensions[get_column_letter(idx)].width = min(max(max_len + 2, 12), 48)


def write_sheet(ws: Any, headers: Sequence[str], rows: List[Dict[str, Any]]) -> None:
    ws.append(list(headers))
    for row in rows:
        ws.append([row.get(header, "") for header in headers])
    apply_sheet_formatting(ws, headers)


def write_output_workbook(
    output_path: Path,
    empirical_rows: List[Dict[str, Any]],
    regression_rows: List[Dict[str, Any]],
) -> None:
    wb = Workbook()
    default_ws = wb.active
    wb.remove(default_ws)

    empirical_ws = wb.create_sheet("empirical_candidates")
    write_sheet(empirical_ws, EMPIRICAL_COLUMNS, empirical_rows)

    regression_ws = wb.create_sheet("regression_candidates")
    write_sheet(regression_ws, REGRESSION_COLUMNS, regression_rows)

    wb.save(output_path)


def should_skip_file(path: Path) -> Optional[str]:
    if not path.is_file():
        return "not a file"
    if path.name.startswith("~"):
        return "temporary file"
    if path.suffix.lower() != ".xlsx":
        return "not an .xlsx file"
    if re.search(r"_PARAM(?:\.\d+)?$", path.stem, flags=re.IGNORECASE):
        return "looks like generated output"
    return None


def main() -> int:
    if not input_dir.exists():
        print(f"Input directory does not exist: {input_dir.resolve()}")
        return 1

    output_path = build_output_path(input_dir, output_dir)

    candidates: List[Path] = []
    for path in sorted(input_dir.iterdir()):
        reason = should_skip_file(path)
        if reason:
            print(f"Skipped file: {path.name} ({reason})")
            continue
        candidates.append(path)

    empirical_rows: List[Dict[str, Any]] = []
    regression_rows: List[Dict[str, Any]] = []
    processed_files = 0

    app = xw.App(visible=False, add_book=False)
    app.display_alerts = False
    app.screen_updating = False

    try:
        for file_path in candidates:
            print(f"Processing file: {file_path.name}")
            try:
                emp_rows, reg_rows = process_workbook_once(app, file_path)
            except Exception as exc:
                print(f"Skipped file: {file_path.name} (processing error: {exc})")
                continue

            processed_files += 1
            empirical_rows.extend(emp_rows)
            regression_rows.extend(reg_rows)
    finally:
        safe_quit_excel_app(app)

    write_output_workbook(output_path, empirical_rows, regression_rows)

    print(f"Output path: {output_path.resolve()}")
    print(f"Number of files processed: {processed_files}")
    print(f"Number of empirical rows: {len(empirical_rows)}")
    print(f"Number of regression rows: {len(regression_rows)}")

    return 0


if __name__ == "__main__":
    sys.exit(main())
