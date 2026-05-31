#!/usr/bin/env python3
from __future__ import annotations

import calendar
import math
import re
from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Sequence, Tuple

import xlwings as xw
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

# Update these paths before running.
input_dir = Path("input")
output_dir = Path("output")

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

DAY_MAP = {"Early": 5, "Mid": 15, "Late": 25}
MONTH_MAP = {m.lower(): i for i, m in enumerate(calendar.month_abbr) if m}
MONTH_MAP.update({m.lower(): i for i, m in enumerate(calendar.month_name) if m})

PERIOD_RE = re.compile(r"^(Early|Mid|Late)([A-Za-z]+)(\d{4})$", re.IGNORECASE)
PARAM_OUTPUT_RE_TEMPLATE = r"^{folder}_PARAM(?:\.\d+)?\.xlsx$"


@dataclass(frozen=True)
class ModelMeta:
    model: str
    ticker: str
    model_period: str
    model_date: str


def normalize_text(value: Any) -> str:
    if value is None:
        return ""
    return str(value).strip().replace("\n", " ").lower()


def to_number(value: Any) -> Optional[float]:
    if value is None or value == "":
        return None
    if isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        if isinstance(value, float) and math.isnan(value):
            return None
        return float(value)
    if isinstance(value, str):
        text = value.strip().replace(",", "")
        if not text:
            return None
        pct = text.endswith("%")
        if pct:
            text = text[:-1].strip()
        try:
            num = float(text)
        except ValueError:
            return None
        return num / 100.0 if pct else num
    return None


def to_int(value: Any) -> Optional[int]:
    num = to_number(value)
    if num is None:
        return None
    return int(round(num))


def to_output_value(value: Any) -> Any:
    if isinstance(value, datetime):
        return value.date().isoformat()
    if isinstance(value, date):
        return value.isoformat()
    return value


def parse_model_meta(file_path: Path) -> Optional[ModelMeta]:
    parts = [part.strip() for part in file_path.stem.split(" - ")]
    if len(parts) < 3:
        return None

    ticker = parts[-2].upper()
    period_token = parts[-1]
    if period_token.endswith("_Send"):
        period_token = period_token[:-5]
    period_match = PERIOD_RE.match(period_token)
    if not period_match:
        return None

    period_label = period_match.group(1).title()
    month_token = period_match.group(2)
    year = int(period_match.group(3))

    month_key = month_token.lower()
    month_num = MONTH_MAP.get(month_key) or MONTH_MAP.get(month_key[:3])
    if not month_num:
        return None

    month_abbr = calendar.month_abbr[month_num]
    model_period = f"{period_label}{month_abbr}_{year}"
    model_date = date(year, month_num, DAY_MAP[period_label]).isoformat()
    model = f"{ticker}_{model_period}"
    return ModelMeta(model=model, ticker=ticker, model_period=model_period, model_date=model_date)


def normalize_2d(values: Any) -> List[List[Any]]:
    if values is None:
        return []
    if isinstance(values, tuple):
        values = list(values)
    if not isinstance(values, list):
        return [[values]]
    if not values:
        return []
    if isinstance(values[0], tuple):
        values = [list(row) for row in values]
    if values and not isinstance(values[0], list):
        return [list(values)]
    return values


def read_sheet_matrix(sheet: xw.Sheet) -> Tuple[List[List[Any]], int, int, int, int]:
    used = sheet.used_range
    matrix = normalize_2d(used.value)
    top_row = used.row
    left_col = used.column
    if not matrix:
        return matrix, top_row, left_col, top_row, left_col
    max_width = max(len(row) for row in matrix if isinstance(row, list))
    last_row = top_row + len(matrix) - 1
    last_col = left_col + max_width - 1
    return matrix, top_row, left_col, last_row, last_col


def matrix_value(matrix: List[List[Any]], top_row: int, left_col: int, row: int, col: int) -> Any:
    row_idx = row - top_row
    col_idx = col - left_col
    if row_idx < 0 or col_idx < 0:
        return None
    if row_idx >= len(matrix):
        return None
    matrix_row = matrix[row_idx]
    if col_idx >= len(matrix_row):
        return None
    return matrix_row[col_idx]


def build_text_cells(matrix: List[List[Any]], top_row: int, left_col: int) -> List[Tuple[int, int, str]]:
    cells: List[Tuple[int, int, str]] = []
    for row_offset, row in enumerate(matrix):
        for col_offset, value in enumerate(row):
            text = normalize_text(value)
            if text:
                cells.append((top_row + row_offset, left_col + col_offset, text))
    return cells


def find_anchor(matrix: List[List[Any]], top_row: int, left_col: int, anchor_text: str = "max") -> Optional[Tuple[int, int]]:
    target = normalize_text(anchor_text)
    for row_offset, row in enumerate(matrix):
        for col_offset, value in enumerate(row):
            if normalize_text(value) == target:
                return top_row + row_offset, left_col + col_offset
    return None


def find_col_by_keyword_groups(
    text_cells: Sequence[Tuple[int, int, str]],
    anchor_row: int,
    anchor_col: int,
    keyword_groups: Sequence[Sequence[str]],
    max_row_distance: int = 20,
) -> Optional[int]:
    for group in keyword_groups:
        best: Optional[Tuple[int, int, int]] = None
        for row, col, text in text_cells:
            if abs(row - anchor_row) > max_row_distance:
                continue
            if all(keyword in text for keyword in group):
                score = (abs(row - anchor_row), abs(col - anchor_col), col)
                if best is None or score < best:
                    best = score
        if best is not None:
            return best[2]
    return None


def find_exact_text_col(
    text_cells: Sequence[Tuple[int, int, str]],
    anchor_row: int,
    anchor_col: int,
    target_text: str,
    row_radius: int = 8,
    col_radius: int = 25,
) -> Optional[int]:
    target = normalize_text(target_text)
    best: Optional[Tuple[int, int, int]] = None
    for row, col, text in text_cells:
        if text != target:
            continue
        if abs(row - anchor_row) > row_radius or abs(col - anchor_col) > col_radius:
            continue
        score = (abs(row - anchor_row), abs(col - anchor_col), col)
        if best is None or score < best:
            best = score
    return best[2] if best else None


def resolve_col(
    text_cells: Sequence[Tuple[int, int, str]],
    anchor_row: int,
    anchor_col: int,
    keyword_groups: Sequence[Sequence[str]],
    default_col: int,
) -> int:
    found_col = find_col_by_keyword_groups(text_cells, anchor_row, anchor_col, keyword_groups)
    return found_col if found_col is not None else default_col


def find_first_data_row(
    matrix: List[List[Any]],
    top_row: int,
    left_col: int,
    start_row: int,
    last_row: int,
    candidate_cols: Iterable[int],
) -> int:
    for row in range(start_row, last_row + 1):
        for col in candidate_cols:
            if to_number(matrix_value(matrix, top_row, left_col, row, col)) is not None:
                return row
    return start_row


def set_formula2_r1c1(cell: xw.Range, formula_r1c1: str) -> None:
    try:
        cell.api.Formula2R1C1 = formula_r1c1
        return
    except Exception:
        pass
    try:
        cell.formula2 = formula_r1c1
        return
    except Exception:
        pass
    cell.api.FormulaR1C1 = formula_r1c1


def safe_close_source_workbook(wb: xw.Book) -> None:
    try:
        wb.close(save=False)
        return
    except TypeError:
        try:
            wb.close(False)
            return
        except Exception:
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


def get_sheet_case_insensitive(wb: xw.Book, sheet_name: str) -> Optional[xw.Sheet]:
    target = sheet_name.strip().lower()
    for sheet in wb.sheets:
        if sheet.name.strip().lower() == target:
            return sheet
    return None


def read_row_cells(sheet: xw.Sheet, row: int, cols: Sequence[int]) -> Dict[int, Any]:
    values: Dict[int, Any] = {}
    for col in sorted(set(cols)):
        values[col] = sheet.range((row, col)).value
    return values


def extract_empirical_rows(wb: xw.Book, meta: ModelMeta, source_file: str) -> List[Dict[str, Any]]:
    sheet = get_sheet_case_insensitive(wb, "Empirical Model")
    if sheet is None:
        print(f"Skipped empirical extraction in {source_file}: sheet 'Empirical Model' not found")
        return []

    matrix, top_row, left_col, last_row, last_col = read_sheet_matrix(sheet)
    if not matrix:
        print(f"Skipped empirical extraction in {source_file}: sheet is empty")
        return []

    text_cells = build_text_cells(matrix, top_row, left_col)
    anchor = find_anchor(matrix, top_row, left_col, "max")
    if anchor is None:
        print(f"Skipped empirical extraction in {source_file}: 'max' anchor not found")
        return []
    anchor_row, anchor_col = anchor

    max_col = anchor_col
    min_col = find_exact_text_col(text_cells, anchor_row, anchor_col, "min") or (anchor_col + 1)
    forecast_col = resolve_col(
        text_cells,
        anchor_row,
        anchor_col,
        [
            ("estimated", "total", "sold"),
            ("forecast", "total"),
            ("forecast",),
        ],
        default_col=anchor_col - 1,
    )
    actual_col = resolve_col(
        text_cells,
        anchor_row,
        anchor_col,
        [
            ("reported", "sales"),
            ("actual", "sales"),
            ("actual",),
        ],
        default_col=anchor_col - 2,
    )
    num_quarters_col = resolve_col(
        text_cells,
        anchor_row,
        anchor_col,
        [
            ("num", "quarters", "used"),
            ("quarters", "used"),
            ("num", "quarter"),
        ],
        default_col=anchor_col - 12,
    )
    last_quarter_col = resolve_col(
        text_cells,
        anchor_row,
        anchor_col,
        [
            ("last", "quarter", "used"),
            ("last", "quarter"),
        ],
        default_col=num_quarters_col + 1,
    )
    avg_penetration_col = resolve_col(
        text_cells,
        anchor_row,
        anchor_col,
        [
            ("avg", "penetration"),
            ("average", "penetration"),
            ("penetration",),
        ],
        default_col=anchor_col - 6,
    )
    quarterly_sales_col = resolve_col(
        text_cells,
        anchor_row,
        anchor_col,
        [
            ("quarterly", "sales"),
            ("quarter", "sales"),
            ("sales",),
        ],
        default_col=anchor_col - 9,
    )
    reported_sales_col = resolve_col(
        text_cells,
        anchor_row,
        anchor_col,
        [
            ("reported", "sales"),
            ("actual", "sales"),
        ],
        default_col=anchor_col - 8,
    )
    growth_col = resolve_col(
        text_cells,
        anchor_row,
        anchor_col,
        [
            ("growth", "rate"),
            ("growth",),
        ],
        default_col=anchor_col - 4,
    )
    captured_col = resolve_col(
        text_cells,
        anchor_row,
        anchor_col,
        [
            ("captured", "db"),
            ("captured", "database"),
            ("captured",),
        ],
        default_col=anchor_col - 3,
    )

    first_data_row = find_first_data_row(
        matrix=matrix,
        top_row=top_row,
        left_col=left_col,
        start_row=anchor_row + 1,
        last_row=last_row,
        candidate_cols=(forecast_col, max_col, min_col, quarterly_sales_col, reported_sales_col),
    )
    row_numbers = [first_data_row + i for i in range(N_QUARTERS)]

    scratch_col = last_col + 20
    formula_rows: List[int] = []
    for idx, row in enumerate(row_numbers):
        num_q = to_int(matrix_value(matrix, top_row, left_col, row, num_quarters_col)) or (idx + 1)
        num_q = max(1, num_q)
        start_row = max(first_data_row, row - num_q + 1)
        formula = (
            f'=IFERROR(AVERAGE(R{start_row}C{quarterly_sales_col}:R{row}C{quarterly_sales_col})/'
            f'AVERAGE(R{start_row}C{reported_sales_col}:R{row}C{reported_sales_col}),"")'
        )
        set_formula2_r1c1(sheet.range((row, scratch_col)), formula)
        formula_rows.append(row)

    if formula_rows:
        wb.app.calculate()

    extracted_rows: List[Dict[str, Any]] = []
    for idx, row in enumerate(row_numbers):
        row_data = read_row_cells(
            sheet,
            row,
            [
                num_quarters_col,
                last_quarter_col,
                forecast_col,
                actual_col,
                max_col,
                min_col,
                avg_penetration_col,
                quarterly_sales_col,
                reported_sales_col,
                growth_col,
                captured_col,
                scratch_col,
            ],
        )

        num_quarters_used = to_int(row_data.get(num_quarters_col)) or (idx + 1)
        last_quarter_used = to_output_value(row_data.get(last_quarter_col))
        forecast_value = to_number(row_data.get(forecast_col))
        actual_value = to_number(row_data.get(actual_col))
        forecast_max = to_number(row_data.get(max_col))
        forecast_min = to_number(row_data.get(min_col))
        range_width = (
            forecast_max - forecast_min if forecast_max is not None and forecast_min is not None else None
        )
        quarterly_sales = to_number(row_data.get(quarterly_sales_col))
        reported_sales = to_number(row_data.get(reported_sales_col))
        growth_rate_pct = to_number(row_data.get(growth_col))
        sales_captured_pct = to_number(row_data.get(captured_col))

        avg_penetration_pct = to_number(row_data.get(scratch_col))
        if avg_penetration_pct is None:
            avg_penetration_pct = to_number(row_data.get(avg_penetration_col))
        if (
            avg_penetration_pct is None
            and quarterly_sales is not None
            and reported_sales not in (None, 0)
        ):
            avg_penetration_pct = quarterly_sales / reported_sales
        if actual_value is None:
            actual_value = reported_sales

        if all(
            value is None
            for value in (
                forecast_value,
                actual_value,
                forecast_max,
                forecast_min,
                avg_penetration_pct,
            )
        ):
            continue

        extracted_rows.append(
            {
                "model": meta.model,
                "ticker": meta.ticker,
                "model_period": meta.model_period,
                "model_date": meta.model_date,
                "method": "empirical",
                "parameter_name": "avg_penetration_pct",
                "parameter_value": avg_penetration_pct,
                "num_quarters_used": num_quarters_used,
                "last_quarter_used": last_quarter_used,
                "forecast_value": forecast_value,
                "actual_value": actual_value,
                "forecast_max": forecast_max,
                "forecast_min": forecast_min,
                "range_width": range_width,
                "avg_penetration_pct": avg_penetration_pct,
                "quarterly_sales": quarterly_sales,
                "reported_sales": reported_sales,
                "growth_rate_pct": growth_rate_pct,
                "sales_captured_in_db_pct": sales_captured_pct,
                "source_file": source_file,
            }
        )

    if formula_rows:
        sheet.range((formula_rows[0], scratch_col), (formula_rows[-1], scratch_col)).value = None
    return extracted_rows


def extract_regression_rows(wb: xw.Book, meta: ModelMeta, source_file: str) -> List[Dict[str, Any]]:
    sheet = get_sheet_case_insensitive(wb, "Regression Model")
    if sheet is None:
        print(f"Skipped regression extraction in {source_file}: sheet 'Regression Model' not found")
        return []

    matrix, top_row, left_col, last_row, last_col = read_sheet_matrix(sheet)
    if not matrix:
        print(f"Skipped regression extraction in {source_file}: sheet is empty")
        return []

    text_cells = build_text_cells(matrix, top_row, left_col)
    anchor = find_anchor(matrix, top_row, left_col, "max")
    if anchor is None:
        print(f"Skipped regression extraction in {source_file}: 'max' anchor not found")
        return []
    anchor_row, anchor_col = anchor

    y_col = anchor_col - 7
    x_col = anchor_col - 11
    max_col = anchor_col
    min_col = find_exact_text_col(text_cells, anchor_row, anchor_col, "min") or (anchor_col + 1)
    num_quarters_col = resolve_col(
        text_cells,
        anchor_row,
        anchor_col,
        [
            ("num", "quarters", "used"),
            ("quarters", "used"),
            ("num", "quarter"),
        ],
        default_col=anchor_col - 12,
    )
    forecast_col = resolve_col(
        text_cells,
        anchor_row,
        anchor_col,
        [
            ("tot", "fcst", "w/o", "sa"),
            ("forecast", "without", "sa"),
            ("forecast", "total"),
            ("forecast",),
        ],
        default_col=anchor_col - 1,
    )
    actual_col = resolve_col(
        text_cells,
        anchor_row,
        anchor_col,
        [
            ("actual", "sales"),
            ("reported", "sales"),
            ("actual",),
        ],
        default_col=anchor_col - 2,
    )
    intercept_col = resolve_col(
        text_cells,
        anchor_row,
        anchor_col,
        [
            ("intercept",),
        ],
        default_col=anchor_col + 2,
    )
    slope_col = resolve_col(
        text_cells,
        anchor_row,
        anchor_col,
        [
            ("slope",),
        ],
        default_col=anchor_col + 3,
    )

    first_data_row = find_first_data_row(
        matrix=matrix,
        top_row=top_row,
        left_col=left_col,
        start_row=anchor_row + 1,
        last_row=last_row,
        candidate_cols=(forecast_col, max_col, min_col, num_quarters_col),
    )
    row_numbers = [first_data_row + i for i in range(N_QUARTERS)]

    intercept_calc_col = last_col + 20
    slope_calc_col = last_col + 21
    formula_rows: List[int] = []
    for idx, row in enumerate(row_numbers):
        num_q = to_int(matrix_value(matrix, top_row, left_col, row, num_quarters_col)) or (idx + 1)
        num_q = max(1, num_q)
        start_row = max(first_data_row, row - num_q + 1)

        intercept_formula = (
            f'=IFERROR(INTERCEPT(R{start_row}C{y_col}:R{row}C{y_col},'
            f'R{start_row}C{x_col}:R{row}C{x_col}),"")'
        )
        slope_formula = (
            f'=IFERROR(SLOPE(R{start_row}C{y_col}:R{row}C{y_col},'
            f'R{start_row}C{x_col}:R{row}C{x_col}),"")'
        )

        set_formula2_r1c1(sheet.range((row, intercept_calc_col)), intercept_formula)
        set_formula2_r1c1(sheet.range((row, slope_calc_col)), slope_formula)
        formula_rows.append(row)

    if formula_rows:
        wb.app.calculate()

    extracted_rows: List[Dict[str, Any]] = []
    prev_key: Optional[Tuple[Any, ...]] = None
    for idx, row in enumerate(row_numbers):
        row_data = read_row_cells(
            sheet,
            row,
            [
                num_quarters_col,
                forecast_col,
                actual_col,
                max_col,
                min_col,
                intercept_col,
                slope_col,
                intercept_calc_col,
                slope_calc_col,
            ],
        )

        num_quarters_used = to_int(row_data.get(num_quarters_col)) or (idx + 1)
        forecast_value = to_number(row_data.get(forecast_col))
        actual_value = to_number(row_data.get(actual_col))
        forecast_max = to_number(row_data.get(max_col))
        forecast_min = to_number(row_data.get(min_col))
        range_width = (
            forecast_max - forecast_min if forecast_max is not None and forecast_min is not None else None
        )

        intercept = to_number(row_data.get(intercept_calc_col))
        slope = to_number(row_data.get(slope_calc_col))
        if intercept is None:
            intercept = to_number(row_data.get(intercept_col))
        if slope is None:
            slope = to_number(row_data.get(slope_col))

        if all(
            value is None
            for value in (
                forecast_value,
                actual_value,
                forecast_max,
                forecast_min,
                intercept,
                slope,
            )
        ):
            continue

        dedupe_key = (
            num_quarters_used,
            forecast_value,
            actual_value,
            forecast_max,
            forecast_min,
            intercept,
            slope,
        )
        if prev_key is not None and dedupe_key == prev_key:
            continue
        prev_key = dedupe_key

        extracted_rows.append(
            {
                "model": meta.model,
                "ticker": meta.ticker,
                "model_period": meta.model_period,
                "model_date": meta.model_date,
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

    if formula_rows:
        sheet.range((formula_rows[0], intercept_calc_col), (formula_rows[-1], intercept_calc_col)).value = None
        sheet.range((formula_rows[0], slope_calc_col), (formula_rows[-1], slope_calc_col)).value = None
    return extracted_rows


def build_output_path(input_folder: Path, out_folder: Path) -> Path:
    out_folder.mkdir(parents=True, exist_ok=True)
    base = f"{input_folder.name}_PARAM"
    candidate = out_folder / f"{base}.xlsx"
    idx = 1
    while candidate.exists():
        candidate = out_folder / f"{base}.{idx}.xlsx"
        idx += 1
    return candidate


def collect_source_files(in_folder: Path) -> List[Path]:
    files: List[Path] = []
    generated_re = re.compile(PARAM_OUTPUT_RE_TEMPLATE.format(folder=re.escape(in_folder.name)), re.IGNORECASE)

    for file_path in sorted(in_folder.iterdir()):
        if not file_path.is_file():
            print(f"Skipped {file_path.name}: not a file")
            continue
        if file_path.name.startswith("~"):
            print(f"Skipped {file_path.name}: temp file")
            continue
        if file_path.suffix.lower() != ".xlsx":
            print(f"Skipped {file_path.name}: not an .xlsx file")
            continue
        if generated_re.match(file_path.name):
            print(f"Skipped {file_path.name}: prior generated output")
            continue
        files.append(file_path)
    return files


def write_sheet(wb: Workbook, name: str, columns: Sequence[str], rows: Sequence[Dict[str, Any]]) -> None:
    ws = wb.create_sheet(name)
    ws.append(list(columns))
    for row in rows:
        ws.append([row.get(col) for col in columns])

    for cell in ws[1]:
        cell.font = Font(bold=True)
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions

    for col_idx, col_name in enumerate(columns, start=1):
        max_len = len(col_name)
        for row in rows:
            value = row.get(col_name)
            if value is None:
                continue
            value_text = str(value)
            if len(value_text) > max_len:
                max_len = len(value_text)
        ws.column_dimensions[get_column_letter(col_idx)].width = min(max_len + 2, 40)


def write_output_workbook(
    output_path: Path,
    empirical_rows: Sequence[Dict[str, Any]],
    regression_rows: Sequence[Dict[str, Any]],
) -> None:
    wb = Workbook()
    wb.remove(wb.active)
    write_sheet(wb, "empirical_candidates", EMPIRICAL_COLUMNS, empirical_rows)
    write_sheet(wb, "regression_candidates", REGRESSION_COLUMNS, regression_rows)
    wb.save(output_path)


def main() -> None:
    if not input_dir.exists():
        raise FileNotFoundError(f"Input folder does not exist: {input_dir}")

    source_files = collect_source_files(input_dir)
    output_path = build_output_path(input_dir, output_dir)

    empirical_rows: List[Dict[str, Any]] = []
    regression_rows: List[Dict[str, Any]] = []
    processed_count = 0

    app: Optional[xw.App] = None
    try:
        app = xw.App(visible=False, add_book=False)
        app.display_alerts = False
        app.screen_updating = False

        try:
            app.api.Calculation = -4135  # xlCalculationManual
        except Exception:
            pass

        for file_path in source_files:
            meta = parse_model_meta(file_path)
            if meta is None:
                print(f"Skipped {file_path.name}: filename does not match expected model pattern")
                continue

            wb: Optional[xw.Book] = None
            try:
                wb = app.books.open(str(file_path), update_links=False)
                empirical_rows.extend(extract_empirical_rows(wb, meta, file_path.name))
                regression_rows.extend(extract_regression_rows(wb, meta, file_path.name))
                processed_count += 1
                print(f"Processed {file_path.name}")
            except Exception as exc:
                print(f"Skipped {file_path.name}: {exc}")
            finally:
                if wb is not None:
                    safe_close_source_workbook(wb)
    finally:
        if app is not None:
            try:
                app.quit()
            except Exception:
                try:
                    app.kill()
                except Exception:
                    pass

    write_output_workbook(output_path, empirical_rows, regression_rows)
    print(f"Output path: {output_path}")
    print(f"Number of files processed: {processed_count}")
    print(f"Number of empirical rows: {len(empirical_rows)}")
    print(f"Number of regression rows: {len(regression_rows)}")


if __name__ == "__main__":
    main()
