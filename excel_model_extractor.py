#!/usr/bin/env python3
from __future__ import annotations

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

# ---- Configure these two paths before running ----
input_dir = "./input"
output_dir = "./output"

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

DAY_BY_PERIOD = {"early": 5, "mid": 15, "late": 25}


@dataclass
class FileMeta:
    model: str
    ticker: str
    model_period: str
    model_date: str
    source_file: str


def normalize_label(value: Any) -> str:
    if value is None:
        return ""
    text = str(value).strip().lower()
    text = text.replace("%", " pct ")
    text = re.sub(r"[^a-z0-9]+", " ", text)
    return re.sub(r"\s+", " ", text).strip()


def to_matrix(values: Any) -> List[List[Any]]:
    if isinstance(values, list):
        if not values:
            return []
        if isinstance(values[0], list):
            return values
        return [values]
    return [[values]]


def flatten_vertical(values: Any) -> List[Any]:
    matrix = to_matrix(values)
    if not matrix:
        return []
    return [row[0] if row else None for row in matrix]


def is_number(value: Any) -> bool:
    if isinstance(value, bool):
        return False
    if isinstance(value, (int, float)):
        if isinstance(value, float) and math.isnan(value):
            return False
        return True
    return False


def to_float(value: Any) -> Optional[float]:
    if value is None or value == "":
        return None
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        if isinstance(value, float) and math.isnan(value):
            return None
        return float(value)
    text = str(value).strip().replace(",", "")
    if not text:
        return None
    try:
        return float(text)
    except ValueError:
        return None


def normalize_output_value(value: Any) -> Any:
    if value is None:
        return ""
    if isinstance(value, float) and math.isnan(value):
        return ""
    if isinstance(value, datetime):
        return value.date().isoformat()
    if isinstance(value, date):
        return value.isoformat()
    return value


def safe_close_workbook(wb: xw.Book) -> None:
    try:
        wb.close(save=False)
    except TypeError:
        try:
            wb.close(False)
        except Exception:
            wb.api.Close(False)
    except Exception:
        try:
            wb.close(False)
        except Exception:
            wb.api.Close(False)


def parse_file_meta(file_path: Path) -> FileMeta:
    stem = file_path.stem
    parts = [part.strip() for part in stem.split(" - ")]
    ticker = parts[1] if len(parts) >= 2 else ""

    period_match = re.search(
        r"(Early|Mid|Late)(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)(\d{4})",
        stem,
        flags=re.IGNORECASE,
    )
    model_period = ""
    model_date = ""
    if period_match:
        period_word = period_match.group(1).title()
        month_word = period_match.group(2).title()
        year = int(period_match.group(3))
        model_period = f"{period_word}{month_word}_{year}"
        month_num = MONTH_MAP[month_word.lower()]
        day = DAY_BY_PERIOD[period_word.lower()]
        model_date = date(year, month_num, day).isoformat()

    if not ticker:
        fallback_ticker = re.search(r"\b[A-Z]{2,6}\b", stem)
        ticker = fallback_ticker.group(0) if fallback_ticker else ""

    if ticker and model_period:
        model = f"{ticker}_{model_period}"
    elif ticker:
        model = ticker
    else:
        model = stem

    return FileMeta(
        model=model,
        ticker=ticker,
        model_period=model_period,
        model_date=model_date,
        source_file=file_path.name,
    )


def build_label_cells(
    matrix: Sequence[Sequence[Any]], base_row: int, base_col: int
) -> List[Tuple[str, int, int]]:
    labels: List[Tuple[str, int, int]] = []
    for r_idx, row_values in enumerate(matrix):
        for c_idx, value in enumerate(row_values):
            if isinstance(value, str) and value.strip():
                labels.append(
                    (normalize_label(value), base_row + r_idx, base_col + c_idx)
                )
    return labels


def matrix_value(
    matrix: Sequence[Sequence[Any]], base_row: int, base_col: int, row: int, col: int
) -> Any:
    r_idx = row - base_row
    c_idx = col - base_col
    if r_idx < 0 or c_idx < 0:
        return None
    if r_idx >= len(matrix):
        return None
    current_row = matrix[r_idx]
    if c_idx >= len(current_row):
        return None
    return current_row[c_idx]


def find_anchor_max(label_cells: Sequence[Tuple[str, int, int]]) -> Optional[Tuple[int, int]]:
    exact = [(row, col) for text, row, col in label_cells if text == "max"]
    if exact:
        return exact[0]

    partial = [
        (text, row, col) for text, row, col in label_cells if re.search(r"\bmax\b", text)
    ]
    if not partial:
        return None
    partial.sort(key=lambda entry: len(entry[0]))
    _, row, col = partial[0]
    return (row, col)


def find_col_offset(
    label_cells: Sequence[Tuple[str, int, int]],
    anchor_row: int,
    anchor_col: int,
    synonyms: Iterable[str],
    default: int,
) -> int:
    best: Optional[Tuple[int, int]] = None
    synonym_list = [normalize_label(token) for token in synonyms]
    for text, row, col in label_cells:
        if not any(token and token in text for token in synonym_list):
            continue
        score = abs(row - anchor_row) * 10 + abs(col - anchor_col)
        if row == anchor_row:
            score -= 6
        if best is None or score < best[0]:
            best = (score, col - anchor_col)
    if best is None:
        return default
    return best[1]


def detect_data_direction(
    matrix: Sequence[Sequence[Any]],
    base_row: int,
    base_col: int,
    anchor_row: int,
    anchor_col: int,
) -> int:
    below_count = 0
    above_count = 0
    for step in range(1, N_QUARTERS + 1):
        below_val = matrix_value(matrix, base_row, base_col, anchor_row + step, anchor_col)
        above_val = matrix_value(matrix, base_row, base_col, anchor_row - step, anchor_col)
        if is_number(below_val):
            below_count += 1
        if is_number(above_val):
            above_count += 1
    return 1 if below_count >= above_count else -1


def candidate_rows(
    matrix: Sequence[Sequence[Any]],
    base_row: int,
    anchor_row: int,
    direction: int,
) -> List[Tuple[int, int]]:
    first_row = base_row
    last_row = base_row + len(matrix) - 1
    rows: List[Tuple[int, int]] = []
    for n_quarters in range(1, N_QUARTERS + 1):
        row = anchor_row + (direction * n_quarters)
        if row < first_row or row > last_row:
            break
        rows.append((n_quarters, row))
    return rows


def safe_range_width(max_value: Any, min_value: Any) -> Any:
    max_num = to_float(max_value)
    min_num = to_float(min_value)
    if max_num is None or min_num is None:
        return ""
    return max_num - min_num


def regression_row_is_duplicate(prev_row: Dict[str, Any], next_row: Dict[str, Any]) -> bool:
    keys = [
        "num_quarters_used",
        "forecast_value",
        "forecast_max",
        "forecast_min",
        "intercept",
        "slope",
    ]
    for key in keys:
        prev_num = to_float(prev_row.get(key))
        next_num = to_float(next_row.get(key))
        if prev_num is not None and next_num is not None:
            if abs(prev_num - next_num) > 1e-9:
                return False
            continue
        if normalize_output_value(prev_row.get(key)) != normalize_output_value(next_row.get(key)):
            return False
    return True


def extract_empirical_rows(sheet: xw.Sheet, meta: FileMeta) -> List[Dict[str, Any]]:
    used = sheet.used_range
    matrix = to_matrix(used.value)
    if not matrix:
        return []

    base_row = used.row
    base_col = used.column
    labels = build_label_cells(matrix, base_row, base_col)
    anchor = find_anchor_max(labels)
    if anchor is None:
        return []

    anchor_row, anchor_col = anchor
    direction = detect_data_direction(matrix, base_row, base_col, anchor_row, anchor_col)
    rows = candidate_rows(matrix, base_row, anchor_row, direction)
    if not rows:
        return []

    min_offset = find_col_offset(labels, anchor_row, anchor_col, ["min"], default=1)
    forecast_offset = find_col_offset(
        labels,
        anchor_row,
        anchor_col,
        [
            "estimated total sold",
            "est total sold",
            "forecast value",
            "tot fcst",
            "total sold",
        ],
        default=-1,
    )
    actual_offset = find_col_offset(
        labels, anchor_row, anchor_col, ["reported sales", "actual sales", "actual value"], default=-2
    )
    num_quarters_offset = find_col_offset(
        labels, anchor_row, anchor_col, ["num quarters used", "quarters used", "n quarters"], default=-3
    )
    last_quarter_offset = find_col_offset(
        labels, anchor_row, anchor_col, ["last quarter used", "last quarter"], default=-11
    )
    quarterly_sales_offset = find_col_offset(
        labels, anchor_row, anchor_col, ["quarterly sales"], default=-7
    )
    growth_rate_offset = find_col_offset(labels, anchor_row, anchor_col, ["growth rate"], default=-4)
    captured_pct_offset = find_col_offset(
        labels,
        anchor_row,
        anchor_col,
        ["sales captured in db", "captured in db", "penetration"],
        default=-5,
    )
    reported_sales_offset = find_col_offset(
        labels, anchor_row, anchor_col, ["reported sales"], default=actual_offset
    )

    penetration_col = anchor_col + captured_pct_offset
    scratch_col = used.last_cell.column + 3
    scratch_start_row = used.last_cell.row + 1
    formula_rows: List[List[str]] = []

    first_data_row = rows[0][1]
    for _, row in rows:
        if direction == 1:
            start_row = first_data_row
            end_row = row
        else:
            start_row = row
            end_row = first_data_row
        formula_rows.append(
            [f'=IFERROR(AVERAGE(R{start_row}C{penetration_col}:R{end_row}C{penetration_col}),"")']
        )

    formula_target = sheet.range(
        (scratch_start_row, scratch_col),
        (scratch_start_row + len(formula_rows) - 1, scratch_col),
    )
    formula_target.formula2 = formula_rows
    sheet.book.app.calculate()
    avg_penetration_values = flatten_vertical(formula_target.value)

    extracted: List[Dict[str, Any]] = []
    for idx, (n_quarters, row) in enumerate(rows):
        forecast_max = matrix_value(matrix, base_row, base_col, row, anchor_col)
        forecast_min = matrix_value(matrix, base_row, base_col, row, anchor_col + min_offset)
        forecast_value = matrix_value(matrix, base_row, base_col, row, anchor_col + forecast_offset)
        actual_value = matrix_value(matrix, base_row, base_col, row, anchor_col + actual_offset)

        num_quarters_used = matrix_value(
            matrix, base_row, base_col, row, anchor_col + num_quarters_offset
        )
        if num_quarters_used in (None, ""):
            num_quarters_used = n_quarters

        last_quarter_used = matrix_value(
            matrix, base_row, base_col, row, anchor_col + last_quarter_offset
        )
        quarterly_sales = matrix_value(
            matrix, base_row, base_col, row, anchor_col + quarterly_sales_offset
        )
        reported_sales = matrix_value(
            matrix, base_row, base_col, row, anchor_col + reported_sales_offset
        )
        growth_rate_pct = matrix_value(
            matrix, base_row, base_col, row, anchor_col + growth_rate_offset
        )
        sales_captured_pct = matrix_value(
            matrix, base_row, base_col, row, anchor_col + captured_pct_offset
        )

        avg_penetration_pct = avg_penetration_values[idx] if idx < len(avg_penetration_values) else ""
        if avg_penetration_pct in (None, ""):
            avg_penetration_pct = sales_captured_pct

        if all(
            value in (None, "")
            for value in [forecast_max, forecast_min, forecast_value, actual_value, quarterly_sales]
        ):
            continue

        out_row = {
            "model": meta.model,
            "ticker": meta.ticker,
            "model_period": meta.model_period,
            "model_date": meta.model_date,
            "method": "empirical",
            "parameter_name": "avg_penetration_pct",
            "parameter_value": normalize_output_value(avg_penetration_pct),
            "num_quarters_used": normalize_output_value(num_quarters_used),
            "last_quarter_used": normalize_output_value(last_quarter_used),
            "forecast_value": normalize_output_value(forecast_value),
            "actual_value": normalize_output_value(actual_value),
            "forecast_max": normalize_output_value(forecast_max),
            "forecast_min": normalize_output_value(forecast_min),
            "range_width": normalize_output_value(safe_range_width(forecast_max, forecast_min)),
            "avg_penetration_pct": normalize_output_value(avg_penetration_pct),
            "quarterly_sales": normalize_output_value(quarterly_sales),
            "reported_sales": normalize_output_value(reported_sales),
            "growth_rate_pct": normalize_output_value(growth_rate_pct),
            "sales_captured_in_db_pct": normalize_output_value(sales_captured_pct),
            "source_file": meta.source_file,
        }
        extracted.append(out_row)

    return extracted


def extract_regression_rows(sheet: xw.Sheet, meta: FileMeta) -> List[Dict[str, Any]]:
    used = sheet.used_range
    matrix = to_matrix(used.value)
    if not matrix:
        return []

    base_row = used.row
    base_col = used.column
    labels = build_label_cells(matrix, base_row, base_col)
    anchor = find_anchor_max(labels)
    if anchor is None:
        return []

    anchor_row, anchor_col = anchor
    direction = detect_data_direction(matrix, base_row, base_col, anchor_row, anchor_col)
    rows = candidate_rows(matrix, base_row, anchor_row, direction)
    if not rows:
        return []

    min_offset = find_col_offset(labels, anchor_row, anchor_col, ["min"], default=1)
    num_quarters_offset = find_col_offset(
        labels, anchor_row, anchor_col, ["num quarters used", "quarters used", "n quarters"], default=-3
    )
    forecast_without_sa_offset = find_col_offset(
        labels,
        anchor_row,
        anchor_col,
        ["tot fcst w o sa", "tot fcst wo sa", "tot fcst without sa", "forecast without sa"],
        default=-1,
    )
    actual_value_offset = find_col_offset(
        labels, anchor_row, anchor_col, ["actual sales", "actual value", "reported sales"], default=-2
    )

    y_col = anchor_col - 7
    x_col = anchor_col - 11
    first_data_row = rows[0][1]

    scratch_col = used.last_cell.column + 3
    scratch_start_row = used.last_cell.row + 1
    formulas: List[List[str]] = []
    for _, row in rows:
        if direction == 1:
            start_row = first_data_row
            end_row = row
        else:
            start_row = row
            end_row = first_data_row
        formulas.append(
            [
                f'=IFERROR(INTERCEPT(R{start_row}C{y_col}:R{end_row}C{y_col},R{start_row}C{x_col}:R{end_row}C{x_col}),"")',
                f'=IFERROR(SLOPE(R{start_row}C{y_col}:R{end_row}C{y_col},R{start_row}C{x_col}:R{end_row}C{x_col}),"")',
            ]
        )

    coeff_target = sheet.range(
        (scratch_start_row, scratch_col),
        (scratch_start_row + len(formulas) - 1, scratch_col + 1),
    )
    coeff_target.formula2 = formulas
    sheet.book.app.calculate()
    coeff_values = to_matrix(coeff_target.value)

    extracted: List[Dict[str, Any]] = []
    for idx, (n_quarters, row) in enumerate(rows):
        forecast_max = matrix_value(matrix, base_row, base_col, row, anchor_col)
        forecast_min = matrix_value(matrix, base_row, base_col, row, anchor_col + min_offset)
        forecast_value = matrix_value(
            matrix, base_row, base_col, row, anchor_col + forecast_without_sa_offset
        )
        actual_value = matrix_value(matrix, base_row, base_col, row, anchor_col + actual_value_offset)
        num_quarters_used = matrix_value(
            matrix, base_row, base_col, row, anchor_col + num_quarters_offset
        )
        if num_quarters_used in (None, ""):
            num_quarters_used = n_quarters

        intercept_value = ""
        slope_value = ""
        if idx < len(coeff_values):
            coeff_row = coeff_values[idx]
            intercept_value = coeff_row[0] if len(coeff_row) > 0 else ""
            slope_value = coeff_row[1] if len(coeff_row) > 1 else ""

        if all(value in (None, "") for value in [forecast_max, forecast_min, forecast_value]):
            continue

        out_row = {
            "model": meta.model,
            "ticker": meta.ticker,
            "model_period": meta.model_period,
            "model_date": meta.model_date,
            "method": "regression",
            "parameter_name": "num_quarters_used",
            "parameter_value": normalize_output_value(num_quarters_used),
            "num_quarters_used": normalize_output_value(num_quarters_used),
            "forecast_value": normalize_output_value(forecast_value),
            "actual_value": normalize_output_value(actual_value),
            "forecast_max": normalize_output_value(forecast_max),
            "forecast_min": normalize_output_value(forecast_min),
            "range_width": normalize_output_value(safe_range_width(forecast_max, forecast_min)),
            "intercept": normalize_output_value(intercept_value),
            "slope": normalize_output_value(slope_value),
            "source_file": meta.source_file,
        }

        if extracted and regression_row_is_duplicate(extracted[-1], out_row):
            continue
        extracted.append(out_row)

    return extracted


def next_output_file(output_path: Path, input_folder_name: str) -> Path:
    base_name = f"{input_folder_name}_PARAM"
    candidate = output_path / f"{base_name}.xlsx"
    index = 1
    while candidate.exists():
        candidate = output_path / f"{base_name}.{index}.xlsx"
        index += 1
    return candidate


def write_sheet(ws: Any, columns: Sequence[str], rows: Sequence[Dict[str, Any]]) -> None:
    ws.append(list(columns))
    for cell in ws[1]:
        cell.font = Font(bold=True)

    for row_data in rows:
        ws.append([row_data.get(column, "") for column in columns])

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions

    max_lengths = [len(column) for column in columns]
    for row_idx in range(2, ws.max_row + 1):
        for col_idx, _ in enumerate(columns, start=1):
            value = ws.cell(row=row_idx, column=col_idx).value
            text = "" if value is None else str(value)
            if len(text) > max_lengths[col_idx - 1]:
                max_lengths[col_idx - 1] = len(text)

    for col_idx, max_length in enumerate(max_lengths, start=1):
        ws.column_dimensions[get_column_letter(col_idx)].width = min(max_length + 2, 45)


def write_output_workbook(
    output_file: Path,
    empirical_rows: Sequence[Dict[str, Any]],
    regression_rows: Sequence[Dict[str, Any]],
) -> None:
    workbook = Workbook()
    default_sheet = workbook.active
    workbook.remove(default_sheet)

    empirical_ws = workbook.create_sheet("empirical_candidates")
    regression_ws = workbook.create_sheet("regression_candidates")

    write_sheet(empirical_ws, EMPIRICAL_COLUMNS, empirical_rows)
    write_sheet(regression_ws, REGRESSION_COLUMNS, regression_rows)
    workbook.save(output_file)


def process_workbooks(input_path: Path, output_path: Path) -> Tuple[Path, int, int, int]:
    output_path.mkdir(parents=True, exist_ok=True)
    output_file = next_output_file(output_path, input_path.name)

    all_files = sorted(input_path.iterdir(), key=lambda path: path.name.lower())
    xlsx_files: List[Path] = []
    for file_path in all_files:
        if not file_path.is_file():
            continue
        if file_path.name.startswith("~"):
            print(f"SKIPPED {file_path.name}: temporary file")
            continue
        if file_path.suffix.lower() != ".xlsx":
            print(f"SKIPPED {file_path.name}: not an .xlsx file")
            continue
        xlsx_files.append(file_path)

    empirical_rows: List[Dict[str, Any]] = []
    regression_rows: List[Dict[str, Any]] = []
    processed_files = 0

    app = xw.App(visible=False, add_book=False)
    app.display_alerts = False
    app.screen_updating = False
    app.calculation = "manual"

    try:
        for file_path in xlsx_files:
            wb: Optional[xw.Book] = None
            try:
                wb = app.books.open(str(file_path), update_links=False)
                sheet_map = {sheet.name: sheet for sheet in wb.sheets}
                meta = parse_file_meta(file_path)

                if "Empirical Model" in sheet_map:
                    empirical_rows.extend(extract_empirical_rows(sheet_map["Empirical Model"], meta))
                else:
                    print(f"SKIPPED {file_path.name}: missing sheet 'Empirical Model'")

                if "Regression Model" in sheet_map:
                    regression_rows.extend(
                        extract_regression_rows(sheet_map["Regression Model"], meta)
                    )
                else:
                    print(f"SKIPPED {file_path.name}: missing sheet 'Regression Model'")

                processed_files += 1
                print(f"PROCESSED {file_path.name}")
            except Exception as exc:
                print(f"SKIPPED {file_path.name}: {exc}")
            finally:
                if wb is not None:
                    safe_close_workbook(wb)
    finally:
        app.quit()

    write_output_workbook(output_file, empirical_rows, regression_rows)
    return output_file, processed_files, len(empirical_rows), len(regression_rows)


def main() -> None:
    input_path = Path(input_dir).expanduser().resolve()
    output_path = Path(output_dir).expanduser().resolve()

    if not input_path.exists() or not input_path.is_dir():
        raise SystemExit(f"Input directory not found: {input_path}")

    output_file, processed_count, empirical_count, regression_count = process_workbooks(
        input_path, output_path
    )

    print(f"OUTPUT {output_file}")
    print(f"FILES_PROCESSED {processed_count}")
    print(f"EMPIRICAL_ROWS {empirical_count}")
    print(f"REGRESSION_ROWS {regression_count}")


if __name__ == "__main__":
    main()
