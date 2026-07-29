#!/usr/bin/env python3
"""Extract empirical/regression candidate rows from Excel model workbooks."""

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

# =========================
# User-editable paths
# =========================
input_dir = "./input"
output_dir = "./output"


EMPIRICAL_SHEET_NAME = "Empirical Model"
REGRESSION_SHEET_NAME = "Regression Model"
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

EMPIRICAL_HEADER_ALIASES = {
    "num_quarters_used": [
        "num quarters used",
        "number of quarters used",
        "quarters used",
        "num qtrs",
    ],
    "last_quarter_used": ["last quarter used", "last quarter", "last qtr"],
    "forecast_value": [
        "estimated total sold",
        "tot fcst",
        "total forecast",
        "forecast value",
    ],
    "actual_value": ["reported sales", "actual value", "actual sales", "actual"],
    "forecast_min": ["min"],
    "avg_penetration_pct": ["avg penetration", "average penetration"],
    "quarterly_sales": ["quarterly sales", "quarter sales"],
    "reported_sales": ["reported sales", "reported"],
    "growth_rate_pct": ["growth rate", "growth pct", "growth %"],
    "sales_captured_in_db_pct": [
        "sales captured in db",
        "captured in db",
        "captured %",
    ],
}

REGRESSION_HEADER_ALIASES = {
    "num_quarters_used": [
        "num quarters used",
        "number of quarters used",
        "quarters used",
        "num qtrs",
    ],
    "forecast_value": [
        "tot fcst w/o sa",
        "tot fcst without sa",
        "total forecast without sa",
    ],
    "actual_value": ["actual value", "actual sales", "actual", "reported sales"],
    "forecast_min": ["min"],
}

MONTH_TO_NUMBER = {
    "Jan": 1,
    "Feb": 2,
    "Mar": 3,
    "Apr": 4,
    "May": 5,
    "Jun": 6,
    "Jul": 7,
    "Aug": 8,
    "Sep": 9,
    "Oct": 10,
    "Nov": 11,
    "Dec": 12,
}

PHASE_TO_DAY = {
    "Early": 5,
    "Mid": 15,
    "Late": 25,
}


@dataclass
class SheetContext:
    sheet: xw.Sheet
    first_row: int
    first_col: int
    last_row: int
    last_col: int
    anchor_row: int
    anchor_col: int


def normalize_text(value: Any) -> str:
    if value is None:
        return ""
    text = str(value).strip().lower()
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
    first = values[0]
    if isinstance(first, tuple):
        values = [list(row) for row in values]
        first = values[0]
    if isinstance(first, list):
        return values
    return [values]


def row_to_list(values: Any) -> List[Any]:
    if isinstance(values, tuple):
        return list(values)
    if isinstance(values, list):
        return values
    return [values]


def clean_value(value: Any) -> Any:
    if isinstance(value, float) and math.isnan(value):
        return None
    return value


def to_float(value: Any) -> Optional[float]:
    value = clean_value(value)
    if value is None:
        return None
    if isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        return float(value)
    if isinstance(value, str):
        text = value.strip().replace(",", "")
        if not text:
            return None
        if text.endswith("%"):
            text = text[:-1].strip()
            try:
                return float(text) / 100.0
            except ValueError:
                return None
        try:
            return float(text)
        except ValueError:
            return None
    return None


def to_int(value: Any) -> Optional[int]:
    float_value = to_float(value)
    if float_value is None:
        return None
    return int(round(float_value))


def maybe_col_index(
    alias_map: Dict[str, List[str]], normalized_header: str
) -> Dict[str, int]:
    results: Dict[str, int] = {}
    for key, aliases in alias_map.items():
        for alias in aliases:
            if normalize_text(alias) in normalized_header:
                results[key] = 1
                break
    return results


def resolve_header_columns(
    context: SheetContext, alias_map: Dict[str, List[str]]
) -> Tuple[int, Dict[str, int]]:
    candidate_rows = [
        row
        for row in {context.anchor_row - 1, context.anchor_row, context.anchor_row + 1}
        if context.first_row <= row <= context.last_row
    ]
    best_row = context.anchor_row
    best_score = -1

    for row in candidate_rows:
        values = row_to_list(
            context.sheet.range((row, context.first_col), (row, context.last_col)).value
        )
        score = 0
        for cell_value in values:
            header = normalize_text(cell_value)
            if not header:
                continue
            if header == "max" or header == "min":
                score += 2
            score += len(maybe_col_index(alias_map, header))
        if score > best_score:
            best_row = row
            best_score = score

    header_values = row_to_list(
        context.sheet.range((best_row, context.first_col), (best_row, context.last_col)).value
    )
    columns: Dict[str, int] = {"forecast_max": context.anchor_col}

    for idx, cell_value in enumerate(header_values):
        header = normalize_text(cell_value)
        if not header:
            continue
        col = context.first_col + idx
        if header == "min":
            columns.setdefault("forecast_min", col)
        for key, aliases in alias_map.items():
            for alias in aliases:
                if normalize_text(alias) in header:
                    columns.setdefault(key, col)
                    break

    if "forecast_min" not in columns and context.anchor_col + 1 <= context.last_col:
        columns["forecast_min"] = context.anchor_col + 1
    return best_row, columns


def safe_close_workbook(book: xw.Book) -> None:
    try:
        book.close(save=False)
        return
    except TypeError:
        pass
    except Exception:
        pass

    try:
        book.close(False)
        return
    except Exception:
        pass

    try:
        book.api.Close(SaveChanges=False)
    except Exception:
        try:
            book.close()
        except Exception:
            pass


def set_formula2_r1c1(cell: xw.Range, formula_r1c1: str) -> None:
    try:
        cell.formula2 = formula_r1c1
        return
    except Exception:
        pass

    try:
        cell.api.Formula2R1C1 = formula_r1c1
        return
    except Exception:
        pass

    cell.api.FormulaR1C1 = formula_r1c1


def read_cell(sheet: xw.Sheet, row: int, col: Optional[int]) -> Any:
    if col is None or col < 1:
        return None
    return sheet.cells(row, col).value


def find_max_anchor(sheet: xw.Sheet) -> Optional[SheetContext]:
    used = sheet.used_range
    first_row = used.row
    first_col = used.column
    last_row = first_row + used.rows.count - 1
    last_col = first_col + used.columns.count - 1
    values = to_2d(used.value)
    if not values:
        return None

    for r_offset, row_values in enumerate(values):
        for c_offset, value in enumerate(row_values):
            if normalize_text(value) == "max":
                return SheetContext(
                    sheet=sheet,
                    first_row=first_row,
                    first_col=first_col,
                    last_row=last_row,
                    last_col=last_col,
                    anchor_row=first_row + r_offset,
                    anchor_col=first_col + c_offset,
                )
    return None


def parse_file_metadata(file_name: str) -> Optional[Dict[str, str]]:
    stem = Path(file_name).stem
    pattern = re.compile(
        r".*-\s*([A-Za-z0-9]+)\s*-\s*(Early|Mid|Late)"
        r"(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)(\d{4})_Send$",
        re.IGNORECASE,
    )
    match = pattern.match(stem)
    if not match:
        return None

    ticker = match.group(1).upper()
    phase = match.group(2).title()
    month_abbrev = match.group(3).title()
    year = int(match.group(4))
    month = MONTH_TO_NUMBER[month_abbrev]
    day = PHASE_TO_DAY[phase]
    model_period = f"{phase}{month_abbrev}_{year}"
    model_date = date(year, month, day).isoformat()

    return {
        "model": f"{ticker}_{model_period}",
        "ticker": ticker,
        "model_period": model_period,
        "model_date": model_date,
    }


def next_output_path(source_input_dir: Path, destination_output_dir: Path) -> Path:
    base_name = f"{source_input_dir.name}_PARAM"
    default_path = destination_output_dir / f"{base_name}.xlsx"
    if not default_path.exists():
        return default_path

    suffix = 1
    while True:
        candidate = destination_output_dir / f"{base_name}.{suffix}.xlsx"
        if not candidate.exists():
            return candidate
        suffix += 1


def derive_avg_penetration_formula(
    sheet: xw.Sheet,
    row: int,
    context: SheetContext,
    temp_col: int,
    requested_quarters: int,
) -> None:
    left_end = context.anchor_col - 1
    if left_end < context.first_col:
        sheet.cells(row, temp_col).value = None
        return

    row_values = row_to_list(sheet.range((row, context.first_col), (row, left_end)).value)
    numeric_cols: List[int] = []
    for idx, cell_value in enumerate(row_values):
        if to_float(cell_value) is not None:
            numeric_cols.append(context.first_col + idx)

    if not numeric_cols:
        sheet.cells(row, temp_col).value = None
        return

    width = min(max(requested_quarters, 1), len(numeric_cols))
    use_cols = numeric_cols[-width:]
    start_col = use_cols[0]
    end_col = use_cols[-1]
    start_offset = start_col - temp_col
    end_offset = end_col - temp_col
    formula = f"=AVERAGE(RC[{start_offset}]:RC[{end_offset}])"
    set_formula2_r1c1(sheet.cells(row, temp_col), formula)


def process_empirical_sheet(
    workbook: xw.Book, file_meta: Dict[str, str], source_file: str
) -> List[Dict[str, Any]]:
    try:
        sheet = workbook.sheets[EMPIRICAL_SHEET_NAME]
    except Exception:
        print(f"  skipped empirical extraction: sheet '{EMPIRICAL_SHEET_NAME}' not found")
        return []

    context = find_max_anchor(sheet)
    if context is None:
        print("  skipped empirical extraction: 'max' anchor not found")
        return []

    _, columns = resolve_header_columns(context, EMPIRICAL_HEADER_ALIASES)
    temp_col = context.last_col + 2
    start_data_row = context.anchor_row + 1
    end_data_row = min(context.anchor_row + N_QUARTERS, context.last_row)
    row_range = list(range(start_data_row, end_data_row + 1))

    for row in row_range:
        num_quarters_used = to_int(read_cell(sheet, row, columns.get("num_quarters_used")))
        if num_quarters_used is None:
            num_quarters_used = row - context.anchor_row
        derive_avg_penetration_formula(
            sheet=sheet,
            row=row,
            context=context,
            temp_col=temp_col,
            requested_quarters=num_quarters_used,
        )

    if row_range:
        workbook.app.calculate()

    results: List[Dict[str, Any]] = []
    for row in row_range:
        num_quarters_used = to_int(read_cell(sheet, row, columns.get("num_quarters_used")))
        if num_quarters_used is None:
            num_quarters_used = row - context.anchor_row

        forecast_max = to_float(sheet.cells(row, context.anchor_col).value)
        forecast_min = to_float(
            read_cell(sheet, row, columns.get("forecast_min", context.anchor_col + 1))
        )
        forecast_value = (
            to_float(read_cell(sheet, row, columns.get("forecast_value")))
            if "forecast_value" in columns
            else None
        )
        reported_sales = (
            to_float(read_cell(sheet, row, columns.get("reported_sales")))
            if "reported_sales" in columns
            else None
        )
        actual_value = (
            to_float(read_cell(sheet, row, columns.get("actual_value")))
            if "actual_value" in columns
            else reported_sales
        )
        avg_penetration_pct = to_float(sheet.cells(row, temp_col).value)
        if avg_penetration_pct is None and "avg_penetration_pct" in columns:
            avg_penetration_pct = to_float(
                read_cell(sheet, row, columns.get("avg_penetration_pct"))
            )

        last_quarter_used = (
            clean_value(read_cell(sheet, row, columns.get("last_quarter_used")))
            if "last_quarter_used" in columns
            else None
        )
        quarterly_sales = (
            to_float(read_cell(sheet, row, columns.get("quarterly_sales")))
            if "quarterly_sales" in columns
            else None
        )
        growth_rate_pct = (
            to_float(read_cell(sheet, row, columns.get("growth_rate_pct")))
            if "growth_rate_pct" in columns
            else None
        )
        sales_captured_in_db_pct = (
            to_float(read_cell(sheet, row, columns.get("sales_captured_in_db_pct")))
            if "sales_captured_in_db_pct" in columns
            else None
        )

        if forecast_value is None and forecast_max is not None and forecast_min is not None:
            forecast_value = (forecast_max + forecast_min) / 2.0

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

        range_width = (
            forecast_max - forecast_min
            if forecast_max is not None and forecast_min is not None
            else None
        )

        results.append(
            {
                "model": file_meta["model"],
                "ticker": file_meta["ticker"],
                "model_period": file_meta["model_period"],
                "model_date": file_meta["model_date"],
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
                "sales_captured_in_db_pct": sales_captured_in_db_pct,
                "source_file": source_file,
            }
        )

    if row_range:
        sheet.range((row_range[0], temp_col), (row_range[-1], temp_col)).value = None

    return results


def build_xy_points(
    sheet: xw.Sheet,
    start_row: int,
    end_row: int,
    x_col: int,
    y_col: int,
) -> List[Tuple[int, float, float]]:
    x_values = row_to_list(sheet.range((start_row, x_col), (end_row, x_col)).value)
    y_values = row_to_list(sheet.range((start_row, y_col), (end_row, y_col)).value)

    points: List[Tuple[int, float, float]] = []
    for idx, (x_value, y_value) in enumerate(zip(x_values, y_values)):
        x_num = to_float(x_value)
        y_num = to_float(y_value)
        if x_num is None or y_num is None:
            continue
        points.append((start_row + idx, x_num, y_num))
    return points


def compact_signature(values: Iterable[Any]) -> Tuple[Any, ...]:
    compacted: List[Any] = []
    for value in values:
        if isinstance(value, float):
            compacted.append(round(value, 10))
        else:
            compacted.append(value)
    return tuple(compacted)


def process_regression_sheet(
    workbook: xw.Book, file_meta: Dict[str, str], source_file: str
) -> List[Dict[str, Any]]:
    try:
        sheet = workbook.sheets[REGRESSION_SHEET_NAME]
    except Exception:
        print(f"  skipped regression extraction: sheet '{REGRESSION_SHEET_NAME}' not found")
        return []

    context = find_max_anchor(sheet)
    if context is None:
        print("  skipped regression extraction: 'max' anchor not found")
        return []

    _, columns = resolve_header_columns(context, REGRESSION_HEADER_ALIASES)
    y_col = context.anchor_col - 7
    x_col = context.anchor_col - 11
    if x_col < context.first_col or y_col < context.first_col:
        print("  skipped regression extraction: x/y columns are outside the used range")
        return []

    points = build_xy_points(
        sheet=sheet,
        start_row=context.anchor_row + 1,
        end_row=context.last_row,
        x_col=x_col,
        y_col=y_col,
    )
    if len(points) < 2:
        print("  skipped regression extraction: not enough x/y points")
        return []

    iterations = min(N_QUARTERS, len(points))
    temp_intercept_col = context.last_col + 2
    temp_slope_col = context.last_col + 3
    formula_rows = []

    for idx in range(1, iterations + 1):
        subset = points[-idx:]
        start = subset[0][0]
        end = subset[-1][0]
        target_row = context.anchor_row + idx

        intercept_formula = (
            f"=INTERCEPT(R{start}C{y_col}:R{end}C{y_col},R{start}C{x_col}:R{end}C{x_col})"
        )
        slope_formula = (
            f"=SLOPE(R{start}C{y_col}:R{end}C{y_col},R{start}C{x_col}:R{end}C{x_col})"
        )
        set_formula2_r1c1(sheet.cells(target_row, temp_intercept_col), intercept_formula)
        set_formula2_r1c1(sheet.cells(target_row, temp_slope_col), slope_formula)
        formula_rows.append(target_row)

    if formula_rows:
        workbook.app.calculate()

    results: List[Dict[str, Any]] = []
    previous_signature: Optional[Tuple[Any, ...]] = None

    for idx in range(1, iterations + 1):
        row = context.anchor_row + idx
        num_quarters_used = to_int(read_cell(sheet, row, columns.get("num_quarters_used")))
        if num_quarters_used is None:
            num_quarters_used = idx

        intercept = to_float(sheet.cells(row, temp_intercept_col).value)
        slope = to_float(sheet.cells(row, temp_slope_col).value)
        forecast_value = (
            to_float(read_cell(sheet, row, columns.get("forecast_value")))
            if "forecast_value" in columns
            else None
        )
        actual_value = (
            to_float(read_cell(sheet, row, columns.get("actual_value")))
            if "actual_value" in columns
            else None
        )
        forecast_max = to_float(sheet.cells(row, context.anchor_col).value)
        forecast_min = to_float(
            read_cell(sheet, row, columns.get("forecast_min", context.anchor_col + 1))
        )
        range_width = (
            forecast_max - forecast_min
            if forecast_max is not None and forecast_min is not None
            else None
        )

        signature = compact_signature(
            [
                num_quarters_used,
                forecast_value,
                forecast_max,
                forecast_min,
                intercept,
                slope,
            ]
        )
        if previous_signature == signature:
            continue
        previous_signature = signature

        if all(
            value is None
            for value in (
                forecast_value,
                forecast_max,
                forecast_min,
                intercept,
                slope,
            )
        ):
            continue

        results.append(
            {
                "model": file_meta["model"],
                "ticker": file_meta["ticker"],
                "model_period": file_meta["model_period"],
                "model_date": file_meta["model_date"],
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
        sheet.range(
            (formula_rows[0], temp_intercept_col),
            (formula_rows[-1], temp_slope_col),
        ).value = None

    return results


def write_sheet(ws: Any, headers: Sequence[str], rows: Sequence[Dict[str, Any]]) -> None:
    ws.append(list(headers))
    for row in rows:
        ws.append([clean_value(row.get(header)) for header in headers])

    for cell in ws[1]:
        cell.font = Font(bold=True)

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions

    for col_idx, header in enumerate(headers, start=1):
        max_len = len(header)
        for row_idx in range(2, ws.max_row + 1):
            value = ws.cell(row=row_idx, column=col_idx).value
            if value is None:
                continue
            max_len = max(max_len, len(str(value)))
        ws.column_dimensions[get_column_letter(col_idx)].width = min(max(max_len + 2, 12), 48)


def save_output_workbook(
    output_path: Path,
    empirical_rows: Sequence[Dict[str, Any]],
    regression_rows: Sequence[Dict[str, Any]],
) -> None:
    workbook = Workbook()
    empirical_sheet = workbook.active
    empirical_sheet.title = "empirical_candidates"
    write_sheet(empirical_sheet, EMPIRICAL_COLUMNS, empirical_rows)

    regression_sheet = workbook.create_sheet("regression_candidates")
    write_sheet(regression_sheet, REGRESSION_COLUMNS, regression_rows)

    workbook.save(output_path)


def main() -> int:
    source_input_dir = Path(input_dir).expanduser().resolve()
    destination_output_dir = Path(output_dir).expanduser().resolve()
    destination_output_dir.mkdir(parents=True, exist_ok=True)

    if not source_input_dir.exists() or not source_input_dir.is_dir():
        print(f"Input directory does not exist: {source_input_dir}")
        return 1

    output_path = next_output_path(source_input_dir, destination_output_dir)
    empirical_rows: List[Dict[str, Any]] = []
    regression_rows: List[Dict[str, Any]] = []
    files_processed = 0

    app: Optional[xw.App] = None
    original_calc_mode: Optional[str] = None
    try:
        app = xw.App(visible=False, add_book=False)
        app.display_alerts = False
        app.screen_updating = False
        original_calc_mode = app.calculation
        app.calculation = "manual"

        for file_path in sorted(source_input_dir.iterdir()):
            if not file_path.is_file():
                continue

            if file_path.name.startswith("~"):
                print(f"skipped file: {file_path.name} (temporary file)")
                continue

            if file_path.suffix.lower() != ".xlsx":
                print(f"skipped file: {file_path.name} (not .xlsx)")
                continue

            file_meta = parse_file_metadata(file_path.name)
            if file_meta is None:
                print(f"skipped file: {file_path.name} (filename format did not match)")
                continue

            workbook: Optional[xw.Book] = None
            try:
                print(f"processing file: {file_path.name}")
                workbook = app.books.open(str(file_path), update_links=False)
                empirical_rows.extend(process_empirical_sheet(workbook, file_meta, file_path.name))
                regression_rows.extend(process_regression_sheet(workbook, file_meta, file_path.name))
                files_processed += 1
                print(f"processed file: {file_path.name}")
            except Exception as exc:
                print(f"skipped file: {file_path.name} (error: {exc})")
            finally:
                if workbook is not None:
                    safe_close_workbook(workbook)
    finally:
        if app is not None:
            if original_calc_mode is not None:
                try:
                    app.calculation = original_calc_mode
                except Exception:
                    pass
            app.quit()

    save_output_workbook(
        output_path=output_path,
        empirical_rows=empirical_rows,
        regression_rows=regression_rows,
    )

    print(f"output path: {output_path}")
    print(f"number of files processed: {files_processed}")
    print(f"number of empirical rows: {len(empirical_rows)}")
    print(f"number of regression rows: {len(regression_rows)}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
