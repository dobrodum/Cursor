#!/usr/bin/env python3
"""
Extract empirical and regression candidate rows from all .xlsx files in input_dir.

The script opens each source workbook once, processes both model sheets while it is
open, and writes one consolidated output workbook with:
  - empirical_candidates
  - regression_candidates
"""

from __future__ import annotations

import math
import re
from dataclasses import dataclass
from datetime import date
from pathlib import Path
from typing import Any, Dict, List, Optional, Sequence, Tuple

import pandas as pd
import xlwings as xw
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

# ---------------------------------------------------------------------------
# User-configurable paths
# ---------------------------------------------------------------------------
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

DAY_BY_PERIOD = {
    "early": 5,
    "mid": 15,
    "late": 25,
}

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


@dataclass(frozen=True)
class FileLabels:
    model: str
    ticker: str
    model_period: str
    model_date: str


def normalize_header(value: Any) -> str:
    if value is None:
        return ""
    return re.sub(r"[^a-z0-9]+", "", str(value).strip().lower())


def clean_value(value: Any) -> Any:
    if isinstance(value, float) and math.isnan(value):
        return None
    if isinstance(value, str):
        trimmed = value.strip()
        return trimmed if trimmed else None
    return value


def to_float(value: Any) -> Optional[float]:
    value = clean_value(value)
    if value is None:
        return None
    if isinstance(value, (int, float)):
        return float(value)
    text = str(value).replace(",", "").strip()
    if not text:
        return None
    if text.endswith("%"):
        text = text[:-1].strip()
    try:
        return float(text)
    except ValueError:
        return None


def parse_file_labels(file_path: Path) -> FileLabels:
    stem = file_path.stem
    parts = [part.strip() for part in stem.split(" - ")]

    ticker = parts[1] if len(parts) >= 2 and parts[1] else "UNKNOWN"

    period_source = parts[2] if len(parts) >= 3 else stem
    period_match = re.search(
        r"(Early|Mid|Late)\s*([A-Za-z]{3})\s*(\d{4})",
        period_source,
        flags=re.IGNORECASE,
    )

    model_period = "UnknownPeriod"
    model_date = ""
    if period_match:
        period_word = period_match.group(1).title()
        month_abbr = period_match.group(2).title()
        year = int(period_match.group(3))

        month_num = MONTH_BY_ABBR.get(month_abbr.lower())
        day = DAY_BY_PERIOD.get(period_word.lower())
        if month_num and day:
            dt = date(year, month_num, day)
            model_date = dt.isoformat()
            model_period = f"{period_word}{month_abbr}_{year}"

    model = f"{ticker}_{model_period}" if model_period else ticker
    return FileLabels(
        model=model,
        ticker=ticker,
        model_period=model_period,
        model_date=model_date,
    )


def choose_output_path(input_path: Path, out_dir: Path) -> Path:
    out_dir.mkdir(parents=True, exist_ok=True)
    base = f"{input_path.name}_PARAM.xlsx"
    candidate = out_dir / base
    suffix_num = 1
    while candidate.exists():
        candidate = out_dir / f"{input_path.name}_PARAM.{suffix_num}.xlsx"
        suffix_num += 1
    return candidate


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
        # Last fallback; still avoid explicit save() calls.
        workbook.close()


def find_anchor_cell(sheet: xw.Sheet, anchor_text: str = "max") -> Optional[Tuple[int, int]]:
    used = sheet.used_range
    values = used.value
    if values is None:
        return None

    if not isinstance(values, list):
        values_2d = [[values]]
    elif values and isinstance(values[0], list):
        values_2d = values
    else:
        values_2d = [values]

    target = anchor_text.strip().lower()
    base_row = used.row
    base_col = used.column
    for row_idx, row_values in enumerate(values_2d):
        for col_idx, cell_value in enumerate(row_values):
            if isinstance(cell_value, str) and cell_value.strip().lower() == target:
                return (base_row + row_idx, base_col + col_idx)
    return None


def build_header_index(
    sheet: xw.Sheet,
    header_row: int,
    start_col: int,
    end_col: int,
) -> Dict[str, int]:
    start_col = max(1, start_col)
    end_col = max(start_col, end_col)
    row_values = sheet.range((header_row, start_col), (header_row, end_col)).value

    if not isinstance(row_values, list):
        values = [row_values]
    elif row_values and isinstance(row_values[0], list):
        values = row_values[0]
    else:
        values = row_values

    index: Dict[str, int] = {}
    for offset, cell_value in enumerate(values):
        key = normalize_header(cell_value)
        if key:
            index[key] = start_col + offset
    return index


def get_column(
    header_index: Dict[str, int],
    aliases: Sequence[str],
    anchor_col: int,
    fallback_offset: int,
) -> int:
    for alias in aliases:
        key = normalize_header(alias)
        if key in header_index:
            return header_index[key]
    return max(1, anchor_col + fallback_offset)


def get_column_if_present(
    header_index: Dict[str, int],
    aliases: Sequence[str],
) -> Optional[int]:
    for alias in aliases:
        key = normalize_header(alias)
        if key in header_index:
            return header_index[key]
    return None


def read_column(sheet: xw.Sheet, col: int, row_start: int, row_end: int) -> List[Any]:
    if col < 1 or row_end < row_start:
        return [None] * max(0, row_end - row_start + 1)

    values = sheet.range((row_start, col), (row_end, col)).value
    if row_start == row_end:
        return [values]
    if isinstance(values, list) and values and isinstance(values[0], list):
        return [row[0] for row in values]
    if isinstance(values, list):
        return values
    return [values]


def write_empirical_average_formulas(
    sheet: xw.Sheet,
    row_start: int,
    row_end: int,
    source_col: int,
    target_col: int,
) -> None:
    rel_col = source_col - target_col
    for offset, row in enumerate(range(row_start, row_end + 1)):
        quarters_used = offset + 1
        if quarters_used == 1:
            formula = f"=RC[{rel_col}]"
        else:
            formula = f"=AVERAGE(R[-{quarters_used - 1}]C[{rel_col}]:RC[{rel_col}])"
        sheet.range((row, target_col)).formula2 = formula


def write_regression_formulas(
    sheet: xw.Sheet,
    row_start: int,
    row_end: int,
    y_col: int,
    x_col: int,
    intercept_col: int,
    slope_col: int,
) -> None:
    rel_y_for_intercept = y_col - intercept_col
    rel_x_for_intercept = x_col - intercept_col
    rel_y_for_slope = y_col - slope_col
    rel_x_for_slope = x_col - slope_col

    for offset, row in enumerate(range(row_start, row_end + 1)):
        quarters_used = offset + 1
        intercept_cell = sheet.range((row, intercept_col))
        slope_cell = sheet.range((row, slope_col))

        if quarters_used < 2:
            intercept_cell.formula2 = '=""'
            slope_cell.formula2 = '=""'
            continue

        intercept_formula = (
            f"=INTERCEPT("
            f"R[-{quarters_used - 1}]C[{rel_y_for_intercept}]:RC[{rel_y_for_intercept}],"
            f"R[-{quarters_used - 1}]C[{rel_x_for_intercept}]:RC[{rel_x_for_intercept}]"
            f")"
        )
        slope_formula = (
            f"=SLOPE("
            f"R[-{quarters_used - 1}]C[{rel_y_for_slope}]:RC[{rel_y_for_slope}],"
            f"R[-{quarters_used - 1}]C[{rel_x_for_slope}]:RC[{rel_x_for_slope}]"
            f")"
        )
        intercept_cell.formula2 = intercept_formula
        slope_cell.formula2 = slope_formula


def extract_empirical_rows(
    workbook: xw.Book,
    labels: FileLabels,
    source_file: str,
) -> List[Dict[str, Any]]:
    rows: List[Dict[str, Any]] = []
    try:
        sheet = workbook.sheets["Empirical Model"]
    except Exception:
        print(f"  - skipped empirical extraction: missing sheet 'Empirical Model'")
        return rows

    anchor = find_anchor_cell(sheet, "max")
    if not anchor:
        print("  - skipped empirical extraction: could not find 'max' anchor")
        return rows

    anchor_row, anchor_col = anchor
    header_index = build_header_index(sheet, anchor_row, anchor_col - 24, anchor_col + 12)

    columns = {
        "num_quarters_used": get_column(
            header_index,
            ["num quarters used", "num_quarters_used", "n quarters", "quarters used"],
            anchor_col,
            -13,
        ),
        "last_quarter_used": get_column(
            header_index,
            ["last quarter used", "last_quarter_used", "last quarter"],
            anchor_col,
            -12,
        ),
        "forecast_value": get_column(
            header_index,
            [
                "estimated total sold",
                "forecast value",
                "forecast",
                "estimated sold",
                "total sold",
            ],
            anchor_col,
            -4,
        ),
        "actual_value": get_column(
            header_index,
            ["reported sales", "actual value", "actual"],
            anchor_col,
            -3,
        ),
        "forecast_max": anchor_col,
        "forecast_min": get_column(header_index, ["min", "forecast min"], anchor_col, 1),
        "avg_penetration_pct": get_column(
            header_index,
            [
                "avg penetration %",
                "avg penetration pct",
                "avg penetration",
                "average penetration",
                "avg_penetration_pct",
            ],
            anchor_col,
            -8,
        ),
        "quarterly_sales": get_column(
            header_index,
            ["quarterly sales", "quarter sales", "sales db quarter"],
            anchor_col,
            -11,
        ),
        "reported_sales": get_column(
            header_index,
            ["reported sales", "reported_sales"],
            anchor_col,
            -10,
        ),
        "growth_rate_pct": get_column(
            header_index,
            ["growth rate %", "growth rate pct", "growth_rate_pct", "growth rate"],
            anchor_col,
            -7,
        ),
        "sales_captured_in_db_pct": get_column(
            header_index,
            [
                "sales captured in db %",
                "sales captured in db pct",
                "sales_captured_in_db_pct",
                "captured in db %",
            ],
            anchor_col,
            -6,
        ),
    }

    row_start = anchor_row + 1
    row_end = row_start + N_QUARTERS - 1

    last_used_col = sheet.used_range.last_cell.column
    avg_formula_col = last_used_col + 2
    avg_source_col = columns["sales_captured_in_db_pct"]
    if avg_source_col < 1:
        avg_source_col = columns["avg_penetration_pct"]

    write_empirical_average_formulas(
        sheet=sheet,
        row_start=row_start,
        row_end=row_end,
        source_col=avg_source_col,
        target_col=avg_formula_col,
    )
    workbook.app.calculate()

    avg_pen_from_formula = read_column(sheet, avg_formula_col, row_start, row_end)
    col_data = {
        key: read_column(sheet, col, row_start, row_end)
        for key, col in columns.items()
    }

    for i in range(N_QUARTERS):
        num_quarters = clean_value(col_data["num_quarters_used"][i]) or (i + 1)
        last_quarter = clean_value(col_data["last_quarter_used"][i])
        forecast_value = clean_value(col_data["forecast_value"][i])
        actual_value = clean_value(col_data["actual_value"][i])
        forecast_max = clean_value(col_data["forecast_max"][i])
        forecast_min = clean_value(col_data["forecast_min"][i])
        quarterly_sales = clean_value(col_data["quarterly_sales"][i])
        reported_sales = clean_value(col_data["reported_sales"][i])
        growth_rate_pct = clean_value(col_data["growth_rate_pct"][i])
        sales_captured = clean_value(col_data["sales_captured_in_db_pct"][i])

        avg_penetration = clean_value(avg_pen_from_formula[i])
        if avg_penetration is None:
            avg_penetration = clean_value(col_data["avg_penetration_pct"][i])

        if all(
            value is None
            for value in (forecast_value, actual_value, forecast_max, forecast_min, quarterly_sales)
        ):
            continue

        max_num = to_float(forecast_max)
        min_num = to_float(forecast_min)
        range_width = (max_num - min_num) if (max_num is not None and min_num is not None) else None

        rows.append(
            {
                "model": labels.model,
                "ticker": labels.ticker,
                "model_period": labels.model_period,
                "model_date": labels.model_date,
                "method": "empirical",
                "parameter_name": "avg_penetration_pct",
                "parameter_value": avg_penetration,
                "num_quarters_used": num_quarters,
                "last_quarter_used": last_quarter,
                "forecast_value": forecast_value,
                "actual_value": actual_value,
                "forecast_max": forecast_max,
                "forecast_min": forecast_min,
                "range_width": range_width,
                "avg_penetration_pct": avg_penetration,
                "quarterly_sales": quarterly_sales,
                "reported_sales": reported_sales,
                "growth_rate_pct": growth_rate_pct,
                "sales_captured_in_db_pct": sales_captured,
                "source_file": source_file,
            }
        )

    return rows


def rounded_signature(values: Sequence[Any]) -> Tuple[Any, ...]:
    output: List[Any] = []
    for value in values:
        numeric = to_float(value)
        if numeric is not None:
            output.append(round(numeric, 10))
        else:
            output.append(clean_value(value))
    return tuple(output)


def extract_regression_rows(
    workbook: xw.Book,
    labels: FileLabels,
    source_file: str,
) -> List[Dict[str, Any]]:
    rows: List[Dict[str, Any]] = []
    try:
        sheet = workbook.sheets["Regression Model"]
    except Exception:
        print(f"  - skipped regression extraction: missing sheet 'Regression Model'")
        return rows

    anchor = find_anchor_cell(sheet, "max")
    if not anchor:
        print("  - skipped regression extraction: could not find 'max' anchor")
        return rows

    anchor_row, anchor_col = anchor
    y_col = max(1, anchor_col - 7)
    x_col = max(1, anchor_col - 11)

    header_index = build_header_index(sheet, anchor_row, anchor_col - 24, anchor_col + 12)
    num_quarters_col = get_column(
        header_index,
        ["num quarters used", "num_quarters_used", "n quarters", "quarters used"],
        anchor_col,
        -13,
    )
    forecast_col = get_column(
        header_index,
        ["tot fcst w/o sa", "tot fcst wo sa", "forecast total without sa", "forecast value"],
        anchor_col,
        -2,
    )
    max_col = anchor_col
    min_col = get_column(header_index, ["min", "forecast min"], anchor_col, 1)
    actual_col = get_column_if_present(
        header_index,
        ["actual value", "actual", "reported sales"],
    )

    row_start = anchor_row + 1
    row_end = row_start + N_QUARTERS - 1

    last_used_col = sheet.used_range.last_cell.column
    intercept_col = last_used_col + 2
    slope_col = last_used_col + 3

    write_regression_formulas(
        sheet=sheet,
        row_start=row_start,
        row_end=row_end,
        y_col=y_col,
        x_col=x_col,
        intercept_col=intercept_col,
        slope_col=slope_col,
    )
    workbook.app.calculate()

    num_quarters_values = read_column(sheet, num_quarters_col, row_start, row_end)
    forecast_values = read_column(sheet, forecast_col, row_start, row_end)
    max_values = read_column(sheet, max_col, row_start, row_end)
    min_values = read_column(sheet, min_col, row_start, row_end)
    if actual_col is None:
        actual_values = [None] * N_QUARTERS
    else:
        actual_values = read_column(sheet, actual_col, row_start, row_end)
    intercept_values = read_column(sheet, intercept_col, row_start, row_end)
    slope_values = read_column(sheet, slope_col, row_start, row_end)

    prev_sig: Optional[Tuple[Any, ...]] = None
    for i in range(N_QUARTERS):
        num_quarters = clean_value(num_quarters_values[i]) or (i + 1)
        forecast_value = clean_value(forecast_values[i])
        forecast_max = clean_value(max_values[i])
        forecast_min = clean_value(min_values[i])
        intercept = clean_value(intercept_values[i])
        slope = clean_value(slope_values[i])
        actual_value = clean_value(actual_values[i])

        if all(
            value is None
            for value in (forecast_value, forecast_max, forecast_min, intercept, slope)
        ):
            continue

        max_num = to_float(forecast_max)
        min_num = to_float(forecast_min)
        range_width = (max_num - min_num) if (max_num is not None and min_num is not None) else None

        current_sig = rounded_signature(
            [num_quarters, forecast_value, forecast_max, forecast_min, intercept, slope]
        )
        if prev_sig is not None and current_sig == prev_sig:
            continue
        prev_sig = current_sig

        rows.append(
            {
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
                "range_width": range_width,
                "intercept": intercept,
                "slope": slope,
                "source_file": source_file,
            }
        )

    return rows


def format_sheet(worksheet, columns: Sequence[str], rows: List[Dict[str, Any]]) -> None:
    for cell in worksheet[1]:
        cell.font = Font(bold=True)

    worksheet.freeze_panes = "A2"
    worksheet.auto_filter.ref = worksheet.dimensions

    for idx, column_name in enumerate(columns, start=1):
        lengths = [len(column_name)]
        for row in rows:
            value = row.get(column_name)
            if value is None:
                continue
            lengths.append(len(str(value)))
        width = min(60, max(12, max(lengths) + 2))
        worksheet.column_dimensions[get_column_letter(idx)].width = width


def write_output_workbook(
    output_path: Path,
    empirical_rows: List[Dict[str, Any]],
    regression_rows: List[Dict[str, Any]],
) -> None:
    empirical_df = pd.DataFrame(empirical_rows, columns=EMPIRICAL_COLUMNS)
    regression_df = pd.DataFrame(regression_rows, columns=REGRESSION_COLUMNS)

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        empirical_df.to_excel(writer, sheet_name="empirical_candidates", index=False)
        regression_df.to_excel(writer, sheet_name="regression_candidates", index=False)

        empirical_ws = writer.sheets["empirical_candidates"]
        regression_ws = writer.sheets["regression_candidates"]
        format_sheet(empirical_ws, EMPIRICAL_COLUMNS, empirical_rows)
        format_sheet(regression_ws, REGRESSION_COLUMNS, regression_rows)


def iter_candidate_files(input_path: Path) -> List[Path]:
    files: List[Path] = []
    for file_path in sorted(input_path.iterdir()):
        if not file_path.is_file():
            continue
        if file_path.name.startswith("~"):
            print(f"Skipped {file_path.name}: temporary file")
            continue
        if file_path.suffix.lower() != ".xlsx":
            print(f"Skipped {file_path.name}: not an .xlsx file")
            continue
        files.append(file_path)
    return files


def main() -> None:
    input_path = Path(input_dir)
    out_path = Path(output_dir)

    if not input_path.exists() or not input_path.is_dir():
        raise FileNotFoundError(f"input_dir does not exist or is not a directory: {input_path}")

    files = iter_candidate_files(input_path)
    output_path = choose_output_path(input_path, out_path)

    empirical_rows: List[Dict[str, Any]] = []
    regression_rows: List[Dict[str, Any]] = []
    processed_files = 0

    app = xw.App(visible=False, add_book=False)
    app.display_alerts = False
    app.screen_updating = False

    try:
        for file_path in files:
            print(f"Processing {file_path.name}")
            workbook: Optional[xw.Book] = None
            try:
                labels = parse_file_labels(file_path)
                workbook = app.books.open(str(file_path), update_links=False)
                empirical_rows.extend(
                    extract_empirical_rows(
                        workbook=workbook,
                        labels=labels,
                        source_file=file_path.name,
                    )
                )
                regression_rows.extend(
                    extract_regression_rows(
                        workbook=workbook,
                        labels=labels,
                        source_file=file_path.name,
                    )
                )
                processed_files += 1
            except Exception as exc:
                print(f"Skipped {file_path.name}: processing error ({exc})")
            finally:
                if workbook is not None:
                    safe_close_workbook(workbook)
    finally:
        try:
            app.quit()
        except Exception:
            pass

    write_output_workbook(
        output_path=output_path,
        empirical_rows=empirical_rows,
        regression_rows=regression_rows,
    )

    print(f"Output path: {output_path}")
    print(f"Number of files processed: {processed_files}")
    print(f"Number of empirical rows: {len(empirical_rows)}")
    print(f"Number of regression rows: {len(regression_rows)}")


if __name__ == "__main__":
    main()
