#!/usr/bin/env python3
from __future__ import annotations

import re
from datetime import date
from pathlib import Path
from typing import Any, Dict, List, Optional, Sequence, Tuple

import xlwings as xw
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

# --------- user inputs ---------
input_dir = Path("input")
output_dir = Path("output")
# -------------------------------

EMPIRICAL_SOURCE_SHEET = "Empirical Model"
REGRESSION_SOURCE_SHEET = "Regression Model"
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

EMPIRICAL_DEFAULT_OFFSETS = {
    "num_quarters_used": -9,
    "last_quarter_used": -8,
    "avg_penetration_pct": -7,
    "quarterly_sales": -6,
    "reported_sales": -5,
    "growth_rate_pct": -4,
    "sales_captured_in_db_pct": -3,
    "forecast_value": -2,
    "actual_value": -1,
    "forecast_max": 0,
    "forecast_min": 1,
}

REGRESSION_DEFAULT_OFFSETS = {
    "num_quarters_used": -4,
    "intercept": -3,
    "slope": -2,
    "forecast_value": -1,
    "forecast_max": 0,
    "forecast_min": 1,
    "actual_value": -5,
}

EMPIRICAL_HEADER_SYNONYMS = {
    "num_quarters_used": {
        "num quarters used",
        "quarters used",
        "num quarters",
        "number of quarters",
    },
    "last_quarter_used": {"last quarter used", "last quarter", "last qtr"},
    "avg_penetration_pct": {"avg penetration pct", "avg penetration", "average penetration"},
    "quarterly_sales": {"quarterly sales", "qtr sales"},
    "reported_sales": {"reported sales", "sales reported"},
    "growth_rate_pct": {"growth rate pct", "growth rate"},
    "sales_captured_in_db_pct": {
        "sales captured in db pct",
        "sales captured in db",
        "captured in db pct",
    },
    "forecast_value": {
        "estimated total sold",
        "est total sold",
        "forecast value",
        "forecast",
    },
    "actual_value": {"actual value", "actual sales", "reported sales"},
    "forecast_max": {"max", "forecast max"},
    "forecast_min": {"min", "forecast min"},
}

REGRESSION_HEADER_SYNONYMS = {
    "num_quarters_used": {
        "num quarters used",
        "quarters used",
        "num quarters",
        "number of quarters",
    },
    "intercept": {"intercept"},
    "slope": {"slope"},
    "forecast_value": {
        "tot fcst w o sa",
        "tot fcst wo sa",
        "tot fcst w sa",
        "tot fcst",
        "forecast total without sa",
    },
    "forecast_max": {"max", "forecast max"},
    "forecast_min": {"min", "forecast min"},
    "actual_value": {"actual value", "actual sales", "reported sales"},
}

MONTH_MAP = {
    "jan": 1,
    "january": 1,
    "feb": 2,
    "february": 2,
    "mar": 3,
    "march": 3,
    "apr": 4,
    "april": 4,
    "may": 5,
    "jun": 6,
    "june": 6,
    "jul": 7,
    "july": 7,
    "aug": 8,
    "august": 8,
    "sep": 9,
    "sept": 9,
    "september": 9,
    "oct": 10,
    "october": 10,
    "nov": 11,
    "november": 11,
    "dec": 12,
    "december": 12,
}

PERIOD_RE = re.compile(r"(Early|Mid|Late)\s*([A-Za-z]+)\s*(20\d{2})", re.IGNORECASE)
DAY_BY_PHASE = {"Early": 5, "Mid": 15, "Late": 25}


def normalize_label(value: Any) -> str:
    if value is None:
        return ""
    text = str(value).strip().lower()
    text = text.replace("%", " pct ")
    text = re.sub(r"[^a-z0-9]+", " ", text)
    return " ".join(text.split())


def as_2d(values: Any) -> List[List[Any]]:
    if isinstance(values, list):
        if values and isinstance(values[0], list):
            return values
        return [values]
    return [[values]]


def to_float(value: Any) -> Optional[float]:
    if value is None or value == "":
        return None
    if isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        return float(value)
    text = str(value).strip().replace(",", "")
    if not text:
        return None
    try:
        return float(text)
    except ValueError:
        return None


def to_int(value: Any) -> Optional[int]:
    parsed = to_float(value)
    if parsed is None:
        return None
    return int(round(parsed))


def pick(*values: Any) -> Any:
    for value in values:
        if value is not None and value != "":
            return value
    return None


def safe_subtract(left: Any, right: Any) -> Optional[float]:
    left_num = to_float(left)
    right_num = to_float(right)
    if left_num is None or right_num is None:
        return None
    return left_num - right_num


def find_sheet(workbook: xw.Book, sheet_name: str) -> Optional[xw.Sheet]:
    target = sheet_name.strip().lower()
    for sheet in workbook.sheets:
        if sheet.name.strip().lower() == target:
            return sheet
    return None


def find_anchor(sheet: xw.Sheet, anchor_text: str = "max") -> Optional[Tuple[int, int]]:
    used = sheet.used_range
    values = as_2d(used.value)
    base_row = used.row
    base_col = used.column
    for r_idx, row_values in enumerate(values):
        for c_idx, value in enumerate(row_values):
            if normalize_label(value) == anchor_text:
                return base_row + r_idx, base_col + c_idx
    return None


def derive_offsets_from_headers(
    sheet: xw.Sheet,
    anchor_row: int,
    anchor_col: int,
    defaults: Dict[str, int],
    synonyms: Dict[str, set],
    window: int = 24,
) -> Dict[str, int]:
    offsets = dict(defaults)
    start_col = max(1, anchor_col - window)
    end_col = anchor_col + window
    header_row_values = as_2d(sheet.range((anchor_row, start_col), (anchor_row, end_col)).value)[0]
    for idx, raw_header in enumerate(header_row_values):
        normalized = normalize_label(raw_header)
        if not normalized:
            continue
        current_col = start_col + idx
        for field_name, aliases in synonyms.items():
            if normalized in aliases:
                offsets[field_name] = current_col - anchor_col
    return offsets


def r1c1_ref(row_offset: int = 0, col_offset: int = 0) -> str:
    if row_offset == 0:
        row_part = "R"
    else:
        row_part = f"R[{row_offset}]"
    if col_offset == 0:
        col_part = "C"
    else:
        col_part = f"C[{col_offset}]"
    return row_part + col_part


def set_formula2_r1c1(cell: xw.Range, formula_r1c1: str) -> None:
    formula = formula_r1c1 if formula_r1c1.startswith("=") else f"={formula_r1c1}"
    try:
        cell.api.Formula2R1C1 = formula
    except Exception:
        # Fallback for environments where Formula2R1C1 is not exposed.
        cell.formula2 = formula


def get_block_values(
    sheet: xw.Sheet,
    start_row: int,
    end_row: int,
    start_col: int,
    end_col: int,
) -> List[List[Any]]:
    if end_row < start_row or end_col < start_col:
        return []
    return as_2d(sheet.range((start_row, start_col), (end_row, end_col)).value)


def block_value(
    row_values: Sequence[Any],
    start_col: int,
    anchor_col: int,
    offset: int,
) -> Any:
    absolute_col = anchor_col + offset
    index = absolute_col - start_col
    if 0 <= index < len(row_values):
        return row_values[index]
    return None


def parse_file_labels(file_name: str) -> Dict[str, str]:
    stem = Path(file_name).stem
    parts = [part.strip() for part in stem.split(" - ")]
    ticker = "UNKNOWN"
    if len(parts) >= 2 and parts[1]:
        ticker = parts[1].upper()

    period_text = parts[2] if len(parts) >= 3 else stem
    period_text = re.sub(r"[_\-\s]*send$", "", period_text, flags=re.IGNORECASE).strip()

    period_match = PERIOD_RE.search(period_text)
    if period_match:
        phase_raw, month_raw, year_raw = period_match.groups()
        phase = phase_raw.capitalize()
        month_key = month_raw.lower()
        month_num = MONTH_MAP.get(month_key)
        if month_num is None:
            month_num = MONTH_MAP.get(month_key[:3])
        if month_num is not None:
            month_abbrev = date(2000, month_num, 1).strftime("%b")
            year = int(year_raw)
            model_period = f"{phase}{month_abbrev}_{year}"
            model_date = date(year, month_num, DAY_BY_PHASE[phase]).isoformat()
        else:
            model_period = period_text.replace(" ", "")
            model_date = ""
    else:
        model_period = period_text.replace(" ", "")
        model_date = ""

    model = f"{ticker}_{model_period}" if model_period else ticker
    return {
        "model": model,
        "ticker": ticker,
        "model_period": model_period,
        "model_date": model_date,
    }


def choose_output_path(in_dir: Path, out_dir: Path) -> Path:
    folder_name = in_dir.resolve().name
    first_candidate = out_dir / f"{folder_name}_PARAM.xlsx"
    if not first_candidate.exists():
        return first_candidate

    index = 1
    while True:
        candidate = out_dir / f"{folder_name}_PARAM.{index}.xlsx"
        if not candidate.exists():
            return candidate
        index += 1


def close_source_workbook(workbook: xw.Book) -> None:
    close_attempts = (
        lambda: workbook.close(save=False),
        lambda: workbook.api.Close(False),
        lambda: workbook.api.Close(SaveChanges=False),
    )
    last_error: Optional[Exception] = None
    for close_fn in close_attempts:
        try:
            close_fn()
            return
        except Exception as exc:
            last_error = exc
    if last_error:
        raise last_error


def extract_empirical_rows(
    workbook: xw.Book,
    labels: Dict[str, str],
    source_file: str,
) -> List[Dict[str, Any]]:
    rows: List[Dict[str, Any]] = []
    sheet = find_sheet(workbook, EMPIRICAL_SOURCE_SHEET)
    if sheet is None:
        return rows

    anchor = find_anchor(sheet, "max")
    if anchor is None:
        return rows
    anchor_row, anchor_col = anchor
    offsets = derive_offsets_from_headers(
        sheet=sheet,
        anchor_row=anchor_row,
        anchor_col=anchor_col,
        defaults=EMPIRICAL_DEFAULT_OFFSETS,
        synonyms=EMPIRICAL_HEADER_SYNONYMS,
    )

    data_start_row = anchor_row + 1
    data_end_row = data_start_row + N_QUARTERS - 1

    avg_col = anchor_col + offsets["avg_penetration_pct"]
    capture_col = anchor_col + offsets["sales_captured_in_db_pct"]
    relative_capture = capture_col - avg_col

    formulas_written = False
    for idx in range(N_QUARTERS):
        row_num = data_start_row + idx
        quarter_count = idx + 1
        start_ref = r1c1_ref(row_offset=-(quarter_count - 1), col_offset=relative_capture)
        end_ref = r1c1_ref(row_offset=0, col_offset=relative_capture)
        formula = f'=IFERROR(AVERAGE({start_ref}:{end_ref}),"")'
        set_formula2_r1c1(sheet.cells(row_num, avg_col), formula)
        formulas_written = True

    if formulas_written:
        workbook.app.calculate()

    needed_offsets = [offsets[field] for field in EMPIRICAL_DEFAULT_OFFSETS.keys()]
    start_col = anchor_col + min(needed_offsets)
    end_col = anchor_col + max(needed_offsets)
    data_block = get_block_values(sheet, data_start_row, data_end_row, start_col, end_col)

    for idx, row_values in enumerate(data_block):
        num_quarters_used = pick(to_int(block_value(row_values, start_col, anchor_col, offsets["num_quarters_used"])), idx + 1)
        last_quarter_used = block_value(row_values, start_col, anchor_col, offsets["last_quarter_used"])
        avg_penetration_pct = block_value(row_values, start_col, anchor_col, offsets["avg_penetration_pct"])
        quarterly_sales = block_value(row_values, start_col, anchor_col, offsets["quarterly_sales"])
        reported_sales = block_value(row_values, start_col, anchor_col, offsets["reported_sales"])
        growth_rate_pct = block_value(row_values, start_col, anchor_col, offsets["growth_rate_pct"])
        sales_captured_in_db_pct = block_value(
            row_values,
            start_col,
            anchor_col,
            offsets["sales_captured_in_db_pct"],
        )
        forecast_value = block_value(row_values, start_col, anchor_col, offsets["forecast_value"])
        actual_value = pick(
            block_value(row_values, start_col, anchor_col, offsets["actual_value"]),
            reported_sales,
        )
        forecast_max = block_value(row_values, start_col, anchor_col, offsets["forecast_max"])
        forecast_min = block_value(row_values, start_col, anchor_col, offsets["forecast_min"])

        has_data = any(
            value not in (None, "")
            for value in (
                num_quarters_used,
                forecast_value,
                forecast_max,
                forecast_min,
                avg_penetration_pct,
            )
        )
        if not has_data:
            continue

        rows.append(
            {
                "model": labels["model"],
                "ticker": labels["ticker"],
                "model_period": labels["model_period"],
                "model_date": labels["model_date"],
                "method": "empirical",
                "parameter_name": "avg_penetration_pct",
                "parameter_value": avg_penetration_pct,
                "num_quarters_used": num_quarters_used,
                "last_quarter_used": last_quarter_used,
                "forecast_value": forecast_value,
                "actual_value": actual_value,
                "forecast_max": forecast_max,
                "forecast_min": forecast_min,
                "range_width": safe_subtract(forecast_max, forecast_min),
                "avg_penetration_pct": avg_penetration_pct,
                "quarterly_sales": quarterly_sales,
                "reported_sales": reported_sales,
                "growth_rate_pct": growth_rate_pct,
                "sales_captured_in_db_pct": sales_captured_in_db_pct,
                "source_file": source_file,
            }
        )
    return rows


def extract_regression_rows(
    workbook: xw.Book,
    labels: Dict[str, str],
    source_file: str,
) -> List[Dict[str, Any]]:
    rows: List[Dict[str, Any]] = []
    sheet = find_sheet(workbook, REGRESSION_SOURCE_SHEET)
    if sheet is None:
        return rows

    anchor = find_anchor(sheet, "max")
    if anchor is None:
        return rows
    anchor_row, anchor_col = anchor

    offsets = derive_offsets_from_headers(
        sheet=sheet,
        anchor_row=anchor_row,
        anchor_col=anchor_col,
        defaults=REGRESSION_DEFAULT_OFFSETS,
        synonyms=REGRESSION_HEADER_SYNONYMS,
    )

    y_col = anchor_col - 7
    x_col = anchor_col - 11
    intercept_col = anchor_col + offsets["intercept"]
    slope_col = anchor_col + offsets["slope"]
    num_quarters_col = anchor_col + offsets["num_quarters_used"]

    data_start_row = anchor_row + 1
    data_end_row = data_start_row + N_QUARTERS - 1

    formulas_written = False
    for idx in range(N_QUARTERS):
        row_num = data_start_row + idx
        num_quarters = pick(to_int(sheet.cells(row_num, num_quarters_col).value), idx + 1)
        span = max(num_quarters - 1, 0)

        intercept_y_start = r1c1_ref(row_offset=-span, col_offset=(y_col - intercept_col))
        intercept_y_end = r1c1_ref(row_offset=0, col_offset=(y_col - intercept_col))
        intercept_x_start = r1c1_ref(row_offset=-span, col_offset=(x_col - intercept_col))
        intercept_x_end = r1c1_ref(row_offset=0, col_offset=(x_col - intercept_col))
        intercept_formula = (
            f'=IFERROR(INTERCEPT({intercept_y_start}:{intercept_y_end},'
            f"{intercept_x_start}:{intercept_x_end}),\"\")"
        )
        set_formula2_r1c1(sheet.cells(row_num, intercept_col), intercept_formula)

        slope_y_start = r1c1_ref(row_offset=-span, col_offset=(y_col - slope_col))
        slope_y_end = r1c1_ref(row_offset=0, col_offset=(y_col - slope_col))
        slope_x_start = r1c1_ref(row_offset=-span, col_offset=(x_col - slope_col))
        slope_x_end = r1c1_ref(row_offset=0, col_offset=(x_col - slope_col))
        slope_formula = (
            f'=IFERROR(SLOPE({slope_y_start}:{slope_y_end},'
            f"{slope_x_start}:{slope_x_end}),\"\")"
        )
        set_formula2_r1c1(sheet.cells(row_num, slope_col), slope_formula)
        formulas_written = True

    if formulas_written:
        workbook.app.calculate()

    needed_offsets = [offsets[field] for field in REGRESSION_DEFAULT_OFFSETS.keys()]
    start_col = anchor_col + min(needed_offsets)
    end_col = anchor_col + max(needed_offsets)
    data_block = get_block_values(sheet, data_start_row, data_end_row, start_col, end_col)

    previous_signature: Optional[Tuple[Any, ...]] = None
    for idx, row_values in enumerate(data_block):
        num_quarters_used = pick(to_int(block_value(row_values, start_col, anchor_col, offsets["num_quarters_used"])), idx + 1)
        intercept = block_value(row_values, start_col, anchor_col, offsets["intercept"])
        slope = block_value(row_values, start_col, anchor_col, offsets["slope"])
        forecast_total_without_sa = block_value(row_values, start_col, anchor_col, offsets["forecast_value"])
        forecast_max = block_value(row_values, start_col, anchor_col, offsets["forecast_max"])
        forecast_min = block_value(row_values, start_col, anchor_col, offsets["forecast_min"])
        actual_value = block_value(row_values, start_col, anchor_col, offsets["actual_value"])

        signature = (
            num_quarters_used,
            to_float(intercept),
            to_float(slope),
            to_float(forecast_total_without_sa),
            to_float(forecast_max),
            to_float(forecast_min),
        )
        if previous_signature is not None and signature == previous_signature:
            continue
        previous_signature = signature

        has_data = any(
            value not in (None, "")
            for value in (num_quarters_used, intercept, slope, forecast_total_without_sa, forecast_max, forecast_min)
        )
        if not has_data:
            continue

        rows.append(
            {
                "model": labels["model"],
                "ticker": labels["ticker"],
                "model_period": labels["model_period"],
                "model_date": labels["model_date"],
                "method": "regression",
                "parameter_name": "num_quarters_used",
                "parameter_value": num_quarters_used,
                "num_quarters_used": num_quarters_used,
                "forecast_value": forecast_total_without_sa,
                "actual_value": actual_value,
                "forecast_max": forecast_max,
                "forecast_min": forecast_min,
                "range_width": safe_subtract(forecast_max, forecast_min),
                "intercept": intercept,
                "slope": slope,
                "source_file": source_file,
            }
        )
    return rows


def write_output_sheet(
    workbook: Workbook,
    sheet_name: str,
    columns: List[str],
    rows: List[Dict[str, Any]],
) -> None:
    ws = workbook.create_sheet(sheet_name)
    ws.append(columns)

    for row in rows:
        ws.append([row.get(column, "") for column in columns])

    for cell in ws[1]:
        cell.font = Font(bold=True)

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions

    for col_idx, header in enumerate(columns, start=1):
        max_len = len(header)
        for row_idx in range(2, ws.max_row + 1):
            value = ws.cell(row=row_idx, column=col_idx).value
            if value is None:
                continue
            max_len = max(max_len, len(str(value)))
        ws.column_dimensions[get_column_letter(col_idx)].width = min(max(12, max_len + 2), 48)


def process_workbooks() -> None:
    if not input_dir.exists():
        print(f"Input directory not found: {input_dir.resolve()}")
        return

    output_dir.mkdir(parents=True, exist_ok=True)
    output_path = choose_output_path(input_dir, output_dir)

    empirical_rows: List[Dict[str, Any]] = []
    regression_rows: List[Dict[str, Any]] = []
    processed_files = 0

    app = xw.App(visible=False, add_book=False)
    app.display_alerts = False
    app.screen_updating = False

    try:
        for file_path in sorted(input_dir.iterdir()):
            file_name = file_path.name
            if not file_path.is_file():
                print(f"skipped: {file_name} (not a file)")
                continue
            if file_name.startswith("~"):
                print(f"skipped: {file_name} (temporary file)")
                continue
            if file_path.suffix.lower() != ".xlsx":
                print(f"skipped: {file_name} (not .xlsx)")
                continue

            workbook: Optional[xw.Book] = None
            try:
                workbook = app.books.open(str(file_path), update_links=False)
                labels = parse_file_labels(file_name)
                empirical_rows.extend(extract_empirical_rows(workbook, labels, file_name))
                regression_rows.extend(extract_regression_rows(workbook, labels, file_name))
                processed_files += 1
                print(f"processed: {file_name}")
            except Exception as exc:
                print(f"skipped: {file_name} (error: {exc})")
            finally:
                if workbook is not None:
                    try:
                        close_source_workbook(workbook)
                    except Exception as close_exc:
                        print(f"warning: close failed for {file_name} ({close_exc})")
    finally:
        app.quit()

    out_wb = Workbook()
    out_wb.remove(out_wb.active)
    write_output_sheet(out_wb, "empirical_candidates", EMPIRICAL_COLUMNS, empirical_rows)
    write_output_sheet(out_wb, "regression_candidates", REGRESSION_COLUMNS, regression_rows)
    out_wb.save(output_path)

    print(f"output_path: {output_path.resolve()}")
    print(f"files_processed: {processed_files}")
    print(f"empirical_rows: {len(empirical_rows)}")
    print(f"regression_rows: {len(regression_rows)}")


if __name__ == "__main__":
    process_workbooks()
