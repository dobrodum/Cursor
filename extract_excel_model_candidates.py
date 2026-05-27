#!/usr/bin/env python3
from __future__ import annotations

import re
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Sequence, Tuple

import xlwings as xw
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

# Configure these two paths before running.
input_dir = Path("./input")
output_dir = Path("./output")

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

N_QUARTERS = 10

EMPIRICAL_HEADER_ALIASES = {
    "num_quarters_used": ("num quarters used", "num quarters", "quarters used", "# quarters"),
    "last_quarter_used": ("last quarter used", "last quarter"),
    "forecast_value": (
        "estimated total sold",
        "tot fcst",
        "forecast value",
        "total forecast",
    ),
    "actual_value": ("reported sales", "actual value", "actual sales", "actual"),
    "forecast_max": ("max",),
    "forecast_min": ("min",),
    "avg_penetration_pct": ("avg penetration", "average penetration"),
    "quarterly_sales": ("quarterly sales", "quarter sales"),
    "reported_sales": ("reported sales",),
    "growth_rate_pct": ("growth rate", "growth pct", "growth %"),
    "sales_captured_in_db_pct": (
        "sales captured in db",
        "sales captured",
        "captured in db",
        "penetration",
    ),
}

REGRESSION_HEADER_ALIASES = {
    "num_quarters_used": ("num quarters used", "num quarters", "quarters used", "# quarters"),
    "forecast_value": (
        "tot fcst w/o sa",
        "tot fcst without sa",
        "forecast total without sa",
        "forecast value",
        "tot fcst",
    ),
    "actual_value": ("actual value", "actual sales", "actual"),
    "forecast_max": ("max",),
    "forecast_min": ("min",),
    "intercept": ("intercept",),
    "slope": ("slope",),
}

# Anchor-relative fallbacks for workbook layouts where headers are missing.
EMPIRICAL_FALLBACK_OFFSETS = {
    "forecast_max": 0,
    "forecast_min": 1,
    "forecast_value": -1,
    "actual_value": -2,
    "num_quarters_used": -6,
    "last_quarter_used": -5,
    "avg_penetration_pct": -4,
    "sales_captured_in_db_pct": -8,
    "growth_rate_pct": -9,
    "reported_sales": -10,
    "quarterly_sales": -11,
}

REGRESSION_FALLBACK_OFFSETS = {
    "forecast_max": 0,
    "forecast_min": 1,
    "forecast_value": -1,
    "actual_value": -2,
    "num_quarters_used": -6,
    "intercept": -4,
    "slope": -3,
}

DAY_BY_BUCKET = {"early": 5, "mid": 15, "late": 25}


def normalize_text(value: Any) -> str:
    if value is None:
        return ""
    return " ".join(str(value).strip().lower().replace("_", " ").split())


def maybe_blank(value: Any) -> Optional[Any]:
    if isinstance(value, str) and value.strip() == "":
        return None
    return value


def to_float(value: Any) -> Optional[float]:
    value = maybe_blank(value)
    if value is None:
        return None
    if isinstance(value, (int, float)):
        return float(value)
    if isinstance(value, str):
        text = value.strip().replace(",", "")
        is_pct = text.endswith("%")
        if is_pct:
            text = text[:-1]
        try:
            number = float(text)
        except ValueError:
            return None
        return number / 100 if is_pct else number
    return None


def safe_subtract(left: Any, right: Any) -> Optional[float]:
    left_num = to_float(left)
    right_num = to_float(right)
    if left_num is None or right_num is None:
        return None
    return left_num - right_num


def parse_filename_label(file_path: Path) -> Dict[str, str]:
    stem = file_path.stem
    parts = [part.strip() for part in stem.split(" - ")]

    ticker = ""
    if len(parts) >= 2:
        ticker = parts[1].strip().upper()
    if not ticker:
        token_match = re.search(r"\b([A-Z]{2,6})\b", stem.upper())
        ticker = token_match.group(1) if token_match else "UNKNOWN"

    period_match = re.search(r"\b(Early|Mid|Late)\s*([A-Za-z]+)\s*(\d{4})\b", stem, re.IGNORECASE)
    model_period = "UNKNOWN"
    model_date = ""
    if period_match:
        bucket_raw, month_raw, year_raw = period_match.groups()
        bucket = bucket_raw.capitalize()
        year = int(year_raw)
        month = parse_month(month_raw)
        if month is not None:
            month_abbrev = datetime(year, month, 1).strftime("%b")
            model_period = f"{bucket}{month_abbrev}_{year}"
            day = DAY_BY_BUCKET[bucket.lower()]
            model_date = f"{year:04d}-{month:02d}-{day:02d}"

    model = f"{ticker}_{model_period}"
    return {
        "model": model,
        "ticker": ticker,
        "model_period": model_period,
        "model_date": model_date,
        "source_file": file_path.name,
    }


def parse_month(month_raw: str) -> Optional[int]:
    token = month_raw.strip()
    for fmt in ("%b", "%B"):
        try:
            return datetime.strptime(token[:3] if fmt == "%b" else token, fmt).month
        except ValueError:
            continue
    try:
        return datetime.strptime(token[:3], "%b").month
    except ValueError:
        return None


def collect_input_files(folder: Path) -> List[Path]:
    files: List[Path] = []
    if not folder.exists():
        print(f"Skipped all files: input directory does not exist: {folder}")
        return files

    for item in sorted(folder.iterdir()):
        if not item.is_file():
            continue
        if item.name.startswith("~"):
            print(f"Skipped {item.name}: temporary file")
            continue
        if item.suffix.lower() != ".xlsx":
            print(f"Skipped {item.name}: not an .xlsx file")
            continue
        files.append(item)
    return files


def build_output_path(input_folder: Path, out_dir: Path) -> Path:
    out_dir.mkdir(parents=True, exist_ok=True)
    base_name = f"{input_folder.name}_PARAM"
    candidate = out_dir / f"{base_name}.xlsx"
    if not candidate.exists():
        return candidate

    version = 1
    while True:
        candidate = out_dir / f"{base_name}.{version}.xlsx"
        if not candidate.exists():
            return candidate
        version += 1


def close_workbook_without_save(book: xw.Book) -> None:
    last_error: Optional[Exception] = None
    for closer in (
        lambda: book.close(save=False),
        lambda: book.close(False),
        lambda: book.api.Close(SaveChanges=False),
        lambda: book.close(),
    ):
        try:
            closer()
            return
        except Exception as exc:  # noqa: PERF203
            last_error = exc
    if last_error is not None:
        raise last_error


def used_range_matrix(sheet: xw.Sheet) -> Tuple[List[List[Any]], int, int]:
    used = sheet.used_range
    top_row = used.row
    left_col = used.column
    values = used.value

    if values is None:
        return [], top_row, left_col
    if not isinstance(values, list):
        return [[values]], top_row, left_col
    if values and not isinstance(values[0], list):
        return [values], top_row, left_col
    return values, top_row, left_col


def matrix_last_col(values: Sequence[Sequence[Any]], left_col: int) -> int:
    if not values:
        return left_col
    width = 1
    for row in values:
        if isinstance(row, list):
            width = max(width, len(row))
        else:
            width = max(width, 1)
    return left_col + width - 1


def matrix_get(values: Sequence[Sequence[Any]], top_row: int, left_col: int, row: int, col: int) -> Any:
    row_idx = row - top_row
    col_idx = col - left_col
    if row_idx < 0 or col_idx < 0 or row_idx >= len(values):
        return None
    row_values = values[row_idx]
    if not isinstance(row_values, list):
        return row_values if col_idx == 0 else None
    if col_idx >= len(row_values):
        return None
    return row_values[col_idx]


def find_anchor_cell(values: Sequence[Sequence[Any]], top_row: int, left_col: int, anchor_text: str) -> Optional[Tuple[int, int]]:
    anchor_norm = normalize_text(anchor_text)
    for row_idx, row_values in enumerate(values):
        if not isinstance(row_values, list):
            row_values = [row_values]
        for col_idx, value in enumerate(row_values):
            cell_norm = normalize_text(value)
            if cell_norm == anchor_norm:
                return top_row + row_idx, left_col + col_idx
    for row_idx, row_values in enumerate(values):
        if not isinstance(row_values, list):
            row_values = [row_values]
        for col_idx, value in enumerate(row_values):
            cell_norm = normalize_text(value)
            if cell_norm.startswith(anchor_norm):
                return top_row + row_idx, left_col + col_idx
    return None


def build_header_lookup(
    values: Sequence[Sequence[Any]],
    top_row: int,
    left_col: int,
    anchor_row: int,
    anchor_col: int,
    window: int = 25,
) -> Dict[int, str]:
    headers: Dict[int, str] = {}
    row_priority = [anchor_row, anchor_row - 1, anchor_row + 1]
    min_col = max(left_col, anchor_col - window)
    max_col = anchor_col + window

    for row in row_priority:
        for col in range(min_col, max_col + 1):
            value = matrix_get(values, top_row, left_col, row, col)
            text = normalize_text(value)
            if not text:
                continue
            if col not in headers or row == anchor_row:
                headers[col] = text
    return headers


def find_column_by_aliases(header_lookup: Dict[int, str], aliases: Iterable[str]) -> Optional[int]:
    for alias in aliases:
        alias_norm = normalize_text(alias)
        for col, text in header_lookup.items():
            if text == alias_norm:
                return col
    for alias in aliases:
        alias_norm = normalize_text(alias)
        for col, text in header_lookup.items():
            if alias_norm in text:
                return col
    return None


def resolve_columns(
    anchor_col: int,
    header_lookup: Dict[int, str],
    alias_lookup: Dict[str, Tuple[str, ...]],
    fallback_offsets: Dict[str, int],
) -> Dict[str, int]:
    resolved: Dict[str, int] = {}
    for field, aliases in alias_lookup.items():
        col = find_column_by_aliases(header_lookup, aliases)
        if col is None and field in fallback_offsets:
            fallback_col = anchor_col + fallback_offsets[field]
            if fallback_col > 0:
                col = fallback_col
        if col is not None:
            resolved[field] = col
    return resolved


def set_formula2_r1c1(cell: xw.Range, formula_r1c1: str) -> bool:
    try:
        cell.formula2 = formula_r1c1
        return True
    except Exception:
        pass
    try:
        cell.api.Formula2R1C1 = formula_r1c1
        return True
    except Exception:
        return False


def read_block(sheet: xw.Sheet, start_row: int, end_row: int, columns: Sequence[int]) -> Tuple[List[List[Any]], int]:
    min_col = min(columns)
    max_col = max(columns)
    data = sheet.range((start_row, min_col), (end_row, max_col)).value
    if data is None:
        return [], min_col
    if not isinstance(data, list):
        return [[data]], min_col
    if data and not isinstance(data[0], list):
        return [data], min_col
    return data, min_col


def block_get(block: Sequence[Sequence[Any]], row_idx: int, col: int, min_col: int) -> Any:
    if row_idx < 0 or row_idx >= len(block):
        return None
    row_values = block[row_idx]
    col_idx = col - min_col
    if col_idx < 0:
        return None
    if isinstance(row_values, list):
        if col_idx >= len(row_values):
            return None
        return row_values[col_idx]
    return row_values if col_idx == 0 else None


def get_col(cols: Dict[str, int], key: str) -> Optional[int]:
    return cols.get(key)


def extract_empirical_rows(book: xw.Book, file_meta: Dict[str, str]) -> List[Dict[str, Any]]:
    try:
        sheet = book.sheets["Empirical Model"]
    except Exception:
        print(f"Skipped {file_meta['source_file']} empirical: sheet 'Empirical Model' not found")
        return []

    values, top_row, left_col = used_range_matrix(sheet)
    anchor = find_anchor_cell(values, top_row, left_col, "max")
    if anchor is None:
        print(f"Skipped {file_meta['source_file']} empirical: 'max' anchor not found")
        return []
    anchor_row, anchor_col = anchor

    header_lookup = build_header_lookup(values, top_row, left_col, anchor_row, anchor_col)
    cols = resolve_columns(anchor_col, header_lookup, EMPIRICAL_HEADER_ALIASES, EMPIRICAL_FALLBACK_OFFSETS)

    start_row = anchor_row + 1
    end_row = start_row + N_QUARTERS - 1
    used_last_col = matrix_last_col(values, left_col)
    helper_col = used_last_col + 2

    sales_capture_col = get_col(cols, "sales_captured_in_db_pct")
    avg_penetration_col = get_col(cols, "avg_penetration_pct") or helper_col

    formulas_updated = False
    if sales_capture_col is not None:
        for row in range(start_row, end_row + 1):
            avg_formula = f"=IFERROR(AVERAGE(R{start_row}C{sales_capture_col}:R{row}C{sales_capture_col}),\"\")"
            formulas_updated = set_formula2_r1c1(sheet.range((row, avg_penetration_col)), avg_formula) or formulas_updated
    if formulas_updated:
        book.app.calculate()

    required_cols = [col for col in cols.values() if col is not None]
    required_cols.append(avg_penetration_col)
    block, min_col = read_block(sheet, start_row, end_row, sorted(set(required_cols)))

    rows: List[Dict[str, Any]] = []
    blank_streak = 0
    for i in range(N_QUARTERS):
        row_num = start_row + i

        forecast_max = maybe_blank(block_get(block, i, cols.get("forecast_max", anchor_col), min_col))
        forecast_min = maybe_blank(block_get(block, i, cols.get("forecast_min", anchor_col + 1), min_col))
        forecast_value = maybe_blank(block_get(block, i, cols.get("forecast_value", anchor_col - 1), min_col))

        if forecast_max is None and forecast_min is None and forecast_value is None:
            blank_streak += 1
            if blank_streak >= 2:
                break
            continue
        blank_streak = 0

        avg_penetration = maybe_blank(block_get(block, i, avg_penetration_col, min_col))
        row_data = {
            "model": file_meta["model"],
            "ticker": file_meta["ticker"],
            "model_period": file_meta["model_period"],
            "model_date": file_meta["model_date"],
            "method": "empirical",
            "parameter_name": "avg_penetration_pct",
            "parameter_value": avg_penetration,
            "num_quarters_used": maybe_blank(
                block_get(block, i, cols.get("num_quarters_used", anchor_col - 6), min_col)
            )
            or (i + 1),
            "last_quarter_used": maybe_blank(block_get(block, i, cols.get("last_quarter_used", anchor_col - 5), min_col)),
            "forecast_value": forecast_value,
            "actual_value": maybe_blank(block_get(block, i, cols.get("actual_value", anchor_col - 2), min_col)),
            "forecast_max": forecast_max,
            "forecast_min": forecast_min,
            "range_width": safe_subtract(forecast_max, forecast_min),
            "avg_penetration_pct": avg_penetration,
            "quarterly_sales": maybe_blank(block_get(block, i, cols.get("quarterly_sales", anchor_col - 11), min_col)),
            "reported_sales": maybe_blank(block_get(block, i, cols.get("reported_sales", anchor_col - 10), min_col)),
            "growth_rate_pct": maybe_blank(block_get(block, i, cols.get("growth_rate_pct", anchor_col - 9), min_col)),
            "sales_captured_in_db_pct": maybe_blank(
                block_get(block, i, cols.get("sales_captured_in_db_pct", anchor_col - 8), min_col)
            ),
            "source_file": file_meta["source_file"],
        }
        rows.append(row_data)

    return rows


def signature_value(value: Any) -> Any:
    number = to_float(value)
    if number is not None:
        return round(number, 10)
    if isinstance(value, str):
        return value.strip()
    return value


def extract_regression_rows(book: xw.Book, file_meta: Dict[str, str]) -> List[Dict[str, Any]]:
    try:
        sheet = book.sheets["Regression Model"]
    except Exception:
        print(f"Skipped {file_meta['source_file']} regression: sheet 'Regression Model' not found")
        return []

    values, top_row, left_col = used_range_matrix(sheet)
    anchor = find_anchor_cell(values, top_row, left_col, "max")
    if anchor is None:
        print(f"Skipped {file_meta['source_file']} regression: 'max' anchor not found")
        return []
    anchor_row, anchor_col = anchor

    y_col = anchor_col - 7
    x_col = anchor_col - 11
    if y_col <= 0 or x_col <= 0:
        print(f"Skipped {file_meta['source_file']} regression: invalid x/y column offsets from anchor")
        return []

    header_lookup = build_header_lookup(values, top_row, left_col, anchor_row, anchor_col)
    cols = resolve_columns(anchor_col, header_lookup, REGRESSION_HEADER_ALIASES, REGRESSION_FALLBACK_OFFSETS)

    start_row = anchor_row + 1
    end_row = start_row + N_QUARTERS - 1
    used_last_col = matrix_last_col(values, left_col)
    helper_start_col = used_last_col + 2

    intercept_col = get_col(cols, "intercept") or helper_start_col
    slope_col = get_col(cols, "slope") or (helper_start_col + 1)

    formulas_updated = False
    for row in range(start_row, end_row + 1):
        intercept_formula = (
            f"=IFERROR(INTERCEPT(R{start_row}C{y_col}:R{row}C{y_col},R{start_row}C{x_col}:R{row}C{x_col}),\"\")"
        )
        slope_formula = f"=IFERROR(SLOPE(R{start_row}C{y_col}:R{row}C{y_col},R{start_row}C{x_col}:R{row}C{x_col}),\"\")"
        formulas_updated = set_formula2_r1c1(sheet.range((row, intercept_col)), intercept_formula) or formulas_updated
        formulas_updated = set_formula2_r1c1(sheet.range((row, slope_col)), slope_formula) or formulas_updated
    if formulas_updated:
        book.app.calculate()

    required_cols = [col for col in cols.values() if col is not None]
    required_cols.extend([intercept_col, slope_col])
    block, min_col = read_block(sheet, start_row, end_row, sorted(set(required_cols)))

    rows: List[Dict[str, Any]] = []
    previous_signature: Optional[Tuple[Any, ...]] = None
    blank_streak = 0

    for i in range(N_QUARTERS):
        forecast_max = maybe_blank(block_get(block, i, cols.get("forecast_max", anchor_col), min_col))
        forecast_min = maybe_blank(block_get(block, i, cols.get("forecast_min", anchor_col + 1), min_col))
        forecast_value = maybe_blank(block_get(block, i, cols.get("forecast_value", anchor_col - 1), min_col))
        intercept_value = maybe_blank(block_get(block, i, intercept_col, min_col))
        slope_value = maybe_blank(block_get(block, i, slope_col, min_col))

        if all(value is None for value in (forecast_max, forecast_min, forecast_value, intercept_value, slope_value)):
            blank_streak += 1
            if blank_streak >= 2:
                break
            continue
        blank_streak = 0

        num_quarters_used = maybe_blank(block_get(block, i, cols.get("num_quarters_used", anchor_col - 6), min_col))
        if num_quarters_used is None:
            num_quarters_used = i + 1

        row_signature = (
            signature_value(num_quarters_used),
            signature_value(forecast_value),
            signature_value(forecast_max),
            signature_value(forecast_min),
            signature_value(intercept_value),
            signature_value(slope_value),
        )
        if previous_signature is not None and row_signature == previous_signature:
            continue
        previous_signature = row_signature

        row_data = {
            "model": file_meta["model"],
            "ticker": file_meta["ticker"],
            "model_period": file_meta["model_period"],
            "model_date": file_meta["model_date"],
            "method": "regression",
            "parameter_name": "num_quarters_used",
            "parameter_value": num_quarters_used,
            "num_quarters_used": num_quarters_used,
            "forecast_value": forecast_value,
            "actual_value": maybe_blank(block_get(block, i, cols.get("actual_value", anchor_col - 2), min_col)),
            "forecast_max": forecast_max,
            "forecast_min": forecast_min,
            "range_width": safe_subtract(forecast_max, forecast_min),
            "intercept": intercept_value,
            "slope": slope_value,
            "source_file": file_meta["source_file"],
        }
        rows.append(row_data)

    return rows


def append_rows(sheet, column_order: Sequence[str], rows: Sequence[Dict[str, Any]]) -> None:
    for row in rows:
        sheet.append([row.get(column) for column in column_order])


def format_sheet(sheet) -> None:
    for cell in sheet[1]:
        cell.font = Font(bold=True)
    sheet.freeze_panes = "A2"
    sheet.auto_filter.ref = sheet.dimensions

    for col_idx in range(1, sheet.max_column + 1):
        max_len = 0
        for row_idx in range(1, sheet.max_row + 1):
            value = sheet.cell(row=row_idx, column=col_idx).value
            length = len(str(value)) if value is not None else 0
            if length > max_len:
                max_len = length
        sheet.column_dimensions[get_column_letter(col_idx)].width = min(max(max_len + 2, 12), 40)


def create_output_workbook() -> Tuple[Workbook, Any, Any]:
    wb = Workbook()
    empirical_sheet = wb.active
    empirical_sheet.title = "empirical_candidates"
    empirical_sheet.append(EMPIRICAL_COLUMNS)

    regression_sheet = wb.create_sheet("regression_candidates")
    regression_sheet.append(REGRESSION_COLUMNS)
    return wb, empirical_sheet, regression_sheet


def main() -> None:
    files = collect_input_files(input_dir)
    output_path = build_output_path(input_dir, output_dir)

    output_wb, empirical_out, regression_out = create_output_workbook()
    processed_files = 0
    empirical_row_count = 0
    regression_row_count = 0

    if not files:
        format_sheet(empirical_out)
        format_sheet(regression_out)
        output_wb.save(output_path)
        print(f"Output path: {output_path}")
        print("Number of files processed: 0")
        print("Number of empirical rows: 0")
        print("Number of regression rows: 0")
        return

    app = xw.App(visible=False, add_book=False)
    app.display_alerts = False
    app.screen_updating = False
    try:
        # xlCalculationManual
        app.api.Calculation = -4135
    except Exception:
        pass

    try:
        for file_path in files:
            print(f"Processing file: {file_path.name}")
            file_meta = parse_filename_label(file_path)
            source_book: Optional[xw.Book] = None
            try:
                source_book = app.books.open(str(file_path), update_links=False)
                empirical_rows = extract_empirical_rows(source_book, file_meta)
                regression_rows = extract_regression_rows(source_book, file_meta)

                append_rows(empirical_out, EMPIRICAL_COLUMNS, empirical_rows)
                append_rows(regression_out, REGRESSION_COLUMNS, regression_rows)

                empirical_row_count += len(empirical_rows)
                regression_row_count += len(regression_rows)
                processed_files += 1
            except Exception as exc:
                print(f"Skipped {file_path.name}: processing error ({exc})")
            finally:
                if source_book is not None:
                    try:
                        close_workbook_without_save(source_book)
                    except Exception as close_exc:
                        print(f"Skipped {file_path.name}: close-without-save fallback failed ({close_exc})")
    finally:
        app.quit()

    format_sheet(empirical_out)
    format_sheet(regression_out)
    output_wb.save(output_path)

    print(f"Output path: {output_path}")
    print(f"Number of files processed: {processed_files}")
    print(f"Number of empirical rows: {empirical_row_count}")
    print(f"Number of regression rows: {regression_row_count}")


if __name__ == "__main__":
    main()
