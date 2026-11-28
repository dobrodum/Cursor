#!/usr/bin/env python3
from __future__ import annotations

import re
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List, Optional, Sequence, Tuple

import xlwings as xw
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

# User-configurable paths
input_dir = "/path/to/input"
output_dir = "/path/to/output"

EMPIRICAL_HEADERS = [
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

REGRESSION_HEADERS = [
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

EMPIRICAL_ALIASES: Dict[str, Sequence[str]] = {
    "num_quarters_used": ("num quarters used", "quarters used", "# quarters", "n quarters"),
    "last_quarter_used": ("last quarter used", "last qtr used", "last quarter"),
    "forecast_value": ("estimated total sold", "total sold", "forecast", "tot fcst"),
    "reported_sales": ("reported sales", "actual sales", "actual value"),
    "forecast_max": ("max",),
    "forecast_min": ("min",),
    "avg_penetration_pct": ("avg penetration", "average penetration", "avg penetration %"),
    "quarterly_sales": ("quarterly sales", "qtr sales", "quarter sales"),
    "growth_rate_pct": ("growth rate", "growth rate %", "growth %"),
    "sales_captured_in_db_pct": (
        "sales captured in db",
        "captured in db",
        "captured in database",
        "db capture %",
    ),
}

REGRESSION_ALIASES: Dict[str, Sequence[str]] = {
    "num_quarters_used": ("num quarters used", "quarters used", "# quarters", "n quarters"),
    "forecast_value": (
        "tot fcst w/o sa",
        "tot fcst wo sa",
        "total forecast without sa",
        "forecast without sa",
        "tot fcst",
    ),
    "actual_value": ("actual sales", "reported sales", "actual value", "actual"),
    "forecast_max": ("max",),
    "forecast_min": ("min",),
}

DAY_BY_PERIOD = {"Early": 5, "Mid": 15, "Late": 25}
MONTH_BY_TOKEN = {
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


def normalize_text(value: Any) -> str:
    if value is None:
        return ""
    text = str(value).strip().lower()
    text = re.sub(r"\s+", " ", text)
    text = re.sub(r"[^a-z0-9%/ ]+", "", text)
    return text


def to_2d(values: Any) -> List[List[Any]]:
    if values is None:
        return []
    if not isinstance(values, (list, tuple)):
        return [[values]]
    if not values:
        return []
    first = values[0]
    if isinstance(first, (list, tuple)):
        normalized: List[List[Any]] = []
        for row in values:
            if isinstance(row, (list, tuple)):
                normalized.append(list(row))
            else:
                normalized.append([row])
        return normalized
    return [list(values)]


def to_number(value: Any) -> Optional[float]:
    if value is None:
        return None
    if isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        return float(value)
    text = str(value).strip()
    if not text:
        return None
    text = text.replace(",", "")
    if text.endswith("%"):
        text = text[:-1]
    try:
        return float(text)
    except ValueError:
        return None


def normalize_output_value(value: Any) -> Any:
    if value is None:
        return None
    if isinstance(value, float) and value.is_integer():
        return int(value)
    return value


def safe_diff(max_value: Any, min_value: Any) -> Optional[float]:
    max_num = to_number(max_value)
    min_num = to_number(min_value)
    if max_num is None or min_num is None:
        return None
    return max_num - min_num


def parse_month_number(token: str) -> Optional[int]:
    cleaned = re.sub(r"[^A-Za-z]", "", token).lower()
    if not cleaned:
        return None
    if cleaned in MONTH_BY_TOKEN:
        return MONTH_BY_TOKEN[cleaned]
    short = cleaned[:3]
    return MONTH_BY_TOKEN.get(short)


def parse_file_metadata(file_path: Path) -> Dict[str, str]:
    stem = file_path.stem
    parts = [part.strip() for part in stem.split("-")]
    ticker = ""
    if len(parts) >= 2:
        ticker = re.sub(r"[^A-Za-z0-9]", "", parts[1]).upper()
    if not ticker:
        ticker_match = re.search(r"\b([A-Z]{2,8})\b", stem)
        if ticker_match:
            ticker = ticker_match.group(1).upper()

    period_source = parts[2] if len(parts) >= 3 else stem
    period_source = period_source.split("_")[0]
    period_match = re.search(
        r"(Early|Mid|Late)\s*([A-Za-z]+)\s*(20\d{2})",
        period_source,
        flags=re.IGNORECASE,
    )
    if not period_match:
        period_match = re.search(
            r"(Early|Mid|Late)\s*([A-Za-z]+)\s*(20\d{2})",
            stem,
            flags=re.IGNORECASE,
        )

    model_period = ""
    model_date = ""
    if period_match:
        period_bucket = period_match.group(1).title()
        month_token = period_match.group(2)
        year = int(period_match.group(3))
        month = parse_month_number(month_token)
        if month is not None:
            month_short = datetime(year, month, 1).strftime("%b")
            model_period = f"{period_bucket}{month_short}_{year}"
            model_date = f"{year:04d}-{month:02d}-{DAY_BY_PERIOD[period_bucket]:02d}"

    if ticker and model_period:
        model = f"{ticker}_{model_period}"
    elif ticker:
        model = ticker
    else:
        model = stem

    return {
        "model": model,
        "ticker": ticker,
        "model_period": model_period,
        "model_date": model_date,
    }


def next_output_path(input_path: Path, output_path: Path) -> Path:
    base_name = f"{input_path.name}_PARAM.xlsx"
    candidate = output_path / base_name
    if not candidate.exists():
        return candidate
    suffix = 1
    while True:
        candidate = output_path / f"{input_path.name}_PARAM.{suffix}.xlsx"
        if not candidate.exists():
            return candidate
        suffix += 1


def safe_close_workbook(wb: xw.Book) -> None:
    try:
        wb.close(save=False)
        return
    except TypeError:
        pass
    except Exception:
        pass

    try:
        wb.close(False)
        return
    except Exception:
        pass

    try:
        wb.api.Close(SaveChanges=False)
    except Exception:
        pass


def find_max_anchor(sheet: xw.Sheet) -> Tuple[int, int]:
    used = sheet.used_range
    values = to_2d(used.value)
    base_row = used.row
    base_col = used.column

    for row_idx, row_values in enumerate(values):
        for col_idx, value in enumerate(row_values):
            if normalize_text(value) == "max":
                return base_row + row_idx, base_col + col_idx

    raise ValueError(f"Could not find 'max' anchor in sheet '{sheet.name}'.")


def build_header_map(
    sheet: xw.Sheet,
    header_row: int,
    anchor_col: int,
    radius: int = 40,
) -> Dict[str, int]:
    start_col = max(1, anchor_col - radius)
    end_col = anchor_col + radius
    row_values = sheet.range((header_row, start_col), (header_row, end_col)).value
    if not isinstance(row_values, (list, tuple)):
        row_values = [row_values]

    header_map: Dict[str, int] = {}
    for idx, value in enumerate(row_values):
        label = normalize_text(value)
        if label and label not in header_map:
            header_map[label] = start_col + idx
    return header_map


def find_col(
    header_map: Dict[str, int],
    aliases: Sequence[str],
    fallback: Optional[int] = None,
) -> Optional[int]:
    for alias in aliases:
        alias_key = normalize_text(alias)
        if alias_key in header_map:
            return header_map[alias_key]
    for alias in aliases:
        alias_key = normalize_text(alias)
        for header, col in header_map.items():
            if alias_key and alias_key in header:
                return col
    if fallback is None or fallback < 1:
        return None
    return fallback


def get_cell(sheet: xw.Sheet, row: int, col: Optional[int]) -> Any:
    if col is None or col < 1:
        return None
    return sheet.range((row, col)).value


def signature_value(value: Any) -> Any:
    num = to_number(value)
    if num is not None:
        return round(num, 10)
    if isinstance(value, str):
        return value.strip()
    return value


def extract_empirical_candidates(
    wb: xw.Book,
    metadata: Dict[str, str],
    source_file: str,
) -> List[List[Any]]:
    try:
        sheet = wb.sheets["Empirical Model"]
    except Exception:
        print(f"skipped file: {source_file} (missing 'Empirical Model' sheet)")
        return []

    anchor_row, anchor_col = find_max_anchor(sheet)
    header_map = build_header_map(sheet, anchor_row, anchor_col)

    num_quarters_col = find_col(header_map, EMPIRICAL_ALIASES["num_quarters_used"], anchor_col - 6)
    last_quarter_col = find_col(header_map, EMPIRICAL_ALIASES["last_quarter_used"], anchor_col - 5)
    forecast_col = find_col(header_map, EMPIRICAL_ALIASES["forecast_value"], anchor_col - 1)
    reported_col = find_col(header_map, EMPIRICAL_ALIASES["reported_sales"], anchor_col - 2)
    max_col = find_col(header_map, EMPIRICAL_ALIASES["forecast_max"], anchor_col)
    min_col = find_col(header_map, EMPIRICAL_ALIASES["forecast_min"], anchor_col + 1)
    avg_pen_col = find_col(header_map, EMPIRICAL_ALIASES["avg_penetration_pct"], anchor_col - 4)
    quarterly_sales_col = find_col(header_map, EMPIRICAL_ALIASES["quarterly_sales"], anchor_col - 8)
    growth_rate_col = find_col(header_map, EMPIRICAL_ALIASES["growth_rate_pct"], anchor_col - 7)
    sales_db_col = find_col(header_map, EMPIRICAL_ALIASES["sales_captured_in_db_pct"], anchor_col - 9)

    helper_avg_col = sheet.used_range.last_cell.column + 3
    formula_rows: Dict[int, bool] = {}
    formulas_written = False
    history_end_row = anchor_row - 1
    avg_source_col = sales_db_col or avg_pen_col

    # Keep n_quarters fixed to 10 and use R1C1 formula2 for average penetration.
    for idx in range(10):
        row = anchor_row + 1 + idx
        num_quarters_raw = get_cell(sheet, row, num_quarters_col)
        num_quarters = to_number(num_quarters_raw)
        n_quarters = int(num_quarters) if num_quarters is not None else idx + 1
        if avg_source_col and history_end_row >= 1 and n_quarters > 0:
            start_row = max(1, history_end_row - n_quarters + 1)
            formula = (
                f'=IFERROR(AVERAGE(R{start_row}C{avg_source_col}:R{history_end_row}C{avg_source_col}),"")'
            )
            sheet.range((row, helper_avg_col)).formula2 = formula
            formulas_written = True
            formula_rows[row] = True

    if formulas_written:
        wb.app.calculate()

    rows: List[List[Any]] = []
    empty_streak = 0
    for idx in range(10):
        row = anchor_row + 1 + idx
        num_quarters_raw = get_cell(sheet, row, num_quarters_col)
        num_quarters = to_number(num_quarters_raw)
        num_quarters_value: Any = int(num_quarters) if num_quarters is not None else idx + 1

        last_quarter_used = get_cell(sheet, row, last_quarter_col)
        forecast_value = get_cell(sheet, row, forecast_col)
        reported_sales = get_cell(sheet, row, reported_col)
        forecast_max = get_cell(sheet, row, max_col)
        forecast_min = get_cell(sheet, row, min_col)
        quarterly_sales = get_cell(sheet, row, quarterly_sales_col)
        growth_rate_pct = get_cell(sheet, row, growth_rate_col)
        sales_captured_in_db_pct = get_cell(sheet, row, sales_db_col)

        if formula_rows.get(row):
            avg_penetration_pct = get_cell(sheet, row, helper_avg_col)
        else:
            avg_penetration_pct = get_cell(sheet, row, avg_pen_col)

        if avg_penetration_pct in (None, "") and avg_pen_col:
            avg_penetration_pct = get_cell(sheet, row, avg_pen_col)

        if forecast_value in (None, ""):
            quarterly_sales_num = to_number(quarterly_sales)
            avg_pen_num = to_number(avg_penetration_pct)
            if quarterly_sales_num is not None and avg_pen_num not in (None, 0):
                forecast_value = quarterly_sales_num / avg_pen_num

        row_has_data = any(
            value not in (None, "")
            for value in (
                forecast_value,
                reported_sales,
                forecast_max,
                forecast_min,
                avg_penetration_pct,
            )
        )
        if not row_has_data:
            empty_streak += 1
            if empty_streak >= 2 and idx > 0:
                break
            continue
        empty_streak = 0

        range_width = safe_diff(forecast_max, forecast_min)
        actual_value = reported_sales

        rows.append(
            [
                metadata["model"],
                metadata["ticker"],
                metadata["model_period"],
                metadata["model_date"],
                "empirical",
                "avg_penetration_pct",
                normalize_output_value(avg_penetration_pct),
                normalize_output_value(num_quarters_value),
                normalize_output_value(last_quarter_used),
                normalize_output_value(forecast_value),
                normalize_output_value(actual_value),
                normalize_output_value(forecast_max),
                normalize_output_value(forecast_min),
                normalize_output_value(range_width),
                normalize_output_value(avg_penetration_pct),
                normalize_output_value(quarterly_sales),
                normalize_output_value(reported_sales),
                normalize_output_value(growth_rate_pct),
                normalize_output_value(sales_captured_in_db_pct),
                source_file,
            ]
        )

    return rows


def extract_regression_candidates(
    wb: xw.Book,
    metadata: Dict[str, str],
    source_file: str,
) -> List[List[Any]]:
    try:
        sheet = wb.sheets["Regression Model"]
    except Exception:
        print(f"skipped file: {source_file} (missing 'Regression Model' sheet)")
        return []

    anchor_row, anchor_col = find_max_anchor(sheet)
    header_map = build_header_map(sheet, anchor_row, anchor_col)

    # Required anchor-based offsets from existing model structure.
    y_col = max(1, anchor_col - 7)
    x_col = max(1, anchor_col - 11)

    num_quarters_col = find_col(header_map, REGRESSION_ALIASES["num_quarters_used"], anchor_col - 6)
    forecast_col = find_col(header_map, REGRESSION_ALIASES["forecast_value"], anchor_col - 1)
    actual_col = find_col(header_map, REGRESSION_ALIASES["actual_value"], None)
    max_col = find_col(header_map, REGRESSION_ALIASES["forecast_max"], anchor_col)
    min_col = find_col(header_map, REGRESSION_ALIASES["forecast_min"], anchor_col + 1)

    helper_intercept_col = sheet.used_range.last_cell.column + 3
    helper_slope_col = helper_intercept_col + 1
    history_end_row = anchor_row - 1
    formulas_written = False

    # Keep n_quarters fixed to 10 and use R1C1 formula2 for INTERCEPT/SLOPE.
    for idx in range(10):
        row = anchor_row + 1 + idx
        num_quarters_raw = get_cell(sheet, row, num_quarters_col)
        num_quarters = to_number(num_quarters_raw)
        n_quarters = int(num_quarters) if num_quarters is not None else idx + 1
        if n_quarters <= 1:
            continue

        start_row = max(1, history_end_row - n_quarters + 1)
        sample_size = history_end_row - start_row + 1
        if sample_size < 2:
            continue

        intercept_formula = (
            f'=IFERROR(INTERCEPT(R{start_row}C{y_col}:R{history_end_row}C{y_col},'
            f'R{start_row}C{x_col}:R{history_end_row}C{x_col}),"")'
        )
        slope_formula = (
            f'=IFERROR(SLOPE(R{start_row}C{y_col}:R{history_end_row}C{y_col},'
            f'R{start_row}C{x_col}:R{history_end_row}C{x_col}),"")'
        )
        sheet.range((row, helper_intercept_col)).formula2 = intercept_formula
        sheet.range((row, helper_slope_col)).formula2 = slope_formula
        formulas_written = True

    if formulas_written:
        wb.app.calculate()

    rows: List[List[Any]] = []
    previous_signature: Optional[Tuple[Any, ...]] = None
    empty_streak = 0
    for idx in range(10):
        row = anchor_row + 1 + idx
        num_quarters_raw = get_cell(sheet, row, num_quarters_col)
        num_quarters = to_number(num_quarters_raw)
        num_quarters_value: Any = int(num_quarters) if num_quarters is not None else idx + 1

        forecast_value = get_cell(sheet, row, forecast_col)
        actual_value = get_cell(sheet, row, actual_col)
        forecast_max = get_cell(sheet, row, max_col)
        forecast_min = get_cell(sheet, row, min_col)
        intercept = get_cell(sheet, row, helper_intercept_col)
        slope = get_cell(sheet, row, helper_slope_col)

        row_has_data = any(
            value not in (None, "")
            for value in (forecast_value, forecast_max, forecast_min, intercept, slope)
        )
        if not row_has_data:
            empty_streak += 1
            if empty_streak >= 2 and idx > 0:
                break
            continue
        empty_streak = 0

        range_width = safe_diff(forecast_max, forecast_min)
        signature = (
            signature_value(num_quarters_value),
            signature_value(intercept),
            signature_value(slope),
            signature_value(forecast_value),
            signature_value(forecast_max),
            signature_value(forecast_min),
        )
        if previous_signature is not None and signature == previous_signature:
            continue
        previous_signature = signature

        rows.append(
            [
                metadata["model"],
                metadata["ticker"],
                metadata["model_period"],
                metadata["model_date"],
                "regression",
                "num_quarters_used",
                normalize_output_value(num_quarters_value),
                normalize_output_value(num_quarters_value),
                normalize_output_value(forecast_value),
                normalize_output_value(actual_value),
                normalize_output_value(forecast_max),
                normalize_output_value(forecast_min),
                normalize_output_value(range_width),
                normalize_output_value(intercept),
                normalize_output_value(slope),
                source_file,
            ]
        )

    return rows


def apply_sheet_formatting(ws) -> None:
    for cell in ws[1]:
        cell.font = Font(bold=True)
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions

    for col_idx in range(1, ws.max_column + 1):
        max_len = 0
        for row_idx in range(1, ws.max_row + 1):
            value = ws.cell(row=row_idx, column=col_idx).value
            if value is None:
                continue
            max_len = max(max_len, len(str(value)))
        ws.column_dimensions[get_column_letter(col_idx)].width = min(max(12, max_len + 2), 48)


def write_output_workbook(
    output_path: Path,
    empirical_rows: List[List[Any]],
    regression_rows: List[List[Any]],
) -> None:
    wb = Workbook()
    default_sheet = wb.active
    wb.remove(default_sheet)

    empirical_ws = wb.create_sheet("empirical_candidates")
    empirical_ws.append(EMPIRICAL_HEADERS)
    for row in empirical_rows:
        empirical_ws.append(row)
    apply_sheet_formatting(empirical_ws)

    regression_ws = wb.create_sheet("regression_candidates")
    regression_ws.append(REGRESSION_HEADERS)
    for row in regression_rows:
        regression_ws.append(row)
    apply_sheet_formatting(regression_ws)

    wb.save(output_path)


def should_skip_file(path: Path) -> Optional[str]:
    if not path.is_file():
        return "not a file"
    if path.name.startswith("~"):
        return "temporary file"
    if path.suffix.lower() != ".xlsx":
        return "not .xlsx"
    return None


def main() -> None:
    input_path = Path(input_dir).expanduser().resolve()
    output_path = Path(output_dir).expanduser().resolve()

    if not input_path.exists() or not input_path.is_dir():
        raise FileNotFoundError(f"Input directory not found: {input_path}")
    output_path.mkdir(parents=True, exist_ok=True)

    output_file = next_output_path(input_path, output_path)
    output_pattern = re.compile(
        rf"^{re.escape(input_path.name)}_PARAM(?:\.\d+)?\.xlsx$",
        flags=re.IGNORECASE,
    )

    empirical_rows: List[List[Any]] = []
    regression_rows: List[List[Any]] = []
    files_processed = 0

    app = xw.App(visible=False, add_book=False)
    app.display_alerts = False
    app.screen_updating = False
    try:
        for file_path in sorted(input_path.iterdir()):
            skip_reason = should_skip_file(file_path)
            if skip_reason:
                print(f"skipped file: {file_path.name} ({skip_reason})")
                continue
            if output_pattern.match(file_path.name):
                print(f"skipped file: {file_path.name} (generated output file pattern)")
                continue

            wb = None
            try:
                wb = app.books.open(str(file_path), update_links=False)
                metadata = parse_file_metadata(file_path)
                empirical_rows.extend(extract_empirical_candidates(wb, metadata, file_path.name))
                regression_rows.extend(extract_regression_candidates(wb, metadata, file_path.name))
                files_processed += 1
                print(f"processed file: {file_path.name}")
            except Exception as exc:
                print(f"skipped file: {file_path.name} (error: {exc})")
            finally:
                if wb is not None:
                    safe_close_workbook(wb)
    finally:
        app.quit()

    write_output_workbook(output_file, empirical_rows, regression_rows)

    print(f"output path: {output_file}")
    print(f"number of files processed: {files_processed}")
    print(f"number of empirical rows: {len(empirical_rows)}")
    print(f"number of regression rows: {len(regression_rows)}")


if __name__ == "__main__":
    main()
