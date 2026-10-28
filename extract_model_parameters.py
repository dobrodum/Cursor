#!/usr/bin/env python3
from __future__ import annotations

import re
from dataclasses import dataclass
from datetime import date
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Sequence, Tuple

import xlwings as xw
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

# --- Configure these two paths ---
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

PHASE_DAY_MAP = {"early": 5, "mid": 15, "late": 25}

FILENAME_PATTERN = re.compile(
    r"-\s*(?P<ticker>[A-Za-z0-9]+)\s*-\s*(?P<period>(Early|Mid|Late)[A-Za-z]{3}\d{4})",
    re.IGNORECASE,
)
PERIOD_PATTERN = re.compile(
    r"^(?P<phase>Early|Mid|Late)(?P<month>[A-Za-z]{3})(?P<year>\d{4})$",
    re.IGNORECASE,
)


@dataclass(frozen=True)
class ModelMeta:
    model: str
    ticker: str
    model_period: str
    model_date: str


def normalize_text(value: Any) -> str:
    if value is None:
        return ""
    text = str(value).strip().lower()
    text = text.replace("\n", " ")
    text = re.sub(r"[^a-z0-9%/ ]+", "", text)
    text = re.sub(r"\s+", " ", text).strip()
    return text


def to_float(value: Any) -> Optional[float]:
    if value in (None, ""):
        return None
    try:
        return float(value)
    except (TypeError, ValueError):
        return None


def to_int(value: Any) -> Optional[int]:
    numeric = to_float(value)
    if numeric is None:
        return None
    return int(numeric)


def as_matrix(values: Any) -> List[List[Any]]:
    if isinstance(values, list):
        if values and isinstance(values[0], list):
            return values
        return [values]
    return [[values]]


def parse_model_meta(file_path: Path) -> ModelMeta:
    stem = file_path.stem
    match = FILENAME_PATTERN.search(stem)
    if not match:
        ticker = ""
        model_period = ""
        model_date = ""
        model = stem
        return ModelMeta(model=model, ticker=ticker, model_period=model_period, model_date=model_date)

    ticker = match.group("ticker").upper()
    period_token = match.group("period")
    period_match = PERIOD_PATTERN.match(period_token)

    if not period_match:
        model_period = period_token
        model_date = ""
        model = f"{ticker}_{model_period}"
        return ModelMeta(model=model, ticker=ticker, model_period=model_period, model_date=model_date)

    phase = period_match.group("phase").capitalize()
    month_abbrev = period_match.group("month").capitalize()
    year_text = period_match.group("year")
    month_number = MONTH_MAP.get(month_abbrev.lower())
    phase_day = PHASE_DAY_MAP.get(phase.lower())

    model_period = f"{phase}{month_abbrev}_{year_text}"
    model_date = ""
    if month_number is not None and phase_day is not None:
        model_date = date(int(year_text), month_number, phase_day).isoformat()

    model = f"{ticker}_{model_period}"
    return ModelMeta(model=model, ticker=ticker, model_period=model_period, model_date=model_date)


def get_output_path(input_folder: Path, target_dir: Path) -> Path:
    base_stem = f"{input_folder.name}_PARAM"
    candidate = target_dir / f"{base_stem}.xlsx"
    index = 1
    while candidate.exists():
        candidate = target_dir / f"{base_stem}.{index}.xlsx"
        index += 1
    return candidate


def find_anchor(sheet: xw.main.Sheet, target: str = "max") -> Optional[Tuple[int, int]]:
    used = sheet.used_range
    values = as_matrix(used.value)
    row0 = used.row
    col0 = used.column
    target_norm = normalize_text(target)

    for r_idx, row_values in enumerate(values):
        for c_idx, cell_value in enumerate(row_values):
            if normalize_text(cell_value) == target_norm:
                return row0 + r_idx, col0 + c_idx
    return None


def collect_headers(
    sheet: xw.main.Sheet, anchor_row: int, anchor_col: int, span: int = 24
) -> Tuple[Dict[str, int], int, int]:
    rows = [r for r in (anchor_row - 1, anchor_row, anchor_row + 1) if r >= 1]
    col_start = max(1, anchor_col - span)
    col_end = anchor_col + span
    header_lookup: Dict[str, int] = {}

    for row in rows:
        row_values = sheet.range((row, col_start), (row, col_end)).value
        if not isinstance(row_values, list):
            row_values = [row_values]
        for offset, value in enumerate(row_values):
            key = normalize_text(value)
            if key and key not in header_lookup:
                header_lookup[key] = col_start + offset

    return header_lookup, col_start, col_end


def find_column(header_lookup: Dict[str, int], keyword_sets: Sequence[Sequence[str]]) -> Optional[int]:
    for keywords in keyword_sets:
        for header, col in header_lookup.items():
            if all(keyword in header for keyword in keywords):
                return col
    return None


def resolve_column(
    header_lookup: Dict[str, int],
    keyword_sets: Sequence[Sequence[str]],
    anchor_col: int,
    fallback_offset: int,
) -> int:
    header_match = find_column(header_lookup, keyword_sets)
    if header_match is not None:
        return header_match
    return max(1, anchor_col + fallback_offset)


def read_cell(sheet: xw.main.Sheet, row: int, col: Optional[int]) -> Any:
    if col is None or col < 1:
        return None
    value = sheet.cells(row, col).value
    if isinstance(value, str):
        value = value.strip()
    if value == "":
        return None
    return value


def write_formula2(cell: xw.main.Range, formula: str) -> None:
    try:
        cell.formula2 = formula
    except Exception:
        # Fallback only for environments where formula2 is unavailable.
        cell.formula = formula


def numeric_signature(values: Iterable[Any]) -> Tuple[Any, ...]:
    signature: List[Any] = []
    for value in values:
        numeric = to_float(value)
        if numeric is None:
            signature.append(value)
        else:
            signature.append(round(numeric, 10))
    return tuple(signature)


def get_sheet(wb: xw.main.Book, sheet_name: str) -> Optional[xw.main.Sheet]:
    target = normalize_text(sheet_name)
    for sheet in wb.sheets:
        if normalize_text(sheet.name) == target:
            return sheet
    return None


def process_empirical_sheet(
    wb: xw.main.Book, sheet: xw.main.Sheet, meta: ModelMeta, source_file: str
) -> List[Dict[str, Any]]:
    anchor = find_anchor(sheet, target="max")
    if anchor is None:
        return []
    anchor_row, anchor_col = anchor
    headers, _, col_end = collect_headers(sheet, anchor_row, anchor_col)

    num_quarters_col = resolve_column(
        headers,
        keyword_sets=(("num", "quarter"), ("quarter", "used"), ("n", "quarter")),
        anchor_col=anchor_col,
        fallback_offset=-7,
    )
    last_quarter_col = resolve_column(
        headers,
        keyword_sets=(("last", "quarter"), ("quarter", "used")),
        anchor_col=anchor_col,
        fallback_offset=-6,
    )
    sales_captured_col = resolve_column(
        headers,
        keyword_sets=(("sales", "captured"), ("captured", "db"), ("penetration",), ("capture", "%")),
        anchor_col=anchor_col,
        fallback_offset=-5,
    )
    growth_col = resolve_column(
        headers,
        keyword_sets=(("growth",), ("growth", "%")),
        anchor_col=anchor_col,
        fallback_offset=-4,
    )
    quarterly_sales_col = resolve_column(
        headers,
        keyword_sets=(("quarterly", "sales"), ("db", "sales"), ("units", "sold")),
        anchor_col=anchor_col,
        fallback_offset=-3,
    )
    reported_sales_col = resolve_column(
        headers,
        keyword_sets=(("reported", "sales"), ("actual", "sales")),
        anchor_col=anchor_col,
        fallback_offset=-2,
    )
    forecast_value_col = resolve_column(
        headers,
        keyword_sets=(
            ("estimated", "total", "sold"),
            ("estimate", "total", "sold"),
            ("forecast", "value"),
            ("tot", "fcst"),
        ),
        anchor_col=anchor_col,
        fallback_offset=-1,
    )
    forecast_max_col = anchor_col
    forecast_min_col = resolve_column(
        headers,
        keyword_sets=(("min",),),
        anchor_col=anchor_col,
        fallback_offset=1,
    )
    actual_value_col = reported_sales_col

    scratch_avg_pen_col = col_end + 2
    history_end_row = anchor_row - 1
    formula_rows: List[Tuple[int, int]] = []

    for index in range(10):
        row = anchor_row + 1 + index
        n_quarters_raw = read_cell(sheet, row, num_quarters_col)
        n_quarters = to_int(n_quarters_raw) or (index + 1)
        if n_quarters <= 0:
            n_quarters = index + 1
        start_row = history_end_row - n_quarters + 1
        if start_row < 1:
            continue

        formula = ""
        if quarterly_sales_col and reported_sales_col:
            formula = (
                f'=IFERROR(AVERAGE('
                f'R{start_row}C{quarterly_sales_col}:R{history_end_row}C{quarterly_sales_col}/'
                f'R{start_row}C{reported_sales_col}:R{history_end_row}C{reported_sales_col}'
                f'),"")'
            )
        elif sales_captured_col:
            formula = (
                f'=IFERROR(AVERAGE('
                f'R{start_row}C{sales_captured_col}:R{history_end_row}C{sales_captured_col}'
                f'),"")'
            )
        if formula:
            write_formula2(sheet.cells(row, scratch_avg_pen_col), formula)
            formula_rows.append((row, n_quarters))

    if formula_rows:
        wb.app.calculate()

    rows: List[Dict[str, Any]] = []
    for index in range(10):
        row = anchor_row + 1 + index
        n_quarters_raw = read_cell(sheet, row, num_quarters_col)
        n_quarters = to_int(n_quarters_raw) or (index + 1)
        if n_quarters <= 0:
            n_quarters = index + 1

        forecast_value = read_cell(sheet, row, forecast_value_col)
        forecast_max = read_cell(sheet, row, forecast_max_col)
        forecast_min = read_cell(sheet, row, forecast_min_col)
        quarterly_sales = read_cell(sheet, row, quarterly_sales_col)
        reported_sales = read_cell(sheet, row, reported_sales_col)
        sales_captured = read_cell(sheet, row, sales_captured_col)
        growth_rate = read_cell(sheet, row, growth_col)
        last_quarter_used = read_cell(sheet, row, last_quarter_col)
        actual_value = read_cell(sheet, row, actual_value_col)
        avg_penetration = read_cell(sheet, row, scratch_avg_pen_col)
        if avg_penetration in (None, ""):
            avg_penetration = sales_captured

        if forecast_value in (None, ""):
            q_sales_num = to_float(quarterly_sales)
            avg_pen_num = to_float(avg_penetration)
            if q_sales_num is not None and avg_pen_num not in (None, 0):
                forecast_value = q_sales_num / avg_pen_num

        max_num = to_float(forecast_max)
        min_num = to_float(forecast_min)
        range_width = (max_num - min_num) if max_num is not None and min_num is not None else None

        if (
            forecast_value in (None, "")
            and forecast_max in (None, "")
            and forecast_min in (None, "")
            and avg_penetration in (None, "")
        ):
            continue

        rows.append(
            {
                "model": meta.model,
                "ticker": meta.ticker,
                "model_period": meta.model_period,
                "model_date": meta.model_date,
                "method": "empirical",
                "parameter_name": "avg_penetration_pct",
                "parameter_value": avg_penetration,
                "num_quarters_used": n_quarters,
                "last_quarter_used": last_quarter_used,
                "forecast_value": forecast_value,
                "actual_value": actual_value,
                "forecast_max": forecast_max,
                "forecast_min": forecast_min,
                "range_width": range_width,
                "avg_penetration_pct": avg_penetration,
                "quarterly_sales": quarterly_sales,
                "reported_sales": reported_sales,
                "growth_rate_pct": growth_rate,
                "sales_captured_in_db_pct": sales_captured,
                "source_file": source_file,
            }
        )

    return rows


def process_regression_sheet(
    wb: xw.main.Book, sheet: xw.main.Sheet, meta: ModelMeta, source_file: str
) -> List[Dict[str, Any]]:
    anchor = find_anchor(sheet, target="max")
    if anchor is None:
        return []
    anchor_row, anchor_col = anchor
    headers, _, col_end = collect_headers(sheet, anchor_row, anchor_col)

    y_col = anchor_col - 7
    x_col = anchor_col - 11

    num_quarters_col = resolve_column(
        headers,
        keyword_sets=(("num", "quarter"), ("quarter", "used"), ("n", "quarter")),
        anchor_col=anchor_col,
        fallback_offset=-2,
    )
    forecast_value_col = resolve_column(
        headers,
        keyword_sets=(
            ("tot", "fcst", "w/o", "sa"),
            ("total", "forecast", "without", "sa"),
            ("forecast", "without", "sa"),
            ("forecast", "value"),
        ),
        anchor_col=anchor_col,
        fallback_offset=-1,
    )
    actual_value_col = find_column(headers, keyword_sets=(("actual",), ("reported", "sales")))
    forecast_max_col = anchor_col
    forecast_min_col = resolve_column(
        headers,
        keyword_sets=(("min",),),
        anchor_col=anchor_col,
        fallback_offset=1,
    )

    scratch_intercept_col = col_end + 2
    scratch_slope_col = col_end + 3
    history_end_row = anchor_row - 1
    formula_rows: List[Tuple[int, int]] = []

    for index in range(10):
        row = anchor_row + 1 + index
        n_quarters_raw = read_cell(sheet, row, num_quarters_col)
        n_quarters = to_int(n_quarters_raw) or (index + 1)
        if n_quarters <= 0:
            n_quarters = index + 1

        start_row = history_end_row - n_quarters + 1
        if start_row < 1:
            continue

        intercept_formula = (
            f'=IFERROR(INTERCEPT(R{start_row}C{y_col}:R{history_end_row}C{y_col},'
            f'R{start_row}C{x_col}:R{history_end_row}C{x_col}),"")'
        )
        slope_formula = (
            f'=IFERROR(SLOPE(R{start_row}C{y_col}:R{history_end_row}C{y_col},'
            f'R{start_row}C{x_col}:R{history_end_row}C{x_col}),"")'
        )

        write_formula2(sheet.cells(row, scratch_intercept_col), intercept_formula)
        write_formula2(sheet.cells(row, scratch_slope_col), slope_formula)
        formula_rows.append((row, n_quarters))

    if formula_rows:
        wb.app.calculate()

    rows: List[Dict[str, Any]] = []
    previous_signature: Optional[Tuple[Any, ...]] = None

    for index in range(10):
        row = anchor_row + 1 + index
        n_quarters_raw = read_cell(sheet, row, num_quarters_col)
        n_quarters = to_int(n_quarters_raw) or (index + 1)
        if n_quarters <= 0:
            n_quarters = index + 1

        intercept = read_cell(sheet, row, scratch_intercept_col)
        slope = read_cell(sheet, row, scratch_slope_col)
        forecast_value = read_cell(sheet, row, forecast_value_col)
        actual_value = read_cell(sheet, row, actual_value_col)
        forecast_max = read_cell(sheet, row, forecast_max_col)
        forecast_min = read_cell(sheet, row, forecast_min_col)

        max_num = to_float(forecast_max)
        min_num = to_float(forecast_min)
        range_width = (max_num - min_num) if max_num is not None and min_num is not None else None

        if (
            intercept in (None, "")
            and slope in (None, "")
            and forecast_value in (None, "")
            and forecast_max in (None, "")
            and forecast_min in (None, "")
        ):
            continue

        signature = numeric_signature(
            (n_quarters, intercept, slope, forecast_value, forecast_max, forecast_min)
        )
        if signature == previous_signature:
            continue
        previous_signature = signature

        rows.append(
            {
                "model": meta.model,
                "ticker": meta.ticker,
                "model_period": meta.model_period,
                "model_date": meta.model_date,
                "method": "regression",
                "parameter_name": "num_quarters_used",
                "parameter_value": n_quarters,
                "num_quarters_used": n_quarters,
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


def close_workbook_without_saving(wb: xw.main.Book) -> None:
    try:
        wb.close(save=False)
        return
    except TypeError:
        pass
    except Exception:
        pass

    try:
        wb.api.Close(False)
        return
    except Exception:
        pass

    try:
        wb.close()
    except Exception:
        pass


def write_rows_with_formatting(
    ws: Any, columns: Sequence[str], rows: Sequence[Dict[str, Any]]
) -> None:
    ws.append(list(columns))
    for row in rows:
        ws.append([row.get(column) for column in columns])

    for header_cell in ws[1]:
        header_cell.font = Font(bold=True)

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions

    for col_idx, name in enumerate(columns, start=1):
        max_width = len(name)
        for row_idx in range(2, ws.max_row + 1):
            value = ws.cell(row=row_idx, column=col_idx).value
            if value is None:
                continue
            max_width = max(max_width, len(str(value)))
        ws.column_dimensions[get_column_letter(col_idx)].width = min(max(max_width + 2, 10), 48)


def write_output(
    output_path: Path,
    empirical_rows: Sequence[Dict[str, Any]],
    regression_rows: Sequence[Dict[str, Any]],
) -> None:
    workbook = Workbook()
    empirical_ws = workbook.active
    empirical_ws.title = "empirical_candidates"
    regression_ws = workbook.create_sheet("regression_candidates")

    write_rows_with_formatting(empirical_ws, EMPIRICAL_COLUMNS, empirical_rows)
    write_rows_with_formatting(regression_ws, REGRESSION_COLUMNS, regression_rows)
    workbook.save(output_path)


def iter_candidate_files(folder: Path) -> Iterable[Path]:
    for path in sorted(folder.iterdir(), key=lambda p: p.name.lower()):
        if not path.is_file():
            continue
        if path.name.startswith("~"):
            print(f"Skipped {path.name}: temporary file.")
            continue
        if path.suffix.lower() != ".xlsx":
            print(f"Skipped {path.name}: not an .xlsx file.")
            continue
        yield path


def main() -> None:
    input_path = Path(input_dir)
    output_path_dir = Path(output_dir)

    if not input_path.exists() or not input_path.is_dir():
        print(f"Input directory does not exist: {input_path}")
        return

    output_path_dir.mkdir(parents=True, exist_ok=True)
    output_path = get_output_path(input_path, output_path_dir)

    empirical_rows: List[Dict[str, Any]] = []
    regression_rows: List[Dict[str, Any]] = []
    processed_files = 0

    app = xw.App(visible=False, add_book=False)
    app.display_alerts = False
    app.screen_updating = False
    app.calculation = "manual"

    try:
        for file_path in iter_candidate_files(input_path):
            print(f"Processing {file_path.name}")
            wb: Optional[xw.main.Book] = None
            try:
                wb = app.books.open(str(file_path), update_links=False)
                meta = parse_model_meta(file_path)

                empirical_sheet = get_sheet(wb, "Empirical Model")
                if empirical_sheet is None:
                    print(f"Skipped empirical extraction for {file_path.name}: missing 'Empirical Model' sheet.")
                else:
                    empirical_rows.extend(
                        process_empirical_sheet(wb, empirical_sheet, meta, file_path.name)
                    )

                regression_sheet = get_sheet(wb, "Regression Model")
                if regression_sheet is None:
                    print(f"Skipped regression extraction for {file_path.name}: missing 'Regression Model' sheet.")
                else:
                    regression_rows.extend(
                        process_regression_sheet(wb, regression_sheet, meta, file_path.name)
                    )

                processed_files += 1
            except Exception as exc:
                print(f"Skipped {file_path.name}: {exc}")
            finally:
                if wb is not None:
                    close_workbook_without_saving(wb)
    finally:
        app.quit()

    write_output(output_path, empirical_rows, regression_rows)

    print(f"Output path: {output_path}")
    print(f"Number of files processed: {processed_files}")
    print(f"Number of empirical rows: {len(empirical_rows)}")
    print(f"Number of regression rows: {len(regression_rows)}")


if __name__ == "__main__":
    main()
