#!/usr/bin/env python3
from __future__ import annotations

import re
from dataclasses import dataclass
from datetime import date
from pathlib import Path
from typing import Any, Dict, List, Optional, Sequence, Tuple

import xlwings as xw
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

# Configure these two directories before running.
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

MONTH_TO_NUMBER = {
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

PHASE_TO_DAY = {
    "early": 5,
    "mid": 15,
    "late": 25,
}

FILE_PERIOD_RE = re.compile(r"^(Early|Mid|Late)([A-Za-z]{3,9})(\d{4})$", flags=re.IGNORECASE)


@dataclass
class FileMetadata:
    model: str
    ticker: str
    model_period: str
    model_date: str


@dataclass
class SheetScan:
    row0: int
    col0: int
    values: List[List[Any]]

    @property
    def height(self) -> int:
        return len(self.values)

    @property
    def width(self) -> int:
        return len(self.values[0]) if self.values else 0

    @property
    def last_row(self) -> int:
        return self.row0 + self.height - 1

    @property
    def last_col(self) -> int:
        return self.col0 + self.width - 1


def normalize_text(value: Any) -> str:
    if value is None:
        return ""
    text = str(value).strip().lower()
    return re.sub(r"[^a-z0-9]+", " ", text).strip()


def to_matrix(value: Any) -> List[List[Any]]:
    if isinstance(value, tuple):
        value = list(value)

    if value is None:
        return []

    if isinstance(value, list):
        if not value:
            return []
        if isinstance(value[0], tuple):
            value = [list(row) for row in value]
        if isinstance(value[0], list):
            rows = [list(row) if isinstance(row, list) else [row] for row in value]
        else:
            rows = [list(value)]
    else:
        rows = [[value]]

    width = max(len(row) for row in rows)
    padded_rows: List[List[Any]] = []
    for row in rows:
        padded_rows.append(row + [None] * (width - len(row)))
    return padded_rows


def scan_sheet(sheet: xw.Sheet) -> Optional[SheetScan]:
    used_range = sheet.used_range
    matrix = to_matrix(used_range.value)
    if not matrix:
        return None
    return SheetScan(row0=used_range.row, col0=used_range.column, values=matrix)


def get_scan_value(scan: SheetScan, row: int, col: int) -> Any:
    row_idx = row - scan.row0
    col_idx = col - scan.col0
    if row_idx < 0 or col_idx < 0 or row_idx >= scan.height or col_idx >= scan.width:
        return None
    return scan.values[row_idx][col_idx]


def find_anchor(scan: SheetScan, target: str = "max") -> Optional[Tuple[int, int]]:
    target_normalized = normalize_text(target)
    for row_idx, row_values in enumerate(scan.values):
        for col_idx, cell_value in enumerate(row_values):
            if normalize_text(cell_value) == target_normalized:
                return scan.row0 + row_idx, scan.col0 + col_idx
    return None


def get_header_entries(scan: SheetScan, header_row: int) -> List[Tuple[int, str]]:
    entries: List[Tuple[int, str]] = []
    for col in range(scan.col0, scan.last_col + 1):
        normalized = normalize_text(get_scan_value(scan, header_row, col))
        if normalized:
            entries.append((col, normalized))
    return entries


def find_col_by_tokens(
    headers: Sequence[Tuple[int, str]], token_groups: Sequence[Sequence[str]]
) -> Optional[int]:
    normalized_groups = [[normalize_text(token) for token in group] for group in token_groups]
    for group in normalized_groups:
        for col, header in headers:
            if all(re.search(rf"\b{re.escape(token)}\b", header) for token in group):
                return col
    return None


def as_float(value: Any) -> Optional[float]:
    if value is None or isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        return float(value)

    text = str(value).strip()
    if not text:
        return None
    text = text.replace(",", "")
    is_percent = text.endswith("%")
    if is_percent:
        text = text[:-1]
    try:
        number = float(text)
    except ValueError:
        return None
    return number / 100.0 if is_percent else number


def as_int(value: Any) -> Optional[int]:
    number = as_float(value)
    if number is None:
        return None
    return int(round(number))


def rounded_number(value: Any) -> Optional[float]:
    number = as_float(value)
    if number is None:
        return None
    return round(number, 10)


def parse_file_metadata(file_name: str) -> FileMetadata:
    stem = Path(file_name).stem
    parts = [part.strip() for part in stem.split(" - ")]

    ticker = ""
    period_token = ""
    if len(parts) >= 3:
        ticker = parts[-2].upper()
        period_token = parts[-1]
    elif len(parts) == 2:
        ticker = parts[-1].upper()
    elif len(parts) == 1:
        ticker = parts[0].upper()

    period_token = re.sub(r"_send$", "", period_token, flags=re.IGNORECASE).strip()
    model_period = period_token if period_token else "unknown_period"
    model_date = ""

    period_match = FILE_PERIOD_RE.match(period_token)
    if period_match:
        phase = period_match.group(1).title()
        month_key = period_match.group(2)[:3].lower()
        year = period_match.group(3)
        month_number = MONTH_TO_NUMBER.get(month_key)
        if month_number:
            model_period = f"{phase}{month_key.title()}_{year}"
            model_date = date(
                int(year), month_number, PHASE_TO_DAY[phase.lower()]
            ).isoformat()

    if not ticker:
        ticker = "UNKNOWN"
    model = f"{ticker}_{model_period}"

    return FileMetadata(
        model=model,
        ticker=ticker,
        model_period=model_period,
        model_date=model_date,
    )


def next_output_path(input_path: Path, target_output_dir: Path) -> Path:
    target_output_dir.mkdir(parents=True, exist_ok=True)
    base_name = f"{input_path.name}_PARAM"
    candidate = target_output_dir / f"{base_name}.xlsx"
    if not candidate.exists():
        return candidate

    suffix = 1
    while True:
        candidate = target_output_dir / f"{base_name}.{suffix}.xlsx"
        if not candidate.exists():
            return candidate
        suffix += 1


def get_sheet_by_name(workbook: xw.Book, target_name: str) -> Optional[xw.Sheet]:
    target_normalized = target_name.strip().lower()
    for sheet in workbook.sheets:
        if sheet.name.strip().lower() == target_normalized:
            return sheet
    return None


def close_workbook_safely(workbook: xw.Book) -> None:
    try:
        workbook.close(save=False)
        return
    except TypeError:
        pass
    except Exception:
        pass

    for closer in (
        lambda: workbook.close(False),
        lambda: workbook.api.Close(SaveChanges=False),
        lambda: workbook.api.Close(False),
    ):
        try:
            closer()
            return
        except Exception:
            continue


def ensure_formula_matrix(formulas: Any) -> List[List[str]]:
    if isinstance(formulas, str):
        return [[formulas]]
    if isinstance(formulas, list):
        if not formulas:
            return []
        if isinstance(formulas[0], list):
            return formulas
        return [[str(item)] for item in formulas]
    return [[str(formulas)]]


def write_formula2_r1c1(target_range: xw.Range, formulas: Any) -> None:
    formula_matrix = ensure_formula_matrix(formulas)
    if not formula_matrix:
        return

    assign_value: Any
    if len(formula_matrix) == 1 and len(formula_matrix[0]) == 1:
        assign_value = formula_matrix[0][0]
    else:
        assign_value = formula_matrix

    try:
        target_range.formula2 = assign_value
        return
    except Exception:
        pass

    range_api = target_range.api
    for row_idx, row_values in enumerate(formula_matrix, start=1):
        for col_idx, formula in enumerate(row_values, start=1):
            if not formula:
                continue
            cell_api = range_api.Cells(row_idx, col_idx)
            try:
                cell_api.Formula2R1C1 = formula
            except Exception:
                cell_api.FormulaR1C1 = formula


def build_quarter_row_lookup(scan: SheetScan, quarter_col: Optional[int], start_row: int) -> Dict[int, int]:
    row_lookup: Dict[int, int] = {}
    if quarter_col is None:
        return row_lookup

    max_scan_row = min(scan.last_row, start_row + 3 * N_QUARTERS)
    for row in range(start_row, max_scan_row + 1):
        quarter_value = as_int(get_scan_value(scan, row, quarter_col))
        if quarter_value is None:
            continue
        if 1 <= quarter_value <= N_QUARTERS and quarter_value not in row_lookup:
            row_lookup[quarter_value] = row
    return row_lookup


def collect_numeric_rows(
    scan: SheetScan, main_col: int, start_row: int, paired_col: Optional[int] = None
) -> List[int]:
    rows: List[int] = []
    if main_col < scan.col0 or main_col > scan.last_col:
        return rows

    for row in range(start_row, scan.last_row + 1):
        main_value = as_float(get_scan_value(scan, row, main_col))
        if main_value is None:
            continue
        if paired_col is not None:
            paired_value = as_float(get_scan_value(scan, row, paired_col))
            if paired_value is None:
                continue
        rows.append(row)
    return rows


def col_with_fallback(scan: SheetScan, preferred: Optional[int], fallback: Optional[int]) -> Optional[int]:
    col = preferred if preferred is not None else fallback
    if col is None:
        return None
    if scan.col0 <= col <= scan.last_col:
        return col
    return None


def extract_empirical_candidates(
    workbook: xw.Book, metadata: FileMetadata, source_file: str
) -> List[Dict[str, Any]]:
    sheet = get_sheet_by_name(workbook, "Empirical Model")
    if sheet is None:
        print(f"Skipped empirical for {source_file}: sheet 'Empirical Model' not found")
        return []

    scan = scan_sheet(sheet)
    if scan is None:
        print(f"Skipped empirical for {source_file}: sheet is empty")
        return []

    anchor = find_anchor(scan, "max")
    if anchor is None:
        print(f"Skipped empirical for {source_file}: 'max' anchor not found")
        return []

    anchor_row, anchor_col = anchor
    headers = get_header_entries(scan, anchor_row)

    num_quarters_col = col_with_fallback(
        scan,
        find_col_by_tokens(headers, (("num", "quarter"), ("quarters", "used"))),
        anchor_col - 8,
    )
    last_quarter_col = col_with_fallback(
        scan, find_col_by_tokens(headers, (("last", "quarter"),)), anchor_col - 7
    )
    forecast_col = col_with_fallback(
        scan,
        find_col_by_tokens(
            headers,
            (
                ("estimated", "total", "sold"),
                ("est", "total", "sold"),
                ("forecast", "total"),
                ("forecast", "value"),
            ),
        ),
        anchor_col - 6,
    )
    actual_col = col_with_fallback(
        scan,
        find_col_by_tokens(
            headers, (("reported", "sales"), ("actual", "sales"), ("actual", "value"))
        ),
        anchor_col - 5,
    )
    avg_penetration_col = col_with_fallback(
        scan,
        find_col_by_tokens(
            headers,
            (
                ("avg", "penetration"),
                ("average", "penetration"),
                ("penetration", "pct"),
            ),
        ),
        anchor_col - 4,
    )
    penetration_series_col = col_with_fallback(
        scan,
        find_col_by_tokens(headers, (("penetration", "pct"), ("penetration",))),
        avg_penetration_col,
    )
    quarterly_sales_col = col_with_fallback(
        scan,
        find_col_by_tokens(
            headers, (("quarterly", "sales"), ("qtr", "sales"), ("quarter", "sales"))
        ),
        anchor_col - 3,
    )
    reported_sales_col = col_with_fallback(
        scan, find_col_by_tokens(headers, (("reported", "sales"),)), anchor_col - 2
    )
    growth_rate_col = col_with_fallback(
        scan,
        find_col_by_tokens(headers, (("growth", "rate"), ("growth", "pct"))),
        anchor_col - 1,
    )
    captured_col = col_with_fallback(
        scan,
        find_col_by_tokens(
            headers,
            (
                ("captured", "db"),
                ("captured", "database"),
                ("sales", "captured"),
            ),
        ),
        anchor_col + 2,
    )

    max_col = anchor_col
    min_col = col_with_fallback(scan, find_col_by_tokens(headers, (("min",),)), anchor_col + 1)
    data_start_row = anchor_row + 1

    row_lookup = build_quarter_row_lookup(scan, num_quarters_col, data_start_row)

    avg_penetration_by_quarter: Dict[int, Optional[float]] = {}
    if penetration_series_col is not None:
        penetration_rows = collect_numeric_rows(scan, penetration_series_col, data_start_row)
        if penetration_rows:
            n_formula_rows = min(N_QUARTERS, len(penetration_rows))
            helper_row = scan.last_row + 2
            helper_col = scan.last_col + 2
            formulas: List[List[str]] = []
            for n in range(1, n_formula_rows + 1):
                start_row = penetration_rows[-n]
                end_row = penetration_rows[-1]
                formulas.append(
                    [f"=AVERAGE(R{start_row}C{penetration_series_col}:R{end_row}C{penetration_series_col})"]
                )

            helper_range = sheet.range(
                (helper_row, helper_col), (helper_row + n_formula_rows - 1, helper_col)
            )
            write_formula2_r1c1(helper_range, formulas)
            workbook.app.calculate()
            helper_values = to_matrix(helper_range.value)
            for idx in range(1, n_formula_rows + 1):
                value = helper_values[idx - 1][0] if idx - 1 < len(helper_values) else None
                avg_penetration_by_quarter[idx] = as_float(value)

    rows: List[Dict[str, Any]] = []
    empty_streak = 0
    for n in range(1, N_QUARTERS + 1):
        row = row_lookup.get(n, data_start_row + (n - 1))

        num_quarters_used = as_int(get_scan_value(scan, row, num_quarters_col)) if num_quarters_col else n
        if num_quarters_used is None:
            num_quarters_used = n

        last_quarter_used = get_scan_value(scan, row, last_quarter_col) if last_quarter_col else None
        forecast_value = as_float(get_scan_value(scan, row, forecast_col)) if forecast_col else None
        actual_value = as_float(get_scan_value(scan, row, actual_col)) if actual_col else None
        forecast_max = as_float(get_scan_value(scan, row, max_col))
        forecast_min = as_float(get_scan_value(scan, row, min_col))

        avg_penetration_pct = avg_penetration_by_quarter.get(n)
        if avg_penetration_pct is None and avg_penetration_col is not None:
            avg_penetration_pct = as_float(get_scan_value(scan, row, avg_penetration_col))

        quarterly_sales = as_float(get_scan_value(scan, row, quarterly_sales_col)) if quarterly_sales_col else None
        reported_sales = as_float(get_scan_value(scan, row, reported_sales_col)) if reported_sales_col else None
        growth_rate_pct = as_float(get_scan_value(scan, row, growth_rate_col)) if growth_rate_col else None
        sales_captured_in_db_pct = as_float(get_scan_value(scan, row, captured_col)) if captured_col else None

        if reported_sales is None:
            reported_sales = actual_value
        if actual_value is None:
            actual_value = reported_sales
        if (
            forecast_value is None
            and quarterly_sales is not None
            and avg_penetration_pct is not None
            and avg_penetration_pct != 0
        ):
            forecast_value = quarterly_sales / avg_penetration_pct

        range_width = (
            forecast_max - forecast_min
            if forecast_max is not None and forecast_min is not None
            else None
        )

        has_signal = any(
            value is not None
            for value in (
                forecast_value,
                actual_value,
                forecast_max,
                forecast_min,
                avg_penetration_pct,
                quarterly_sales,
                reported_sales,
                growth_rate_pct,
                sales_captured_in_db_pct,
            )
        )
        if not has_signal:
            empty_streak += 1
            if empty_streak >= 3 and n > 3:
                break
            continue
        empty_streak = 0

        rows.append(
            {
                "model": metadata.model,
                "ticker": metadata.ticker,
                "model_period": metadata.model_period,
                "model_date": metadata.model_date,
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

    return rows


def extract_regression_candidates(
    workbook: xw.Book, metadata: FileMetadata, source_file: str
) -> List[Dict[str, Any]]:
    sheet = get_sheet_by_name(workbook, "Regression Model")
    if sheet is None:
        print(f"Skipped regression for {source_file}: sheet 'Regression Model' not found")
        return []

    scan = scan_sheet(sheet)
    if scan is None:
        print(f"Skipped regression for {source_file}: sheet is empty")
        return []

    anchor = find_anchor(scan, "max")
    if anchor is None:
        print(f"Skipped regression for {source_file}: 'max' anchor not found")
        return []

    anchor_row, anchor_col = anchor
    headers = get_header_entries(scan, anchor_row)

    num_quarters_col = col_with_fallback(
        scan,
        find_col_by_tokens(headers, (("num", "quarter"), ("quarters", "used"))),
        anchor_col - 9,
    )
    forecast_col = col_with_fallback(
        scan,
        find_col_by_tokens(
            headers,
            (
                ("tot", "fcst", "sa"),
                ("tot", "forecast", "sa"),
                ("forecast", "without", "sa"),
            ),
        ),
        anchor_col - 1,
    )
    actual_col = col_with_fallback(
        scan,
        find_col_by_tokens(
            headers, (("actual", "sales"), ("reported", "sales"), ("actual", "value"))
        ),
        anchor_col - 2,
    )
    max_col = anchor_col
    min_col = col_with_fallback(scan, find_col_by_tokens(headers, (("min",),)), anchor_col + 1)

    y_col = anchor_col - 7
    x_col = anchor_col - 11
    data_start_row = anchor_row + 1

    row_lookup = build_quarter_row_lookup(scan, num_quarters_col, data_start_row)

    coeff_by_quarter: Dict[int, Tuple[Optional[float], Optional[float]]] = {}
    valid_rows = collect_numeric_rows(scan, x_col, data_start_row, paired_col=y_col)
    if valid_rows:
        n_formula_rows = min(N_QUARTERS, len(valid_rows))
        helper_row = scan.last_row + 2
        helper_col = scan.last_col + 3
        formulas: List[List[str]] = []
        for n in range(1, n_formula_rows + 1):
            start_row = valid_rows[-n]
            end_row = valid_rows[-1]
            y_range = f"R{start_row}C{y_col}:R{end_row}C{y_col}"
            x_range = f"R{start_row}C{x_col}:R{end_row}C{x_col}"
            formulas.append(
                [
                    f"=INTERCEPT({y_range},{x_range})",
                    f"=SLOPE({y_range},{x_range})",
                ]
            )

        helper_range = sheet.range(
            (helper_row, helper_col),
            (helper_row + n_formula_rows - 1, helper_col + 1),
        )
        write_formula2_r1c1(helper_range, formulas)
        workbook.app.calculate()
        helper_values = to_matrix(helper_range.value)
        for idx in range(1, n_formula_rows + 1):
            row_values = helper_values[idx - 1] if idx - 1 < len(helper_values) else []
            intercept = as_float(row_values[0]) if len(row_values) > 0 else None
            slope = as_float(row_values[1]) if len(row_values) > 1 else None
            coeff_by_quarter[idx] = (intercept, slope)

    rows: List[Dict[str, Any]] = []
    previous_key: Optional[Tuple[Any, ...]] = None
    empty_streak = 0
    for n in range(1, N_QUARTERS + 1):
        row = row_lookup.get(n, data_start_row + (n - 1))

        num_quarters_used = as_int(get_scan_value(scan, row, num_quarters_col)) if num_quarters_col else n
        if num_quarters_used is None:
            num_quarters_used = n

        forecast_value = as_float(get_scan_value(scan, row, forecast_col)) if forecast_col else None
        actual_value = as_float(get_scan_value(scan, row, actual_col)) if actual_col else None
        forecast_max = as_float(get_scan_value(scan, row, max_col))
        forecast_min = as_float(get_scan_value(scan, row, min_col))
        intercept, slope = coeff_by_quarter.get(n, (None, None))

        range_width = (
            forecast_max - forecast_min
            if forecast_max is not None and forecast_min is not None
            else None
        )

        has_signal = any(
            value is not None
            for value in (
                forecast_value,
                actual_value,
                forecast_max,
                forecast_min,
                intercept,
                slope,
            )
        )
        if not has_signal:
            empty_streak += 1
            if empty_streak >= 3 and n > 3:
                break
            continue
        empty_streak = 0

        dedupe_key = (
            num_quarters_used,
            rounded_number(forecast_value),
            rounded_number(forecast_max),
            rounded_number(forecast_min),
            rounded_number(intercept),
            rounded_number(slope),
        )
        if dedupe_key == previous_key:
            continue
        previous_key = dedupe_key

        rows.append(
            {
                "model": metadata.model,
                "ticker": metadata.ticker,
                "model_period": metadata.model_period,
                "model_date": metadata.model_date,
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

    return rows


def write_sheet(worksheet: Any, columns: Sequence[str], rows: Sequence[Dict[str, Any]]) -> None:
    worksheet.append(list(columns))
    for row in rows:
        worksheet.append([row.get(col) for col in columns])

    header_font = Font(bold=True)
    for cell in worksheet[1]:
        cell.font = header_font

    worksheet.freeze_panes = "A2"
    if worksheet.max_row >= 1 and worksheet.max_column >= 1:
        worksheet.auto_filter.ref = worksheet.dimensions

    for col_idx, column_name in enumerate(columns, start=1):
        width = len(column_name)
        for row_idx in range(2, worksheet.max_row + 1):
            value = worksheet.cell(row=row_idx, column=col_idx).value
            if value is None:
                continue
            width = max(width, len(str(value)))
        worksheet.column_dimensions[get_column_letter(col_idx)].width = min(width + 2, 48)


def write_output_workbook(
    output_path: Path,
    empirical_rows: Sequence[Dict[str, Any]],
    regression_rows: Sequence[Dict[str, Any]],
) -> None:
    workbook = Workbook()
    default_sheet = workbook.active
    workbook.remove(default_sheet)

    empirical_sheet = workbook.create_sheet("empirical_candidates")
    regression_sheet = workbook.create_sheet("regression_candidates")
    write_sheet(empirical_sheet, EMPIRICAL_COLUMNS, empirical_rows)
    write_sheet(regression_sheet, REGRESSION_COLUMNS, regression_rows)

    workbook.save(output_path)


def main() -> None:
    input_path = Path(input_dir).expanduser()
    output_path = Path(output_dir).expanduser()

    if not input_path.exists():
        print(f"Input directory not found: {input_path}")
        return
    if not input_path.is_dir():
        print(f"Input path is not a directory: {input_path}")
        return

    files = sorted(input_path.iterdir(), key=lambda file_path: file_path.name.lower())
    empirical_rows: List[Dict[str, Any]] = []
    regression_rows: List[Dict[str, Any]] = []
    processed_count = 0

    try:
        app = xw.App(visible=False, add_book=False)
    except Exception as exc:
        print(f"Unable to start hidden Excel app: {exc}")
        return

    app.display_alerts = False
    app.screen_updating = False
    try:
        app.enable_events = False
    except Exception:
        pass
    try:
        app.calculation = "manual"
    except Exception:
        pass

    try:
        for file_path in files:
            if not file_path.is_file():
                print(f"Skipped {file_path.name}: not a file")
                continue
            if file_path.name.startswith("~"):
                print(f"Skipped {file_path.name}: temporary file")
                continue
            if file_path.suffix.lower() != ".xlsx":
                print(f"Skipped {file_path.name}: not an .xlsx file")
                continue

            print(f"Processing {file_path.name}")
            metadata = parse_file_metadata(file_path.name)

            try:
                workbook = app.books.open(str(file_path), update_links=False)
            except Exception as exc:
                print(f"Skipped {file_path.name}: failed to open ({exc})")
                continue

            try:
                empirical_rows.extend(
                    extract_empirical_candidates(workbook, metadata, file_path.name)
                )
                regression_rows.extend(
                    extract_regression_candidates(workbook, metadata, file_path.name)
                )
                processed_count += 1
            except Exception as exc:
                print(f"Skipped {file_path.name}: processing error ({exc})")
            finally:
                close_workbook_safely(workbook)
    finally:
        try:
            app.quit()
        except Exception:
            pass

    output_file = next_output_path(input_path, output_path)
    write_output_workbook(output_file, empirical_rows, regression_rows)

    print(f"Output path: {output_file}")
    print(f"Number of files processed: {processed_count}")
    print(f"Number of empirical rows: {len(empirical_rows)}")
    print(f"Number of regression rows: {len(regression_rows)}")


if __name__ == "__main__":
    main()
