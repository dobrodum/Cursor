#!/usr/bin/env python3
from __future__ import annotations

from dataclasses import dataclass
from datetime import date
from pathlib import Path
import re
from typing import Any, Dict, List, Optional, Sequence, Tuple

import xlwings as xw
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

# -----------------------------
# User-configurable directories
# -----------------------------
input_dir = r"/path/to/input"
output_dir = r"/path/to/output"


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

WINDOW_DAY = {
    "early": 5,
    "mid": 15,
    "late": 25,
}

MONTH_TO_NUM = {
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

NON_ALNUM_RE = re.compile(r"[^a-z0-9]+")
FILE_LABEL_RE = re.compile(
    r"-\s*(?P<ticker>[A-Za-z0-9]+)\s*-\s*(?P<window>Early|Mid|Late)"
    r"(?P<month>[A-Za-z]{3,9})(?P<year>\d{4})",
    flags=re.IGNORECASE,
)
MAX_RE = re.compile(r"_PARAM(?:\.\d+)?$", flags=re.IGNORECASE)

# Fallback offsets from the "max" anchor cell when headers are not found.
EMPIRICAL_FALLBACK_OFFSETS = {
    "num_quarters_used": -9,
    "last_quarter_used": -8,
    "quarterly_sales": -7,
    "reported_sales": -6,
    "growth_rate_pct": -5,
    "sales_captured_in_db_pct": -4,
    "avg_penetration_pct": -3,
    "forecast_value": -2,
    "actual_value": -1,
    "forecast_max": 0,
    "forecast_min": 1,
}

REGRESSION_FALLBACK_OFFSETS = {
    "num_quarters_used": -4,
    "forecast_value": -1,
    "forecast_max": 0,
    "forecast_min": 1,
}


@dataclass
class FileLabel:
    model: str
    ticker: str
    model_period: str
    model_date: str


@dataclass
class SheetCache:
    sheet: xw.Sheet
    start_row: int
    start_col: int
    values: List[List[Any]]

    @classmethod
    def from_sheet(cls, sheet: xw.Sheet) -> "SheetCache":
        used = sheet.used_range
        values = ensure_2d(used.value)
        return cls(sheet=sheet, start_row=used.row, start_col=used.column, values=values)

    @property
    def row_count(self) -> int:
        return len(self.values)

    @property
    def col_count(self) -> int:
        if not self.values:
            return 0
        return max(len(row) for row in self.values)

    @property
    def end_row(self) -> int:
        return self.start_row + self.row_count - 1

    @property
    def end_col(self) -> int:
        return self.start_col + self.col_count - 1

    def get(self, row: int, col: int) -> Any:
        if row < self.start_row or col < self.start_col or row > self.end_row or col > self.end_col:
            return None
        row_idx = row - self.start_row
        col_idx = col - self.start_col
        row_values = self.values[row_idx]
        if col_idx >= len(row_values):
            return None
        return row_values[col_idx]

    def row_values(self, row: int) -> List[Tuple[int, Any]]:
        if row < self.start_row or row > self.end_row:
            return []
        row_values = self.values[row - self.start_row]
        return [(self.start_col + idx, row_values[idx]) for idx in range(len(row_values))]


def ensure_2d(value: Any) -> List[List[Any]]:
    if value is None:
        return []
    if not isinstance(value, list):
        return [[value]]
    if not value:
        return []
    if isinstance(value[0], list):
        return value
    return [value]


def normalize_text(value: Any) -> str:
    if value is None:
        return ""
    text = str(value).strip().lower()
    if not text:
        return ""
    return NON_ALNUM_RE.sub(" ", text).strip()


def clean_cell_value(value: Any) -> Any:
    if isinstance(value, str):
        trimmed = value.strip()
        if trimmed in {"", "#N/A", "#VALUE!", "#DIV/0!", "#REF!"}:
            return None
        return trimmed
    return value


def coerce_float(value: Any) -> Optional[float]:
    value = clean_cell_value(value)
    if value is None or isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        return float(value)
    if isinstance(value, str):
        cleaned = value.replace(",", "").strip()
        if not cleaned:
            return None
        if cleaned.endswith("%"):
            base = cleaned[:-1].strip()
            if not base:
                return None
            try:
                return float(base) / 100.0
            except ValueError:
                return None
        try:
            return float(cleaned)
        except ValueError:
            return None
    return None


def coerce_int(value: Any) -> Optional[int]:
    number = coerce_float(value)
    if number is None:
        return None
    return int(round(number))


def safe_diff(a: Any, b: Any) -> Optional[float]:
    av = coerce_float(a)
    bv = coerce_float(b)
    if av is None or bv is None:
        return None
    return av - bv


def parse_file_label(filename: str) -> FileLabel:
    stem = Path(filename).stem
    match = FILE_LABEL_RE.search(stem)
    if not match:
        ticker_guess = "UNKNOWN"
        parts = [part.strip() for part in stem.split("-")]
        if len(parts) >= 2 and parts[1]:
            ticker_guess = parts[1].split()[0].upper()
        model_period = "Unknown_0000"
        return FileLabel(
            model=f"{ticker_guess}_{model_period}",
            ticker=ticker_guess,
            model_period=model_period,
            model_date="",
        )

    ticker = match.group("ticker").upper()
    window = match.group("window").title()
    month_token = match.group("month")[:3].title()
    year = int(match.group("year"))
    month_num = MONTH_TO_NUM.get(month_token)

    model_period = f"{window}{month_token}_{year}"
    model_date = ""
    if month_num is not None:
        model_date = date(year, month_num, WINDOW_DAY[window.lower()]).isoformat()

    return FileLabel(
        model=f"{ticker}_{model_period}",
        ticker=ticker,
        model_period=model_period,
        model_date=model_date,
    )


def next_output_path(input_path: Path, output_path_root: Path) -> Path:
    base_name = input_path.resolve().name
    candidate = output_path_root / f"{base_name}_PARAM.xlsx"
    if not candidate.exists():
        return candidate

    suffix = 1
    while True:
        candidate = output_path_root / f"{base_name}_PARAM.{suffix}.xlsx"
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

    close_attempts = [
        lambda: wb.api.Close(False),
        lambda: wb.api.Close(SaveChanges=False),
        lambda: wb.api.close(saving=False),
    ]
    for close_fn in close_attempts:
        try:
            close_fn()
            return
        except Exception:
            continue


def find_anchor_max(cache: SheetCache) -> Optional[Tuple[int, int]]:
    candidates: List[Tuple[int, int, int]] = []
    for r_idx, row_values in enumerate(cache.values):
        for c_idx, value in enumerate(row_values):
            if normalize_text(value) != "max":
                continue
            row = cache.start_row + r_idx
            col = cache.start_col + c_idx
            score = 0
            if normalize_text(cache.get(row, col + 1)) == "min":
                score += 2
            if normalize_text(cache.get(row, col - 1)) == "min":
                score += 1
            candidates.append((score, row, col))

    if not candidates:
        return None
    candidates.sort(key=lambda item: (-item[0], item[1], item[2]))
    _, row, col = candidates[0]
    return row, col


def choose_header_row(cache: SheetCache, anchor_row: int, terms: Sequence[str]) -> int:
    candidate_rows = [anchor_row - 1, anchor_row, anchor_row + 1]
    best_row = anchor_row
    best_score = -1

    for row in candidate_rows:
        if row < cache.start_row or row > cache.end_row:
            continue
        normalized_row = [normalize_text(value) for _, value in cache.row_values(row)]
        score = 0
        for term in terms:
            if any(term in text for text in normalized_row):
                score += 1
        if score > best_score:
            best_score = score
            best_row = row
    return best_row


def find_column_by_phrases(
    cache: SheetCache,
    row: int,
    phrases: Sequence[str],
    anchor_col: int,
) -> Optional[int]:
    matches: List[Tuple[int, int]] = []
    for col, value in cache.row_values(row):
        text = normalize_text(value)
        if not text:
            continue
        for phrase in phrases:
            if phrase in text:
                matches.append((abs(col - anchor_col), col))
                break
    if not matches:
        return None
    matches.sort(key=lambda item: (item[0], item[1]))
    return matches[0][1]


def resolve_col(
    cache: SheetCache,
    header_row: int,
    anchor_col: int,
    phrases: Sequence[str],
    fallback_offset: Optional[int],
) -> Optional[int]:
    header_col = find_column_by_phrases(
        cache=cache,
        row=header_row,
        phrases=phrases,
        anchor_col=anchor_col,
    )
    if header_col is not None:
        return header_col
    if fallback_offset is None:
        return None
    return anchor_col + fallback_offset


def read_cache_value(cache: SheetCache, row: int, col: Optional[int]) -> Any:
    if col is None:
        return None
    return clean_cell_value(cache.get(row, col))


def is_row_empty(values: Sequence[Any]) -> bool:
    for value in values:
        if clean_cell_value(value) is not None:
            return False
    return True


def numeric_rows_for_xy(cache: SheetCache, x_col: int, y_col: int) -> List[int]:
    rows: List[int] = []
    for row in range(cache.start_row, cache.end_row + 1):
        x_value = coerce_float(cache.get(row, x_col))
        y_value = coerce_float(cache.get(row, y_col))
        if x_value is not None and y_value is not None:
            rows.append(row)
    return rows


def normalize_for_compare(value: Any) -> Any:
    if isinstance(value, float):
        return round(value, 10)
    number = coerce_float(value)
    if number is not None:
        return round(number, 10)
    return clean_cell_value(value)


def extract_empirical_rows(wb: xw.Book, file_label: FileLabel, source_file: str) -> List[Dict[str, Any]]:
    sheet_name = "Empirical Model"
    rows: List[Dict[str, Any]] = []

    if sheet_name not in [sheet.name for sheet in wb.sheets]:
        print(f"  skipped empirical: sheet '{sheet_name}' not found")
        return rows

    sheet = wb.sheets[sheet_name]
    cache = SheetCache.from_sheet(sheet)
    anchor = find_anchor_max(cache)
    if anchor is None:
        print("  skipped empirical: could not find 'max' anchor")
        return rows

    anchor_row, anchor_col = anchor
    header_row = choose_header_row(
        cache=cache,
        anchor_row=anchor_row,
        terms=["max", "min", "forecast", "sales", "quarter"],
    )

    empirical_phrases = {
        "num_quarters_used": ("num quarter", "quarters used", "num qtr", "n quarter"),
        "last_quarter_used": ("last quarter", "last qtr"),
        "forecast_value": ("estimated total sold", "forecast value", "total forecast", "tot fcst", "est total"),
        "actual_value": ("actual value", "actual sales", "reported sales"),
        "forecast_max": ("max",),
        "forecast_min": ("min",),
        "avg_penetration_pct": ("avg penetration", "average penetration"),
        "quarterly_sales": ("quarterly sales", "quarter sales"),
        "reported_sales": ("reported sales",),
        "growth_rate_pct": ("growth rate", "growth pct", "growth %"),
        "sales_captured_in_db_pct": ("sales captured in db", "captured in db", "sales captured", "penetration"),
    }

    cols: Dict[str, Optional[int]] = {}
    for field, fallback_offset in EMPIRICAL_FALLBACK_OFFSETS.items():
        cols[field] = resolve_col(
            cache=cache,
            header_row=header_row,
            anchor_col=anchor_col,
            phrases=empirical_phrases.get(field, ()),
            fallback_offset=fallback_offset,
        )

    cols["forecast_max"] = anchor_col
    if normalize_text(cache.get(anchor_row, anchor_col + 1)) == "min":
        cols["forecast_min"] = anchor_col + 1

    start_row = header_row + 1
    data_rows = [start_row + idx for idx in range(N_QUARTERS)]
    scratch_col = cache.end_col + 2

    source_col_for_avg = cols.get("sales_captured_in_db_pct") or cols.get("avg_penetration_pct")
    formula_rows: List[int] = []
    if source_col_for_avg is not None:
        for idx, row in enumerate(data_rows):
            num_quarters_value = coerce_int(read_cache_value(cache, row, cols.get("num_quarters_used")))
            n_quarters = num_quarters_value if num_quarters_value and num_quarters_value > 0 else (idx + 1)
            avg_start_row = max(start_row, row - n_quarters + 1)
            formula = (
                f'=IFERROR(AVERAGE(R{avg_start_row}C{source_col_for_avg}:'
                f"R{row}C{source_col_for_avg}),\"\")"
            )
            sheet.range((row, scratch_col)).formula2 = formula
            formula_rows.append(row)

    if formula_rows:
        wb.app.calculate()

    for idx, row in enumerate(data_rows):
        num_quarters_used = read_cache_value(cache, row, cols.get("num_quarters_used"))
        if num_quarters_used is None:
            num_quarters_used = idx + 1

        last_quarter_used = read_cache_value(cache, row, cols.get("last_quarter_used"))
        forecast_value = read_cache_value(cache, row, cols.get("forecast_value"))
        forecast_max = read_cache_value(cache, row, cols.get("forecast_max"))
        forecast_min = read_cache_value(cache, row, cols.get("forecast_min"))
        quarterly_sales = read_cache_value(cache, row, cols.get("quarterly_sales"))
        reported_sales = read_cache_value(cache, row, cols.get("reported_sales"))
        growth_rate_pct = read_cache_value(cache, row, cols.get("growth_rate_pct"))
        sales_captured_in_db_pct = read_cache_value(cache, row, cols.get("sales_captured_in_db_pct"))

        avg_penetration_pct = None
        if row in formula_rows:
            avg_penetration_pct = clean_cell_value(sheet.range((row, scratch_col)).value)
        if avg_penetration_pct is None:
            avg_penetration_pct = read_cache_value(cache, row, cols.get("avg_penetration_pct"))

        actual_value = reported_sales
        range_width = safe_diff(forecast_max, forecast_min)

        if is_row_empty(
            (
                num_quarters_used,
                forecast_value,
                forecast_max,
                forecast_min,
                avg_penetration_pct,
                reported_sales,
            )
        ):
            continue

        rows.append(
            {
                "model": file_label.model,
                "ticker": file_label.ticker,
                "model_period": file_label.model_period,
                "model_date": file_label.model_date,
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

    if formula_rows:
        sheet.range((start_row, scratch_col), (start_row + N_QUARTERS - 1, scratch_col)).clear_contents()

    return rows


def extract_regression_rows(wb: xw.Book, file_label: FileLabel, source_file: str) -> List[Dict[str, Any]]:
    sheet_name = "Regression Model"
    rows: List[Dict[str, Any]] = []

    if sheet_name not in [sheet.name for sheet in wb.sheets]:
        print(f"  skipped regression: sheet '{sheet_name}' not found")
        return rows

    sheet = wb.sheets[sheet_name]
    cache = SheetCache.from_sheet(sheet)
    anchor = find_anchor_max(cache)
    if anchor is None:
        print("  skipped regression: could not find 'max' anchor")
        return rows

    anchor_row, anchor_col = anchor
    header_row = choose_header_row(
        cache=cache,
        anchor_row=anchor_row,
        terms=["max", "min", "forecast", "quarter"],
    )

    regression_phrases = {
        "num_quarters_used": ("num quarter", "quarters used", "num qtr", "n quarter"),
        "forecast_value": ("tot fcst w o sa", "tot fcst without sa", "forecast w o sa", "without sa", "tot fcst"),
        "forecast_max": ("max",),
        "forecast_min": ("min",),
    }

    cols: Dict[str, Optional[int]] = {}
    for field, fallback_offset in REGRESSION_FALLBACK_OFFSETS.items():
        cols[field] = resolve_col(
            cache=cache,
            header_row=header_row,
            anchor_col=anchor_col,
            phrases=regression_phrases.get(field, ()),
            fallback_offset=fallback_offset,
        )

    cols["forecast_max"] = anchor_col
    if normalize_text(cache.get(anchor_row, anchor_col + 1)) == "min":
        cols["forecast_min"] = anchor_col + 1

    y_col = anchor_col - 7
    x_col = anchor_col - 11
    source_rows = numeric_rows_for_xy(cache=cache, x_col=x_col, y_col=y_col)

    start_row = header_row + 1
    data_rows = [start_row + idx for idx in range(N_QUARTERS)]
    scratch_intercept_col = cache.end_col + 2
    scratch_slope_col = cache.end_col + 3

    formula_rows: List[int] = []
    if len(source_rows) >= 2:
        for idx, row in enumerate(data_rows):
            num_q_value = coerce_int(read_cache_value(cache, row, cols.get("num_quarters_used")))
            num_quarters_used = num_q_value if num_q_value and num_q_value > 1 else (idx + 2)
            n_points = max(2, min(num_quarters_used, len(source_rows)))
            selected_rows = source_rows[-n_points:]
            range_start = selected_rows[0]
            range_end = selected_rows[-1]

            intercept_formula = (
                f'=IFERROR(INTERCEPT(R{range_start}C{y_col}:R{range_end}C{y_col},'
                f"R{range_start}C{x_col}:R{range_end}C{x_col}),\"\")"
            )
            slope_formula = (
                f'=IFERROR(SLOPE(R{range_start}C{y_col}:R{range_end}C{y_col},'
                f"R{range_start}C{x_col}:R{range_end}C{x_col}),\"\")"
            )

            sheet.range((row, scratch_intercept_col)).formula2 = intercept_formula
            sheet.range((row, scratch_slope_col)).formula2 = slope_formula
            formula_rows.append(row)

    if formula_rows:
        wb.app.calculate()

    previous_signature: Optional[Tuple[Any, ...]] = None
    for idx, row in enumerate(data_rows):
        num_quarters_used = read_cache_value(cache, row, cols.get("num_quarters_used"))
        if num_quarters_used is None:
            num_quarters_used = idx + 1

        forecast_value = read_cache_value(cache, row, cols.get("forecast_value"))
        forecast_max = read_cache_value(cache, row, cols.get("forecast_max"))
        forecast_min = read_cache_value(cache, row, cols.get("forecast_min"))
        range_width = safe_diff(forecast_max, forecast_min)

        intercept = clean_cell_value(sheet.range((row, scratch_intercept_col)).value) if row in formula_rows else None
        slope = clean_cell_value(sheet.range((row, scratch_slope_col)).value) if row in formula_rows else None

        if is_row_empty((num_quarters_used, forecast_value, forecast_max, forecast_min, intercept, slope)):
            continue

        signature = (
            normalize_for_compare(num_quarters_used),
            normalize_for_compare(forecast_value),
            normalize_for_compare(forecast_max),
            normalize_for_compare(forecast_min),
            normalize_for_compare(intercept),
            normalize_for_compare(slope),
        )
        if previous_signature is not None and signature == previous_signature:
            continue
        previous_signature = signature

        rows.append(
            {
                "model": file_label.model,
                "ticker": file_label.ticker,
                "model_period": file_label.model_period,
                "model_date": file_label.model_date,
                "method": "regression",
                "parameter_name": "num_quarters_used",
                "parameter_value": num_quarters_used,
                "num_quarters_used": num_quarters_used,
                "forecast_value": forecast_value,
                "actual_value": "",
                "forecast_max": forecast_max,
                "forecast_min": forecast_min,
                "range_width": range_width,
                "intercept": intercept,
                "slope": slope,
                "source_file": source_file,
            }
        )

    if formula_rows:
        sheet.range((start_row, scratch_intercept_col), (start_row + N_QUARTERS - 1, scratch_slope_col)).clear_contents()

    return rows


def apply_sheet_formatting(worksheet) -> None:
    worksheet.freeze_panes = "A2"
    worksheet.auto_filter.ref = worksheet.dimensions

    for col_idx, column_cells in enumerate(
        worksheet.iter_cols(min_row=1, max_row=worksheet.max_row),
        start=1,
    ):
        max_length = 0
        for cell in column_cells:
            if cell.value is None:
                continue
            max_length = max(max_length, len(str(cell.value)))
        worksheet.column_dimensions[get_column_letter(col_idx)].width = min(max(12, max_length + 2), 40)


def write_output_workbook(
    output_path: Path,
    empirical_rows: List[Dict[str, Any]],
    regression_rows: List[Dict[str, Any]],
) -> None:
    workbook = Workbook()
    empirical_sheet = workbook.active
    empirical_sheet.title = "empirical_candidates"
    empirical_sheet.append(EMPIRICAL_COLUMNS)
    for row in empirical_rows:
        empirical_sheet.append([row.get(col) for col in EMPIRICAL_COLUMNS])

    regression_sheet = workbook.create_sheet("regression_candidates")
    regression_sheet.append(REGRESSION_COLUMNS)
    for row in regression_rows:
        regression_sheet.append([row.get(col) for col in REGRESSION_COLUMNS])

    header_font = Font(bold=True)
    for worksheet in (empirical_sheet, regression_sheet):
        for cell in worksheet[1]:
            cell.font = header_font
        apply_sheet_formatting(worksheet)

    workbook.save(output_path)


def process_workbooks(
    input_path: Path,
    output_path: Path,
) -> Tuple[int, List[Dict[str, Any]], List[Dict[str, Any]]]:
    empirical_rows: List[Dict[str, Any]] = []
    regression_rows: List[Dict[str, Any]] = []
    processed_count = 0

    app = xw.App(visible=False, add_book=False)
    app.display_alerts = False
    app.screen_updating = False
    try:
        try:
            app.calculation = "manual"
        except Exception:
            pass

        generated_prefix = f"{input_path.name}_PARAM"

        for file_path in sorted(input_path.iterdir()):
            if not file_path.is_file():
                print(f"Skipped {file_path.name}: not a file")
                continue
            if file_path.name.startswith("~"):
                print(f"Skipped {file_path.name}: temporary file")
                continue
            if file_path.suffix.lower() != ".xlsx":
                print(f"Skipped {file_path.name}: not an .xlsx file")
                continue
            if file_path.resolve() == output_path.resolve():
                print(f"Skipped {file_path.name}: output target file")
                continue
            if file_path.stem.startswith(generated_prefix) or MAX_RE.search(file_path.stem):
                print(f"Skipped {file_path.name}: generated output pattern")
                continue

            print(f"Processing {file_path.name}")
            wb = None
            try:
                wb = app.books.open(str(file_path), update_links=False)
                file_label = parse_file_label(file_path.name)
                empirical_rows.extend(extract_empirical_rows(wb=wb, file_label=file_label, source_file=file_path.name))
                regression_rows.extend(extract_regression_rows(wb=wb, file_label=file_label, source_file=file_path.name))
                processed_count += 1
            except Exception as exc:
                print(f"Skipped {file_path.name}: {exc}")
            finally:
                if wb is not None:
                    safe_close_workbook(wb)
    finally:
        try:
            app.quit()
        except Exception:
            pass

    return processed_count, empirical_rows, regression_rows


def main() -> None:
    input_path = Path(input_dir).expanduser().resolve()
    output_path_root = Path(output_dir).expanduser().resolve()

    if not input_path.exists() or not input_path.is_dir():
        raise FileNotFoundError(f"Input directory does not exist: {input_path}")

    output_path_root.mkdir(parents=True, exist_ok=True)
    output_path = next_output_path(input_path=input_path, output_path_root=output_path_root)

    processed_count, empirical_rows, regression_rows = process_workbooks(
        input_path=input_path,
        output_path=output_path,
    )

    write_output_workbook(
        output_path=output_path,
        empirical_rows=empirical_rows,
        regression_rows=regression_rows,
    )

    print(f"Output path: {output_path}")
    print(f"Files processed: {processed_count}")
    print(f"Empirical rows: {len(empirical_rows)}")
    print(f"Regression rows: {len(regression_rows)}")


if __name__ == "__main__":
    main()
