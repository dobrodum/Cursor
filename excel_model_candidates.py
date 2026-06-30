#!/usr/bin/env python3
"""
Extract empirical/regression model candidates from all .xlsx files in input_dir.

Design goals:
- Open each source workbook only once.
- Process "Empirical Model" and "Regression Model" while workbook is open.
- Never save source workbooks.
- Build a single output workbook with two sheets:
  - empirical_candidates
  - regression_candidates
"""

from __future__ import annotations

import re
from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Sequence, Tuple

import xlwings as xw
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter


# =========================
# User-configurable paths
# =========================
input_dir = Path("./input")
output_dir = Path("./output")


N_QUARTERS = 10

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

PERIOD_DAY_MAP = {
    "early": 5,
    "mid": 15,
    "late": 25,
}


@dataclass(frozen=True)
class FileLabel:
    model: str
    ticker: str
    model_period: str
    model_date: str


def normalize_text(value: Any) -> str:
    if value is None:
        return ""
    text = str(value).strip().lower()
    text = re.sub(r"[\s\-_/%]+", " ", text)
    text = re.sub(r"[^a-z0-9 ]", "", text)
    return text.strip()


def to_2d(values: Any) -> List[List[Any]]:
    if values is None:
        return []
    if not isinstance(values, list):
        return [[values]]
    if values and not isinstance(values[0], list):
        return [values]
    rows = values
    max_cols = max((len(row) for row in rows), default=0)
    normalized: List[List[Any]] = []
    for row in rows:
        padded = list(row) + [None] * (max_cols - len(row))
        normalized.append(padded)
    return normalized


class UsedRangeCache:
    def __init__(self, ws: xw.Sheet):
        self.ws = ws
        used = ws.used_range
        self.top = used.row
        self.left = used.column
        self.values = to_2d(used.value)
        self.row_count = len(self.values)
        self.col_count = len(self.values[0]) if self.values else 0
        self.bottom = self.top + self.row_count - 1
        self.right = self.left + self.col_count - 1

    def in_bounds(self, row: int, col: int) -> bool:
        return (
            self.values
            and self.top <= row <= self.bottom
            and self.left <= col <= self.right
        )

    def get(self, row: int, col: int) -> Any:
        if not self.in_bounds(row, col):
            return None
        return self.values[row - self.top][col - self.left]

    def find_max_anchor(self) -> Optional[Tuple[int, int]]:
        matches: List[Tuple[int, int]] = []
        for row in range(self.top, self.bottom + 1):
            for col in range(self.left, self.right + 1):
                if normalize_text(self.get(row, col)) == "max":
                    matches.append((row, col))

        if not matches:
            return None

        # Prefer anchors with a nearby "min" label on same row.
        for row, col in matches:
            for min_col in range(col, min(col + 4, self.right + 1)):
                if normalize_text(self.get(row, min_col)) == "min":
                    return row, col

        return matches[0]

    def find_col_on_row(
        self,
        row: int,
        aliases: Sequence[str],
        col_min: int,
        col_max: int,
        prefer_col: Optional[int] = None,
    ) -> Optional[int]:
        if not self.values:
            return None
        left = max(self.left, col_min)
        right = min(self.right, col_max)
        if left > right:
            return None

        normalized_aliases = [normalize_text(alias) for alias in aliases]
        hits: List[int] = []
        for col in range(left, right + 1):
            cell_text = normalize_text(self.get(row, col))
            if not cell_text:
                continue
            for alias in normalized_aliases:
                if alias and alias in cell_text:
                    hits.append(col)
                    break

        if not hits:
            return None
        if prefer_col is None:
            return hits[0]
        return min(hits, key=lambda c: abs(c - prefer_col))

    def find_col_in_block(
        self,
        aliases: Sequence[str],
        row_min: int,
        row_max: int,
        col_min: int,
        col_max: int,
        prefer_col: Optional[int] = None,
    ) -> Optional[int]:
        if not self.values:
            return None
        top = max(self.top, row_min)
        bottom = min(self.bottom, row_max)
        left = max(self.left, col_min)
        right = min(self.right, col_max)
        if top > bottom or left > right:
            return None

        normalized_aliases = [normalize_text(alias) for alias in aliases]
        hits: List[Tuple[int, int]] = []
        for row in range(top, bottom + 1):
            for col in range(left, right + 1):
                cell_text = normalize_text(self.get(row, col))
                if not cell_text:
                    continue
                for alias in normalized_aliases:
                    if alias and alias in cell_text:
                        hits.append((row, col))
                        break

        if not hits:
            return None

        if prefer_col is None:
            return hits[-1][1]
        return min((col for _, col in hits), key=lambda c: abs(c - prefer_col))


def to_float(value: Any) -> Optional[float]:
    if value is None:
        return None
    if isinstance(value, (int, float)):
        return float(value)
    text = str(value).strip().replace(",", "")
    if not text:
        return None
    if text.endswith("%"):
        text = text[:-1]
        try:
            return float(text) / 100.0
        except ValueError:
            return None
    try:
        return float(text)
    except ValueError:
        return None


def to_int(value: Any) -> Optional[int]:
    number = to_float(value)
    if number is None:
        return None
    return int(round(number))


def display_value(value: Any) -> Any:
    if isinstance(value, datetime):
        return value.date().isoformat()
    if isinstance(value, date):
        return value.isoformat()
    return value


def all_blank(values: Iterable[Any]) -> bool:
    for value in values:
        if value is None:
            continue
        if isinstance(value, str) and not value.strip():
            continue
        return False
    return True


def set_formula2(cell: xw.Range, formula: str) -> None:
    try:
        cell.formula2 = formula
    except Exception:
        # Formula2 may not be supported by all Excel builds.
        cell.formula = formula


def safe_close_workbook(wb: xw.Book) -> None:
    try:
        wb.close(save=False)
        return
    except TypeError:
        pass
    except Exception:
        pass

    # Safe fallback for environments where wb.close(save=False) is unsupported.
    try:
        wb.api.Close(SaveChanges=False)
    except Exception:
        try:
            wb.close()
        except Exception:
            pass


def get_unique_output_path(in_dir: Path, out_dir: Path) -> Path:
    base_name = f"{in_dir.name}_PARAM.xlsx"
    candidate = out_dir / base_name
    if not candidate.exists():
        return candidate

    index = 1
    while True:
        candidate = out_dir / f"{in_dir.name}_PARAM.{index}.xlsx"
        if not candidate.exists():
            return candidate
        index += 1


def parse_file_label(file_path: Path) -> FileLabel:
    stem = file_path.stem
    parts = [part.strip() for part in stem.split(" - ")]
    ticker = parts[1] if len(parts) >= 2 else "UNKNOWN"
    period_token = parts[2] if len(parts) >= 3 else ""
    period_token = re.sub(r"_send$", "", period_token, flags=re.IGNORECASE)
    period_token = period_token.split("_")[0]

    match = re.search(r"(?i)\b(early|mid|late)([a-z]{3})(\d{4})\b", period_token)
    if not match:
        match = re.search(r"(?i)\b(early|mid|late)([a-z]{3})(\d{4})\b", stem)

    model_period = "unknown_0000"
    model_date = ""
    if match:
        period_word, month_abbr, year_str = match.groups()
        month_num = MONTH_MAP.get(month_abbr.lower())
        day_num = PERIOD_DAY_MAP[period_word.lower()]
        if month_num is not None:
            year_num = int(year_str)
            model_period = f"{period_word.capitalize()}{month_abbr.capitalize()}_{year_num}"
            model_date = date(year_num, month_num, day_num).isoformat()

    model = f"{ticker}_{model_period}"
    return FileLabel(
        model=model,
        ticker=ticker,
        model_period=model_period,
        model_date=model_date,
    )


def map_columns_near_anchor(
    cache: UsedRangeCache,
    anchor_row: int,
    anchor_col: int,
    aliases_by_field: Dict[str, Sequence[str]],
    default_offsets: Dict[str, int],
    search_left: int = 30,
    search_right: int = 15,
) -> Dict[str, int]:
    mapped: Dict[str, int] = {}
    row_min_col = max(1, anchor_col - search_left)
    row_max_col = anchor_col + search_right

    for field, aliases in aliases_by_field.items():
        found = cache.find_col_on_row(
            row=anchor_row,
            aliases=aliases,
            col_min=row_min_col,
            col_max=row_max_col,
            prefer_col=anchor_col,
        )
        if found is not None:
            mapped[field] = found
            continue
        mapped[field] = max(1, anchor_col + default_offsets[field])

    # Strong anchor defaults for max/min headers.
    mapped["forecast_max"] = anchor_col
    min_on_right = normalize_text(cache.get(anchor_row, anchor_col + 1)) == "min"
    mapped["forecast_min"] = anchor_col + 1 if min_on_right else mapped["forecast_min"]
    return mapped


def extract_empirical_rows(
    wb: xw.Book,
    ws: xw.Sheet,
    label: FileLabel,
    source_file: str,
) -> List[Dict[str, Any]]:
    cache = UsedRangeCache(ws)
    anchor = cache.find_max_anchor()
    if anchor is None:
        print("  skipped empirical extraction: 'max' anchor not found")
        return []

    anchor_row, anchor_col = anchor

    aliases = {
        "num_quarters_used": ["num quarters", "quarters used", "# quarters", "n quarters"],
        "last_quarter_used": ["last quarter"],
        "forecast_value": ["estimated total sold", "forecast", "tot fcst"],
        "actual_value": ["actual", "reported sales"],
        "forecast_max": ["max"],
        "forecast_min": ["min"],
        "quarterly_sales": ["quarterly sales"],
        "reported_sales": ["reported sales"],
        "growth_rate_pct": ["growth rate", "growth %"],
        "sales_captured_in_db_pct": ["sales captured", "captured in db"],
    }
    default_offsets = {
        "num_quarters_used": -9,
        "last_quarter_used": -8,
        "forecast_value": -7,
        "actual_value": -6,
        "forecast_max": 0,
        "forecast_min": 1,
        "quarterly_sales": -5,
        "reported_sales": -4,
        "growth_rate_pct": -3,
        "sales_captured_in_db_pct": -2,
    }
    col_map = map_columns_near_anchor(cache, anchor_row, anchor_col, aliases, default_offsets)

    # Try to locate a penetration series column for formula-based average penetration.
    penetration_col = cache.find_col_in_block(
        aliases=["penetration", "penetration %", "pen pct"],
        row_min=cache.top,
        row_max=anchor_row,
        col_min=max(1, anchor_col - 40),
        col_max=anchor_col + 5,
        prefer_col=anchor_col,
    )
    data_end_row = anchor_row - 1
    data_floor = cache.top + 1

    rows_meta: List[Dict[str, Any]] = []
    temp_formula_rows: List[int] = []
    temp_formula_col = cache.right + 2
    temp_formula_start = cache.bottom + 2

    for i in range(1, N_QUARTERS + 1):
        row = anchor_row + i
        num_quarters_used = to_int(cache.get(row, col_map["num_quarters_used"])) or i
        last_quarter_used = display_value(cache.get(row, col_map["last_quarter_used"]))
        forecast_value = display_value(cache.get(row, col_map["forecast_value"]))
        actual_value = display_value(cache.get(row, col_map["actual_value"]))
        forecast_max = display_value(cache.get(row, col_map["forecast_max"]))
        forecast_min = display_value(cache.get(row, col_map["forecast_min"]))
        quarterly_sales = display_value(cache.get(row, col_map["quarterly_sales"]))
        reported_sales = display_value(cache.get(row, col_map["reported_sales"]))
        growth_rate_pct = display_value(cache.get(row, col_map["growth_rate_pct"]))
        sales_captured_in_db_pct = display_value(
            cache.get(row, col_map["sales_captured_in_db_pct"])
        )

        if all_blank(
            [
                forecast_value,
                actual_value,
                forecast_max,
                forecast_min,
                quarterly_sales,
                reported_sales,
            ]
        ):
            continue

        avg_penetration_pct = None
        temp_row = temp_formula_start + len(rows_meta)
        if penetration_col is not None and num_quarters_used > 0 and data_end_row >= data_floor:
            start_row = max(data_floor, data_end_row - num_quarters_used + 1)
            formula = f"=AVERAGE(R{start_row}C{penetration_col}:R{data_end_row}C{penetration_col})"
            set_formula2(ws.cells(temp_row, temp_formula_col), formula)
            temp_formula_rows.append(temp_row)

        rows_meta.append(
            {
                "sheet_row": row,
                "num_quarters_used": num_quarters_used,
                "last_quarter_used": last_quarter_used,
                "forecast_value": forecast_value,
                "actual_value": actual_value,
                "forecast_max": forecast_max,
                "forecast_min": forecast_min,
                "quarterly_sales": quarterly_sales,
                "reported_sales": reported_sales,
                "growth_rate_pct": growth_rate_pct,
                "sales_captured_in_db_pct": sales_captured_in_db_pct,
                "avg_penetration_pct": avg_penetration_pct,
                "temp_formula_row": temp_row if temp_row in temp_formula_rows else None,
            }
        )

    if temp_formula_rows:
        wb.app.calculate()
        for meta in rows_meta:
            temp_row = meta["temp_formula_row"]
            if temp_row is not None:
                meta["avg_penetration_pct"] = display_value(
                    ws.cells(temp_row, temp_formula_col).value
                )
        ws.range(
            (temp_formula_start, temp_formula_col),
            (temp_formula_start + len(rows_meta) - 1, temp_formula_col),
        ).clear_contents()

    output_rows: List[Dict[str, Any]] = []
    for meta in rows_meta:
        max_val = to_float(meta["forecast_max"])
        min_val = to_float(meta["forecast_min"])
        range_width = (max_val - min_val) if max_val is not None and min_val is not None else None
        avg_pen = meta["avg_penetration_pct"]

        output_rows.append(
            {
                "model": label.model,
                "ticker": label.ticker,
                "model_period": label.model_period,
                "model_date": label.model_date,
                "method": "empirical",
                "parameter_name": "avg_penetration_pct",
                "parameter_value": avg_pen,
                "num_quarters_used": meta["num_quarters_used"],
                "last_quarter_used": meta["last_quarter_used"],
                "forecast_value": meta["forecast_value"],
                "actual_value": meta["actual_value"],
                "forecast_max": meta["forecast_max"],
                "forecast_min": meta["forecast_min"],
                "range_width": range_width,
                "avg_penetration_pct": avg_pen,
                "quarterly_sales": meta["quarterly_sales"],
                "reported_sales": meta["reported_sales"],
                "growth_rate_pct": meta["growth_rate_pct"],
                "sales_captured_in_db_pct": meta["sales_captured_in_db_pct"],
                "source_file": source_file,
            }
        )

    return output_rows


def signature_for_regression_row(row: Dict[str, Any]) -> Tuple[Any, ...]:
    def rounded(value: Any) -> Any:
        number = to_float(value)
        if number is None:
            return value
        return round(number, 10)

    return (
        row.get("num_quarters_used"),
        rounded(row.get("forecast_value")),
        rounded(row.get("forecast_max")),
        rounded(row.get("forecast_min")),
        rounded(row.get("intercept")),
        rounded(row.get("slope")),
    )


def extract_regression_rows(
    wb: xw.Book,
    ws: xw.Sheet,
    label: FileLabel,
    source_file: str,
) -> List[Dict[str, Any]]:
    cache = UsedRangeCache(ws)
    anchor = cache.find_max_anchor()
    if anchor is None:
        print("  skipped regression extraction: 'max' anchor not found")
        return []

    anchor_row, anchor_col = anchor
    y_col = anchor_col - 7
    x_col = anchor_col - 11

    aliases = {
        "num_quarters_used": ["num quarters", "quarters used", "# quarters", "n quarters"],
        "forecast_value": ["tot fcst w/o sa", "tot fcst without sa", "forecast", "tot fcst"],
        "actual_value": ["actual", "reported sales"],
        "forecast_max": ["max"],
        "forecast_min": ["min"],
    }
    default_offsets = {
        "num_quarters_used": -9,
        "forecast_value": -6,
        "actual_value": -5,
        "forecast_max": 0,
        "forecast_min": 1,
    }
    col_map = map_columns_near_anchor(cache, anchor_row, anchor_col, aliases, default_offsets)

    temp_start_row = cache.bottom + 2
    temp_intercept_col = cache.right + 2
    temp_slope_col = cache.right + 3
    rows_meta: List[Dict[str, Any]] = []

    data_end_row = anchor_row - 1
    data_floor = cache.top + 1

    for i in range(1, N_QUARTERS + 1):
        row = anchor_row + i
        num_quarters_used = to_int(cache.get(row, col_map["num_quarters_used"])) or i
        forecast_value = display_value(cache.get(row, col_map["forecast_value"]))
        actual_value = display_value(cache.get(row, col_map["actual_value"]))
        forecast_max = display_value(cache.get(row, col_map["forecast_max"]))
        forecast_min = display_value(cache.get(row, col_map["forecast_min"]))

        if all_blank([forecast_value, forecast_max, forecast_min, actual_value]):
            continue
        if num_quarters_used <= 0:
            continue

        data_start_row = max(data_floor, data_end_row - num_quarters_used + 1)
        if data_start_row > data_end_row:
            continue

        temp_row = temp_start_row + len(rows_meta)
        intercept_formula = (
            f"=INTERCEPT(R{data_start_row}C{y_col}:R{data_end_row}C{y_col},"
            f"R{data_start_row}C{x_col}:R{data_end_row}C{x_col})"
        )
        slope_formula = (
            f"=SLOPE(R{data_start_row}C{y_col}:R{data_end_row}C{y_col},"
            f"R{data_start_row}C{x_col}:R{data_end_row}C{x_col})"
        )
        set_formula2(ws.cells(temp_row, temp_intercept_col), intercept_formula)
        set_formula2(ws.cells(temp_row, temp_slope_col), slope_formula)

        rows_meta.append(
            {
                "sheet_row": row,
                "num_quarters_used": num_quarters_used,
                "forecast_value": forecast_value,
                "actual_value": actual_value,
                "forecast_max": forecast_max,
                "forecast_min": forecast_min,
                "temp_row": temp_row,
            }
        )

    if rows_meta:
        wb.app.calculate()
        for meta in rows_meta:
            temp_row = meta["temp_row"]
            intercept = display_value(ws.cells(temp_row, temp_intercept_col).value)
            slope = display_value(ws.cells(temp_row, temp_slope_col).value)
            meta["intercept"] = intercept
            meta["slope"] = slope

            if meta["forecast_value"] is None:
                x_future = to_float(cache.get(meta["sheet_row"], x_col))
                intercept_num = to_float(intercept)
                slope_num = to_float(slope)
                if (
                    x_future is not None
                    and intercept_num is not None
                    and slope_num is not None
                ):
                    meta["forecast_value"] = intercept_num + slope_num * x_future

        ws.range(
            (temp_start_row, temp_intercept_col),
            (temp_start_row + len(rows_meta) - 1, temp_slope_col),
        ).clear_contents()

    output_rows: List[Dict[str, Any]] = []
    prev_signature: Optional[Tuple[Any, ...]] = None
    for meta in rows_meta:
        max_val = to_float(meta["forecast_max"])
        min_val = to_float(meta["forecast_min"])
        range_width = (max_val - min_val) if max_val is not None and min_val is not None else None

        row = {
            "model": label.model,
            "ticker": label.ticker,
            "model_period": label.model_period,
            "model_date": label.model_date,
            "method": "regression",
            "parameter_name": "num_quarters_used",
            "parameter_value": meta["num_quarters_used"],
            "num_quarters_used": meta["num_quarters_used"],
            "forecast_value": meta["forecast_value"],
            "actual_value": meta["actual_value"] if meta["actual_value"] is not None else "",
            "forecast_max": meta["forecast_max"],
            "forecast_min": meta["forecast_min"],
            "range_width": range_width,
            "intercept": meta.get("intercept"),
            "slope": meta.get("slope"),
            "source_file": source_file,
        }

        sig = signature_for_regression_row(row)
        if prev_signature is not None and sig == prev_signature:
            continue
        prev_signature = sig
        output_rows.append(row)

    return output_rows


def write_sheet(
    ws,
    headers: Sequence[str],
    rows: Sequence[Dict[str, Any]],
) -> None:
    ws.append(list(headers))
    for header_cell in ws[1]:
        header_cell.font = Font(bold=True)

    for row in rows:
        ws.append([row.get(header, "") for header in headers])

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(headers))}{max(ws.max_row, 1)}"

    for idx, header in enumerate(headers, start=1):
        max_len = len(header)
        for row_idx in range(2, ws.max_row + 1):
            value = ws.cell(row=row_idx, column=idx).value
            text_len = len(str(value)) if value is not None else 0
            if text_len > max_len:
                max_len = text_len
        ws.column_dimensions[get_column_letter(idx)].width = min(max(12, max_len + 2), 48)


def write_output_workbook(
    output_path: Path,
    empirical_rows: Sequence[Dict[str, Any]],
    regression_rows: Sequence[Dict[str, Any]],
) -> None:
    wb = Workbook()
    default_sheet = wb.active
    wb.remove(default_sheet)

    ws_empirical = wb.create_sheet("empirical_candidates")
    write_sheet(ws_empirical, EMPIRICAL_HEADERS, empirical_rows)

    ws_regression = wb.create_sheet("regression_candidates")
    write_sheet(ws_regression, REGRESSION_HEADERS, regression_rows)

    wb.save(output_path)


def iter_input_files(in_dir: Path) -> Iterable[Path]:
    for path in sorted(in_dir.iterdir()):
        if not path.is_file():
            continue
        if path.name.startswith("~"):
            print(f"Skipping {path.name}: temporary file")
            continue
        if path.suffix.lower() != ".xlsx":
            print(f"Skipping {path.name}: not an .xlsx file")
            continue
        yield path


def process_all_workbooks() -> None:
    if not input_dir.exists():
        raise FileNotFoundError(f"input_dir does not exist: {input_dir}")

    output_dir.mkdir(parents=True, exist_ok=True)
    output_path = get_unique_output_path(input_dir, output_dir)

    empirical_rows: List[Dict[str, Any]] = []
    regression_rows: List[Dict[str, Any]] = []
    processed_file_count = 0

    app = xw.App(visible=False, add_book=False)
    app.display_alerts = False
    app.screen_updating = False

    try:
        for file_path in iter_input_files(input_dir):
            print(f"Processing {file_path.name}")
            wb: Optional[xw.Book] = None
            try:
                label = parse_file_label(file_path)
                wb = app.books.open(str(file_path), update_links=False)

                try:
                    empirical_ws = wb.sheets["Empirical Model"]
                except Exception:
                    empirical_ws = None
                    print(f"  skipped empirical extraction: missing sheet 'Empirical Model'")

                try:
                    regression_ws = wb.sheets["Regression Model"]
                except Exception:
                    regression_ws = None
                    print(f"  skipped regression extraction: missing sheet 'Regression Model'")

                if empirical_ws is not None:
                    empirical_rows.extend(
                        extract_empirical_rows(
                            wb=wb,
                            ws=empirical_ws,
                            label=label,
                            source_file=file_path.name,
                        )
                    )
                if regression_ws is not None:
                    regression_rows.extend(
                        extract_regression_rows(
                            wb=wb,
                            ws=regression_ws,
                            label=label,
                            source_file=file_path.name,
                        )
                    )

                processed_file_count += 1
            except Exception as exc:
                print(f"Skipping {file_path.name}: {exc}")
            finally:
                if wb is not None:
                    safe_close_workbook(wb)
    finally:
        app.quit()

    write_output_workbook(output_path, empirical_rows, regression_rows)

    print(f"Output workbook: {output_path}")
    print(f"Files processed: {processed_file_count}")
    print(f"Empirical rows: {len(empirical_rows)}")
    print(f"Regression rows: {len(regression_rows)}")


if __name__ == "__main__":
    process_all_workbooks()
