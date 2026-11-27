#!/usr/bin/env python3
"""Extract empirical/regression model candidate parameters from Excel workbooks."""

from __future__ import annotations

import re
import sys
from dataclasses import dataclass
from datetime import date
from pathlib import Path
from typing import Any, Iterable

import xlwings as xw
from openpyxl import Workbook
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

EMPIRICAL_LABEL_ALIASES = {
    "forecast_max": ["max"],
    "forecast_min": ["min"],
    "forecast_value": ["estimated total sold", "est total sold", "forecast value"],
    "actual_value": ["reported sales", "actual value", "actual sales"],
    "quarterly_sales": ["quarterly sales", "qtr sales"],
    "growth_rate_pct": ["growth rate", "growth rate pct"],
    "sales_captured_in_db_pct": [
        "sales captured in db pct",
        "sales captured in db",
        "captured in db pct",
    ],
    "last_quarter_used": ["last quarter used", "last qtr used", "last quarter"],
}

REGRESSION_LABEL_ALIASES = {
    "forecast_max": ["max"],
    "forecast_min": ["min"],
    "forecast_value": [
        "tot fcst w o sa",
        "tot fcst wo sa",
        "tot fcst without sa",
        "total forecast without sa",
    ],
    "actual_value": ["actual value", "actual sales", "reported sales"],
}

EMPIRICAL_DEFAULT_OFFSETS = {
    "forecast_max": 0,
    "forecast_min": 1,
    "forecast_value": -1,
    "actual_value": -2,
    "quarterly_sales": -3,
    "growth_rate_pct": -4,
    "sales_captured_in_db_pct": -5,
    "last_quarter_used": -6,
}

REGRESSION_DEFAULT_OFFSETS = {
    "forecast_max": 0,
    "forecast_min": 1,
    "forecast_value": -1,
}

MONTH_TO_NUM = {
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

PERIOD_DAY_MAP = {"Early": 5, "Mid": 15, "Late": 25}


@dataclass
class FileLabels:
    model: str
    ticker: str
    model_period: str
    model_date: str


@dataclass
class SheetSnapshot:
    start_row: int
    start_col: int
    values: list[list[Any]]

    def value_at(self, row: int, col: int) -> Any:
        row_idx = row - self.start_row
        col_idx = col - self.start_col
        if row_idx < 0 or col_idx < 0:
            return None
        if row_idx >= len(self.values):
            return None
        row_vals = self.values[row_idx]
        if col_idx >= len(row_vals):
            return None
        return row_vals[col_idx]


def normalize_label(value: Any) -> str:
    if value is None:
        return ""
    text = str(value).strip().lower()
    text = text.replace("%", " pct ")
    text = re.sub(r"[^a-z0-9]+", " ", text)
    return re.sub(r"\s+", " ", text).strip()


def ensure_2d(values: Any) -> list[list[Any]]:
    if values is None:
        return []
    if isinstance(values, (str, bytes)):
        return [[values]]
    if not isinstance(values, (list, tuple)):
        return [[values]]
    if not values:
        return []
    first = values[0]
    if isinstance(first, (list, tuple)):
        return [list(row) for row in values]
    return [list(values)]


def is_number(value: Any) -> bool:
    return isinstance(value, (int, float)) and not isinstance(value, bool)


def to_float(value: Any) -> float | None:
    if value is None:
        return None
    if is_number(value):
        return float(value)
    if isinstance(value, str):
        stripped = value.strip().replace(",", "")
        if not stripped:
            return None
        try:
            return float(stripped)
        except ValueError:
            return None
    return None


def safe_subtract(a: Any, b: Any) -> float | None:
    a_num = to_float(a)
    b_num = to_float(b)
    if a_num is None or b_num is None:
        return None
    return a_num - b_num


def get_output_path(in_dir: Path, out_dir: Path) -> Path:
    out_dir.mkdir(parents=True, exist_ok=True)
    base_stem = f"{in_dir.name}_PARAM"
    first_choice = out_dir / f"{base_stem}.xlsx"
    if not first_choice.exists():
        return first_choice

    suffix = 1
    while True:
        candidate = out_dir / f"{base_stem}.{suffix}.xlsx"
        if not candidate.exists():
            return candidate
        suffix += 1


def get_sheet_case_insensitive(wb: xw.Book, name: str) -> xw.Sheet | None:
    target = name.strip().lower()
    for sheet in wb.sheets:
        if sheet.name.strip().lower() == target:
            return sheet
    return None


def snapshot_sheet(ws: xw.Sheet) -> SheetSnapshot:
    used = ws.used_range
    values = ensure_2d(used.value)
    return SheetSnapshot(start_row=used.row, start_col=used.column, values=values)


def find_anchor(snapshot: SheetSnapshot, anchor_text: str = "max") -> tuple[int, int] | None:
    target = normalize_label(anchor_text)
    for row_idx, row_vals in enumerate(snapshot.values):
        for col_idx, value in enumerate(row_vals):
            if normalize_label(value) == target:
                return snapshot.start_row + row_idx, snapshot.start_col + col_idx
    return None


def detect_series_start_col(snapshot: SheetSnapshot, anchor_row: int, anchor_col: int) -> int:
    row_idx = anchor_row - snapshot.start_row
    if row_idx < 0 or row_idx >= len(snapshot.values):
        return anchor_col + 1

    row_vals = snapshot.values[row_idx]
    for col_idx, value in enumerate(row_vals):
        col = snapshot.start_col + col_idx
        if col <= anchor_col:
            continue
        if is_number(value) or to_float(value) is not None:
            return col

    return anchor_col + 1


def locate_row_offsets(
    snapshot: SheetSnapshot,
    anchor_row: int,
    anchor_col: int,
    alias_map: dict[str, list[str]],
) -> dict[str, int]:
    normalized_aliases: dict[str, set[str]] = {
        field: {normalize_label(alias) for alias in aliases}
        for field, aliases in alias_map.items()
    }

    best_match: dict[str, tuple[int, int]] = {}
    for row_idx, row_vals in enumerate(snapshot.values):
        for col_idx, value in enumerate(row_vals):
            key = normalize_label(value)
            if not key:
                continue
            abs_row = snapshot.start_row + row_idx
            abs_col = snapshot.start_col + col_idx
            col_distance = abs(abs_col - anchor_col)
            if col_distance > 10:
                continue
            row_distance = abs(abs_row - anchor_row)
            score = row_distance * 10 + col_distance
            for field, aliases in normalized_aliases.items():
                if key in aliases:
                    prior = best_match.get(field)
                    if prior is None or score < prior[0]:
                        best_match[field] = (score, abs_row - anchor_row)

    return {field: data[1] for field, data in best_match.items()}


def merge_offsets(primary: dict[str, int], defaults: dict[str, int]) -> dict[str, int]:
    merged = defaults.copy()
    merged.update(primary)
    return merged


def month_to_number(month_token: str) -> int:
    key = month_token.strip().lower()
    if key in MONTH_TO_NUM:
        return MONTH_TO_NUM[key]
    short_key = key[:3]
    if short_key in MONTH_TO_NUM:
        return MONTH_TO_NUM[short_key]
    raise ValueError(f"Unknown month token: {month_token}")


def parse_file_labels(file_path: Path) -> FileLabels:
    stem = file_path.stem
    parts = [part.strip() for part in stem.split(" - ")]
    ticker = parts[1].upper() if len(parts) > 1 and parts[1].strip() else stem.upper()
    period_candidate = parts[2] if len(parts) > 2 else stem

    match = re.search(r"(Early|Mid|Late)\s*([A-Za-z]+)\s*(\d{4})", period_candidate, flags=re.I)
    if match:
        period_prefix = match.group(1).title()
        month_token = match.group(2).title()
        year = int(match.group(3))
        day = PERIOD_DAY_MAP[period_prefix]
        month_num = month_to_number(month_token)
        model_period = f"{period_prefix}{month_token[:3]}_{year}"
        model_date = date(year, month_num, day).isoformat()
    else:
        model_period = "unknown_period"
        model_date = ""

    model = f"{ticker}_{model_period}"
    return FileLabels(
        model=model,
        ticker=ticker,
        model_period=model_period,
        model_date=model_date,
    )


def set_formula2(cell: xw.Range, formula: str) -> None:
    try:
        cell.formula2 = formula
    except Exception:
        # Some Excel builds expose formula but not formula2 via COM bridge.
        cell.formula = formula


def close_source_workbook(wb: xw.Book) -> None:
    try:
        wb.close(save=False)
        return
    except TypeError:
        pass
    except Exception:
        pass

    # Safe fallback for clients that do not support close(save=False).
    try:
        wb.api.Saved = True
        wb.api.Close(SaveChanges=False)
        return
    except Exception:
        pass

    # Last fallback still avoids any explicit save call.
    wb.close()


def extract_empirical_rows(wb: xw.Book, labels: FileLabels, source_file: str) -> list[dict[str, Any]]:
    ws = get_sheet_case_insensitive(wb, "Empirical Model")
    if ws is None:
        print(f"  Skipped empirical extraction: sheet 'Empirical Model' not found.")
        return []

    snapshot = snapshot_sheet(ws)
    anchor = find_anchor(snapshot, "max")
    if anchor is None:
        print("  Skipped empirical extraction: 'max' anchor not found.")
        return []

    anchor_row, anchor_col = anchor
    dynamic_offsets = locate_row_offsets(snapshot, anchor_row, anchor_col, EMPIRICAL_LABEL_ALIASES)
    offsets = merge_offsets(dynamic_offsets, EMPIRICAL_DEFAULT_OFFSETS)
    data_start_col = detect_series_start_col(snapshot, anchor_row, anchor_col)

    used_last_col = ws.used_range.last_cell.column
    scratch_col = max(used_last_col + 2, data_start_col + N_QUARTERS + 2)
    penetration_col = data_start_col - 1 if data_start_col > 1 else max(anchor_col - 1, 1)
    end_row = anchor_row - 1

    formula_rows: list[tuple[int, int]] = []
    if end_row >= 1:
        for n_quarters in range(1, N_QUARTERS + 1):
            start_row = max(1, end_row - n_quarters + 1)
            target_row = anchor_row + n_quarters
            formula = f"=AVERAGE(R{start_row}C{penetration_col}:R{end_row}C{penetration_col})"
            set_formula2(ws.cells(target_row, scratch_col), formula)
            formula_rows.append((n_quarters, target_row))

    if formula_rows:
        wb.app.calculate()

    avg_pen_by_n: dict[int, Any] = {}
    for n_quarters, row in formula_rows:
        avg_pen_by_n[n_quarters] = ws.cells(row, scratch_col).value

    if formula_rows:
        ws.range(
            (anchor_row + 1, scratch_col),
            (anchor_row + N_QUARTERS, scratch_col),
        ).clear_contents()

    rows: list[dict[str, Any]] = []
    for n_quarters in range(1, N_QUARTERS + 1):
        col = data_start_col + (n_quarters - 1)
        forecast_max = snapshot.value_at(anchor_row + offsets["forecast_max"], col)
        forecast_min = snapshot.value_at(anchor_row + offsets["forecast_min"], col)
        forecast_value = snapshot.value_at(anchor_row + offsets["forecast_value"], col)
        actual_value = snapshot.value_at(anchor_row + offsets["actual_value"], col)
        quarterly_sales = snapshot.value_at(anchor_row + offsets["quarterly_sales"], col)
        growth_rate_pct = snapshot.value_at(anchor_row + offsets["growth_rate_pct"], col)
        sales_captured_in_db_pct = snapshot.value_at(
            anchor_row + offsets["sales_captured_in_db_pct"], col
        )
        last_quarter_used = snapshot.value_at(anchor_row + offsets["last_quarter_used"], col)
        avg_penetration_pct = avg_pen_by_n.get(n_quarters)

        if (
            forecast_value is None
            and forecast_max is None
            and forecast_min is None
            and avg_penetration_pct is None
        ):
            continue

        rows.append(
            {
                "model": labels.model,
                "ticker": labels.ticker,
                "model_period": labels.model_period,
                "model_date": labels.model_date,
                "method": "empirical",
                "parameter_name": "avg_penetration_pct",
                "parameter_value": avg_penetration_pct,
                "num_quarters_used": n_quarters,
                "last_quarter_used": last_quarter_used,
                "forecast_value": forecast_value,
                "actual_value": actual_value,
                "forecast_max": forecast_max,
                "forecast_min": forecast_min,
                "range_width": safe_subtract(forecast_max, forecast_min),
                "avg_penetration_pct": avg_penetration_pct,
                "quarterly_sales": quarterly_sales,
                "reported_sales": actual_value,
                "growth_rate_pct": growth_rate_pct,
                "sales_captured_in_db_pct": sales_captured_in_db_pct,
                "source_file": source_file,
            }
        )

    return rows


def same_value(a: Any, b: Any) -> bool:
    if a is None and b is None:
        return True
    a_num = to_float(a)
    b_num = to_float(b)
    if a_num is not None and b_num is not None:
        return abs(a_num - b_num) < 1e-9
    return str(a).strip() == str(b).strip()


def is_duplicate_final_regression_row(previous: dict[str, Any], current: dict[str, Any]) -> bool:
    fields = ["forecast_value", "forecast_max", "forecast_min", "intercept", "slope"]
    return all(same_value(previous.get(field), current.get(field)) for field in fields)


def extract_regression_rows(wb: xw.Book, labels: FileLabels, source_file: str) -> list[dict[str, Any]]:
    ws = get_sheet_case_insensitive(wb, "Regression Model")
    if ws is None:
        print(f"  Skipped regression extraction: sheet 'Regression Model' not found.")
        return []

    snapshot = snapshot_sheet(ws)
    anchor = find_anchor(snapshot, "max")
    if anchor is None:
        print("  Skipped regression extraction: 'max' anchor not found.")
        return []

    anchor_row, anchor_col = anchor
    dynamic_offsets = locate_row_offsets(snapshot, anchor_row, anchor_col, REGRESSION_LABEL_ALIASES)
    offsets = merge_offsets(dynamic_offsets, REGRESSION_DEFAULT_OFFSETS)
    data_start_col = detect_series_start_col(snapshot, anchor_row, anchor_col)

    y_col = anchor_col - 7
    x_col = anchor_col - 11
    if y_col < 1 or x_col < 1:
        print("  Skipped regression extraction: anchor offsets produced invalid x/y columns.")
        return []

    used_last_col = ws.used_range.last_cell.column
    scratch_col = max(used_last_col + 2, data_start_col + N_QUARTERS + 3)

    formula_rows: list[tuple[int, int]] = []
    for n_quarters in range(1, N_QUARTERS + 1):
        start_row = anchor_row - n_quarters
        end_row = anchor_row - 1
        if start_row < 1 or end_row < 1:
            continue

        out_row = anchor_row + n_quarters
        intercept_formula = (
            f"=INTERCEPT(R{start_row}C{y_col}:R{end_row}C{y_col},"
            f"R{start_row}C{x_col}:R{end_row}C{x_col})"
        )
        slope_formula = (
            f"=SLOPE(R{start_row}C{y_col}:R{end_row}C{y_col},"
            f"R{start_row}C{x_col}:R{end_row}C{x_col})"
        )
        set_formula2(ws.cells(out_row, scratch_col), intercept_formula)
        set_formula2(ws.cells(out_row, scratch_col + 1), slope_formula)
        formula_rows.append((n_quarters, out_row))

    if formula_rows:
        wb.app.calculate()

    coefs_by_n: dict[int, tuple[Any, Any]] = {}
    for n_quarters, out_row in formula_rows:
        intercept_val = ws.cells(out_row, scratch_col).value
        slope_val = ws.cells(out_row, scratch_col + 1).value
        coefs_by_n[n_quarters] = (intercept_val, slope_val)

    if formula_rows:
        ws.range(
            (anchor_row + 1, scratch_col),
            (anchor_row + N_QUARTERS, scratch_col + 1),
        ).clear_contents()

    rows: list[dict[str, Any]] = []
    for n_quarters in range(1, N_QUARTERS + 1):
        col = data_start_col + (n_quarters - 1)
        forecast_value = snapshot.value_at(anchor_row + offsets["forecast_value"], col)
        actual_value = snapshot.value_at(anchor_row + offsets.get("actual_value", -2), col)
        forecast_max = snapshot.value_at(anchor_row + offsets["forecast_max"], col)
        forecast_min = snapshot.value_at(anchor_row + offsets["forecast_min"], col)
        intercept_val, slope_val = coefs_by_n.get(n_quarters, (None, None))

        if (
            forecast_value is None
            and forecast_max is None
            and forecast_min is None
            and intercept_val is None
            and slope_val is None
        ):
            continue

        row = {
            "model": labels.model,
            "ticker": labels.ticker,
            "model_period": labels.model_period,
            "model_date": labels.model_date,
            "method": "regression",
            "parameter_name": "num_quarters_used",
            "parameter_value": n_quarters,
            "num_quarters_used": n_quarters,
            "forecast_value": forecast_value,
            "actual_value": actual_value,
            "forecast_max": forecast_max,
            "forecast_min": forecast_min,
            "range_width": safe_subtract(forecast_max, forecast_min),
            "intercept": intercept_val,
            "slope": slope_val,
            "source_file": source_file,
        }

        if n_quarters == N_QUARTERS and rows and is_duplicate_final_regression_row(rows[-1], row):
            continue

        rows.append(row)

    return rows


def write_sheet(ws, columns: list[str], rows: list[dict[str, Any]]) -> None:
    ws.append(columns)
    for row in rows:
        ws.append([row.get(col) for col in columns])

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    for cell in ws[1]:
        cell.font = Font(bold=True)

    for col_idx, col_name in enumerate(columns, start=1):
        max_len = len(col_name)
        for row_idx in range(2, ws.max_row + 1):
            value = ws.cell(row=row_idx, column=col_idx).value
            if value is None:
                continue
            max_len = max(max_len, len(str(value)))
        ws.column_dimensions[get_column_letter(col_idx)].width = min(max(12, max_len + 2), 45)


def write_output_workbook(
    output_path: Path,
    empirical_rows: list[dict[str, Any]],
    regression_rows: list[dict[str, Any]],
) -> None:
    out_wb = Workbook()
    default_sheet = out_wb.active
    out_wb.remove(default_sheet)

    empirical_ws = out_wb.create_sheet("empirical_candidates")
    regression_ws = out_wb.create_sheet("regression_candidates")

    write_sheet(empirical_ws, EMPIRICAL_COLUMNS, empirical_rows)
    write_sheet(regression_ws, REGRESSION_COLUMNS, regression_rows)

    out_wb.save(output_path)


def iter_input_files(in_dir: Path) -> Iterable[Path]:
    for file_path in sorted(in_dir.iterdir()):
        if not file_path.is_file():
            continue
        yield file_path


def main() -> int:
    if not input_dir.exists() or not input_dir.is_dir():
        print(f"Input directory does not exist or is not a directory: {input_dir}")
        return 1

    output_dir.mkdir(parents=True, exist_ok=True)
    output_path = get_output_path(input_dir, output_dir)
    generated_name_pattern = re.compile(
        rf"^{re.escape(input_dir.name)}_PARAM(?:\.\d+)?\.xlsx$", re.IGNORECASE
    )

    files_processed = 0
    empirical_rows: list[dict[str, Any]] = []
    regression_rows: list[dict[str, Any]] = []

    app = xw.App(visible=False, add_book=False)
    app.display_alerts = False
    app.screen_updating = False
    try:
        for file_path in iter_input_files(input_dir):
            name = file_path.name
            suffix = file_path.suffix.lower()
            if name.startswith("~"):
                print(f"Skipped file: {name} (temporary workbook).")
                continue
            if suffix != ".xlsx":
                print(f"Skipped file: {name} (not an .xlsx file).")
                continue
            if generated_name_pattern.match(name):
                print(f"Skipped file: {name} (prior parameter export).")
                continue

            print(f"Processing file: {name}")
            labels = parse_file_labels(file_path)

            source_wb: xw.Book | None = None
            try:
                source_wb = app.books.open(str(file_path), update_links=False)
                empirical_rows.extend(extract_empirical_rows(source_wb, labels, name))
                regression_rows.extend(extract_regression_rows(source_wb, labels, name))
                files_processed += 1
            except Exception as exc:
                print(f"Skipped file: {name} (error: {exc})")
            finally:
                if source_wb is not None:
                    close_source_workbook(source_wb)
    finally:
        app.quit()

    write_output_workbook(output_path, empirical_rows, regression_rows)

    print(f"Output path: {output_path}")
    print(f"Files processed: {files_processed}")
    print(f"Empirical rows: {len(empirical_rows)}")
    print(f"Regression rows: {len(regression_rows)}")

    return 0


if __name__ == "__main__":
    sys.exit(main())
