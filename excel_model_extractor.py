#!/usr/bin/env python3
"""Extract empirical and regression candidates from Excel model workbooks.

The script scans all `.xlsx` files in `input_dir`, extracts both
`Empirical Model` and `Regression Model` candidates while each source workbook
is open exactly once, and writes a single output workbook with:
  - empirical_candidates
  - regression_candidates
"""

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

# User-configurable paths (required by prompt).
input_dir = Path("./input")
output_dir = Path("./output")

N_QUARTERS = 10
QUARTER_DAY_MAP = {"early": 5, "mid": 15, "late": 25}
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


@dataclass(frozen=True)
class FileMetadata:
    ticker: str
    model_period: str
    model_date: str
    model: str


@dataclass
class SheetSnapshot:
    row0: int
    col0: int
    values: list[list[Any]]
    n_rows: int
    n_cols: int
    text_positions: dict[str, list[tuple[int, int]]]

    @classmethod
    def from_sheet(cls, sheet: xw.Sheet) -> "SheetSnapshot":
        used = sheet.used_range
        row0 = used.row
        col0 = used.column
        matrix = to_matrix(used.value)
        n_rows = len(matrix)
        n_cols = max((len(row) for row in matrix), default=0)

        text_positions: dict[str, list[tuple[int, int]]] = {}
        for r_off, row_vals in enumerate(matrix):
            for c_off, cell_val in enumerate(row_vals):
                if isinstance(cell_val, str):
                    token = normalize_text(cell_val)
                    if token:
                        text_positions.setdefault(token, []).append((row0 + r_off, col0 + c_off))

        return cls(
            row0=row0,
            col0=col0,
            values=matrix,
            n_rows=n_rows,
            n_cols=n_cols,
            text_positions=text_positions,
        )

    def get(self, row: int, col: int) -> Any:
        r_idx = row - self.row0
        c_idx = col - self.col0
        if r_idx < 0 or c_idx < 0 or r_idx >= self.n_rows:
            return None
        row_vals = self.values[r_idx]
        if c_idx >= len(row_vals):
            return None
        return row_vals[c_idx]

    def iter_positions_for_labels(self, labels: Iterable[str]) -> list[tuple[int, int]]:
        normalized = [normalize_text(label) for label in labels]
        results: list[tuple[int, int]] = []
        for text, coords in self.text_positions.items():
            if any(token and (text == token or token in text) for token in normalized):
                results.extend(coords)
        return results


def to_matrix(values: Any) -> list[list[Any]]:
    if values is None:
        return []
    if isinstance(values, list):
        if not values:
            return []
        if isinstance(values[0], list):
            return [list(row) for row in values]
        return [list(values)]
    return [[values]]


def normalize_text(value: str) -> str:
    return re.sub(r"\s+", " ", value.strip().lower())


def is_number(value: Any) -> bool:
    return isinstance(value, (int, float)) and not isinstance(value, bool)


def to_float(value: Any) -> float | None:
    if value is None or value == "":
        return None
    if is_number(value):
        return float(value)
    try:
        return float(str(value).replace(",", ""))
    except (TypeError, ValueError):
        return None


def approx_equal(left: float | None, right: float | None, tol: float = 1e-9) -> bool:
    if left is None and right is None:
        return True
    if left is None or right is None:
        return False
    return abs(left - right) <= tol


def get_sheet_by_name(workbook: xw.Book, expected_name: str) -> xw.Sheet | None:
    expected = normalize_text(expected_name)
    for sheet in workbook.sheets:
        if normalize_text(sheet.name) == expected:
            return sheet
    return None


def find_anchor_max(snapshot: SheetSnapshot) -> tuple[int, int] | None:
    max_positions = snapshot.iter_positions_for_labels(["max"])
    if not max_positions:
        return None

    min_positions = snapshot.iter_positions_for_labels(["min"])
    if not min_positions:
        return max_positions[0]

    def score(position: tuple[int, int]) -> int:
        row, col = position
        return min(abs(row - min_row) + abs(col - min_col) for min_row, min_col in min_positions)

    return min(max_positions, key=score)


def find_numeric_near_label(
    snapshot: SheetSnapshot,
    labels: Iterable[str],
    near: tuple[int, int] | None = None,
) -> float | None:
    candidates = snapshot.iter_positions_for_labels(labels)
    if near is not None and candidates:
        near_row, near_col = near
        candidates.sort(key=lambda rc: abs(rc[0] - near_row) + abs(rc[1] - near_col))

    neighbor_offsets = [
        (0, 1),
        (0, 2),
        (0, -1),
        (1, 0),
        (1, 1),
        (2, 0),
        (-1, 0),
    ]
    for row, col in candidates:
        direct = to_float(snapshot.get(row, col))
        if direct is not None:
            return direct
        for dr, dc in neighbor_offsets:
            value = to_float(snapshot.get(row + dr, col + dc))
            if value is not None:
                return value
    return None


def find_row_for_labels(
    snapshot: SheetSnapshot,
    labels: Iterable[str],
    fallback_row: int,
) -> int:
    candidates = snapshot.iter_positions_for_labels(labels)
    if not candidates:
        return fallback_row
    return min(candidates, key=lambda rc: abs(rc[0] - fallback_row))[0]


def safe_close_workbook(workbook: xw.Book) -> None:
    try:
        workbook.close(save=False)
        return
    except TypeError:
        # Some backends do not expose keyword args.
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
    except Exception as exc:
        print(f"Warning: workbook close fallback failed: {exc}")


def parse_file_metadata(file_path: Path) -> FileMetadata:
    stem = file_path.stem
    parts = [part.strip() for part in stem.split("-")]
    ticker = "UNKNOWN"
    period_fragment = stem

    if len(parts) >= 2:
        ticker = re.sub(r"[^A-Za-z0-9]", "", parts[1]).upper() or "UNKNOWN"
    if len(parts) >= 3:
        period_fragment = parts[2]

    match = re.search(r"(Early|Mid|Late)\s*([A-Za-z]+)\s*(\d{4})", period_fragment, flags=re.IGNORECASE)
    if match is None:
        match = re.search(r"(Early|Mid|Late)\s*([A-Za-z]+)\s*(\d{4})", stem, flags=re.IGNORECASE)

    if match:
        period_word = match.group(1).title()
        month_token = match.group(2)
        year = int(match.group(3))
        month_key = month_token.lower()
        month_number = MONTH_MAP.get(month_key)
        if month_number is None:
            month_number = MONTH_MAP.get(month_key[:3], 1)
        month_short = date(2000, month_number, 1).strftime("%b")
        model_period = f"{period_word}{month_short}_{year}"
        model_day = QUARTER_DAY_MAP[period_word.lower()]
        model_date = date(year, month_number, model_day).isoformat()
    else:
        model_period = "unknown_period"
        model_date = ""

    model = f"{ticker}_{model_period}"
    return FileMetadata(ticker=ticker, model_period=model_period, model_date=model_date, model=model)


def unique_output_path(in_dir: Path, out_dir: Path) -> Path:
    folder_name = in_dir.name
    base = out_dir / f"{folder_name}_PARAM.xlsx"
    if not base.exists():
        return base

    counter = 1
    while True:
        candidate = out_dir / f"{folder_name}_PARAM.{counter}.xlsx"
        if not candidate.exists():
            return candidate
        counter += 1


def r1c1_range(row1: int, col1: int, row2: int, col2: int) -> str:
    return f"R{row1}C{col1}:R{row2}C{col2}"


def extract_empirical_rows(workbook: xw.Book, metadata: FileMetadata, source_file: str) -> list[dict[str, Any]]:
    sheet = get_sheet_by_name(workbook, "Empirical Model")
    if sheet is None:
        print(f"SKIP {source_file}: missing Empirical Model sheet")
        return []

    snapshot = SheetSnapshot.from_sheet(sheet)
    anchor = find_anchor_max(snapshot)
    if anchor is None:
        print(f"SKIP {source_file}: Empirical Model max anchor not found")
        return []

    anchor_row, anchor_col = anchor
    end_col = anchor_col - 1
    if end_col < snapshot.col0:
        print(f"SKIP {source_file}: Empirical Model invalid quarter columns")
        return []

    penetration_row = find_row_for_labels(snapshot, ["penetration"], fallback_row=anchor_row - 8)
    quarterly_sales_row = find_row_for_labels(
        snapshot,
        ["quarterly sales", "qtr sales", "quarter sales"],
        fallback_row=anchor_row - 7,
    )
    reported_sales_row = find_row_for_labels(
        snapshot,
        ["reported sales", "actual sales"],
        fallback_row=anchor_row - 6,
    )
    growth_rate_row = find_row_for_labels(snapshot, ["growth rate"], fallback_row=anchor_row - 5)
    captured_row = find_row_for_labels(
        snapshot,
        ["sales captured in db", "captured in db", "captured"],
        fallback_row=anchor_row - 4,
    )
    quarter_label_row = penetration_row - 1

    available_quarters = end_col - snapshot.col0 + 1
    row_specs: list[tuple[int, int]] = []
    for n_quarters in range(1, min(N_QUARTERS, available_quarters) + 1):
        start_col = end_col - n_quarters + 1
        if start_col >= snapshot.col0:
            row_specs.append((n_quarters, start_col))
    if not row_specs:
        return []

    calc_start_row = max(anchor_row + 2, snapshot.row0 + snapshot.n_rows + 2)
    calc_start_col = max(anchor_col + 2, snapshot.col0 + snapshot.n_cols + 2)
    formulas: list[list[str]] = []
    for idx, (_, start_col) in enumerate(row_specs):
        calc_row = calc_start_row + idx
        avg_formula = f"=AVERAGE({r1c1_range(penetration_row, start_col, penetration_row, end_col)})"
        forecast_formula = f"=R{calc_row}C{calc_start_col}*R{quarterly_sales_row}C{end_col}"
        formulas.append([avg_formula, forecast_formula])

    calc_end_row = calc_start_row + len(formulas) - 1
    calc_range = sheet.range((calc_start_row, calc_start_col), (calc_end_row, calc_start_col + 1))
    calc_range.formula2 = formulas
    workbook.app.calculate()
    calc_values = to_matrix(calc_range.value)
    calc_range.clear_contents()

    forecast_max = find_numeric_near_label(snapshot, ["max"], near=anchor)
    forecast_min = find_numeric_near_label(snapshot, ["min"], near=anchor)
    quarterly_sales = to_float(snapshot.get(quarterly_sales_row, end_col))
    reported_sales = to_float(snapshot.get(reported_sales_row, end_col))
    growth_rate_pct = to_float(snapshot.get(growth_rate_row, end_col))
    sales_captured_pct = to_float(snapshot.get(captured_row, end_col))
    range_width = (
        forecast_max - forecast_min
        if forecast_max is not None and forecast_min is not None
        else None
    )

    rows: list[dict[str, Any]] = []
    for idx, (n_quarters, start_col) in enumerate(row_specs):
        calc_row = calc_values[idx] if idx < len(calc_values) else []
        avg_penetration = to_float(calc_row[0] if len(calc_row) > 0 else None)
        forecast_value = to_float(calc_row[1] if len(calc_row) > 1 else None)
        last_quarter_used = snapshot.get(quarter_label_row, start_col)

        rows.append(
            {
                "model": metadata.model,
                "ticker": metadata.ticker,
                "model_period": metadata.model_period,
                "model_date": metadata.model_date,
                "method": "empirical",
                "parameter_name": "avg_penetration_pct",
                "parameter_value": avg_penetration,
                "num_quarters_used": n_quarters,
                "last_quarter_used": last_quarter_used,
                "forecast_value": forecast_value,
                "actual_value": reported_sales,
                "forecast_max": forecast_max,
                "forecast_min": forecast_min,
                "range_width": range_width,
                "avg_penetration_pct": avg_penetration,
                "quarterly_sales": quarterly_sales,
                "reported_sales": reported_sales,
                "growth_rate_pct": growth_rate_pct,
                "sales_captured_in_db_pct": sales_captured_pct,
                "source_file": source_file,
            }
        )

    return rows


def collect_recent_xy_rows(
    snapshot: SheetSnapshot,
    x_col: int,
    y_col: int,
    anchor_row: int,
    limit: int = 50,
) -> list[int]:
    rows: list[int] = []

    row = anchor_row - 1
    while row >= snapshot.row0:
        x_val = snapshot.get(row, x_col)
        y_val = snapshot.get(row, y_col)
        if is_number(x_val) and is_number(y_val):
            rows.append(row)
            if len(rows) >= limit:
                break
        elif rows:
            break
        row -= 1

    rows.reverse()
    if len(rows) >= 2:
        return rows

    fallback: list[int] = []
    for row in range(snapshot.row0, anchor_row):
        x_val = snapshot.get(row, x_col)
        y_val = snapshot.get(row, y_col)
        if is_number(x_val) and is_number(y_val):
            fallback.append(row)
    return fallback[-limit:]


def is_duplicate_regression_row(previous: dict[str, Any], current: dict[str, Any]) -> bool:
    fields = ["forecast_value", "intercept", "slope", "forecast_max", "forecast_min"]
    return all(
        approx_equal(to_float(previous.get(field)), to_float(current.get(field)))
        for field in fields
    )


def extract_regression_rows(workbook: xw.Book, metadata: FileMetadata, source_file: str) -> list[dict[str, Any]]:
    sheet = get_sheet_by_name(workbook, "Regression Model")
    if sheet is None:
        print(f"SKIP {source_file}: missing Regression Model sheet")
        return []

    snapshot = SheetSnapshot.from_sheet(sheet)
    anchor = find_anchor_max(snapshot)
    if anchor is None:
        print(f"SKIP {source_file}: Regression Model max anchor not found")
        return []

    anchor_row, anchor_col = anchor
    y_col = anchor_col - 7
    x_col = anchor_col - 11

    xy_rows = collect_recent_xy_rows(snapshot, x_col=x_col, y_col=y_col, anchor_row=anchor_row)
    if len(xy_rows) < 2:
        print(f"SKIP {source_file}: Regression Model has insufficient x/y values")
        return []

    max_quarters = min(N_QUARTERS, len(xy_rows))
    calc_start_row = max(anchor_row + 2, snapshot.row0 + snapshot.n_rows + 2)
    calc_start_col = max(anchor_col + 2, snapshot.col0 + snapshot.n_cols + 2)

    forecast_x = to_float(snapshot.get(anchor_row, x_col))
    if forecast_x is None:
        forecast_x = to_float(snapshot.get(xy_rows[-1], x_col))

    row_specs: list[tuple[int, int, int]] = []
    formulas: list[list[str]] = []
    for n_quarters in range(2, max_quarters + 1):
        start_row = xy_rows[-n_quarters]
        end_row = xy_rows[-1]
        calc_row = calc_start_row + len(formulas)

        y_range = r1c1_range(start_row, y_col, end_row, y_col)
        x_range = r1c1_range(start_row, x_col, end_row, x_col)
        intercept_formula = f"=INTERCEPT({y_range},{x_range})"
        slope_formula = f"=SLOPE({y_range},{x_range})"
        if forecast_x is None:
            forecast_formula = f"=R{calc_row}C{calc_start_col}+R{calc_row}C{calc_start_col + 1}*R{end_row}C{x_col}"
        else:
            forecast_formula = f"=R{calc_row}C{calc_start_col}+R{calc_row}C{calc_start_col + 1}*{forecast_x}"

        row_specs.append((n_quarters, start_row, end_row))
        formulas.append([intercept_formula, slope_formula, forecast_formula])

    if not formulas:
        return []

    calc_end_row = calc_start_row + len(formulas) - 1
    calc_range = sheet.range((calc_start_row, calc_start_col), (calc_end_row, calc_start_col + 2))
    calc_range.formula2 = formulas
    workbook.app.calculate()
    calc_values = to_matrix(calc_range.value)
    calc_range.clear_contents()

    forecast_max = find_numeric_near_label(snapshot, ["max"], near=anchor)
    forecast_min = find_numeric_near_label(snapshot, ["min"], near=anchor)
    actual_value = find_numeric_near_label(snapshot, ["actual", "reported sales"], near=anchor)
    range_width = (
        forecast_max - forecast_min
        if forecast_max is not None and forecast_min is not None
        else None
    )

    rows: list[dict[str, Any]] = []
    for idx, (n_quarters, _, _) in enumerate(row_specs):
        calc_row = calc_values[idx] if idx < len(calc_values) else []
        intercept = to_float(calc_row[0] if len(calc_row) > 0 else None)
        slope = to_float(calc_row[1] if len(calc_row) > 1 else None)
        forecast_value = to_float(calc_row[2] if len(calc_row) > 2 else None)

        candidate = {
            "model": metadata.model,
            "ticker": metadata.ticker,
            "model_period": metadata.model_period,
            "model_date": metadata.model_date,
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

        if rows and is_duplicate_regression_row(rows[-1], candidate):
            continue
        rows.append(candidate)

    return rows


def write_output_sheet(
    worksheet,
    columns: list[str],
    rows: list[dict[str, Any]],
) -> None:
    worksheet.append(columns)
    for row in rows:
        worksheet.append([row.get(column, None) for column in columns])

    for column_index in range(1, len(columns) + 1):
        worksheet.cell(row=1, column=column_index).font = Font(bold=True)

    worksheet.freeze_panes = "A2"
    end_col = get_column_letter(len(columns))
    worksheet.auto_filter.ref = f"A1:{end_col}{max(worksheet.max_row, 1)}"

    for column_index, column_name in enumerate(columns, start=1):
        max_len = len(column_name)
        for row_index in range(2, worksheet.max_row + 1):
            cell_value = worksheet.cell(row=row_index, column=column_index).value
            if cell_value is not None:
                max_len = max(max_len, len(str(cell_value)))
        worksheet.column_dimensions[get_column_letter(column_index)].width = min(max_len + 2, 48)


def build_output_workbook(
    output_path: Path,
    empirical_rows: list[dict[str, Any]],
    regression_rows: list[dict[str, Any]],
) -> None:
    workbook = Workbook()
    empirical_sheet = workbook.active
    empirical_sheet.title = "empirical_candidates"
    write_output_sheet(empirical_sheet, EMPIRICAL_COLUMNS, empirical_rows)

    regression_sheet = workbook.create_sheet("regression_candidates")
    write_output_sheet(regression_sheet, REGRESSION_COLUMNS, regression_rows)
    workbook.save(output_path)


def main() -> int:
    in_dir = Path(input_dir).expanduser().resolve()
    out_dir = Path(output_dir).expanduser().resolve()

    if not in_dir.exists():
        print(f"Input directory does not exist: {in_dir}")
        return 1

    out_dir.mkdir(parents=True, exist_ok=True)
    output_path = unique_output_path(in_dir, out_dir)
    output_stem_pattern = re.compile(rf"^{re.escape(in_dir.name)}_PARAM(?:\.\d+)?$", flags=re.IGNORECASE)

    files_processed = 0
    empirical_rows: list[dict[str, Any]] = []
    regression_rows: list[dict[str, Any]] = []

    app: xw.App | None = None
    try:
        app = xw.App(visible=False, add_book=False)
        app.display_alerts = False
        app.screen_updating = False
        app.enable_events = False
        try:
            app.calculation = "manual"
        except Exception:
            pass

        for entry in sorted(in_dir.iterdir(), key=lambda path: path.name.lower()):
            if not entry.is_file():
                print(f"SKIP {entry.name}: not a file")
                continue
            if entry.name.startswith("~"):
                print(f"SKIP {entry.name}: temporary workbook")
                continue
            if entry.suffix.lower() != ".xlsx":
                print(f"SKIP {entry.name}: not an .xlsx file")
                continue
            if output_stem_pattern.match(entry.stem):
                print(f"SKIP {entry.name}: prior PARAM output file")
                continue

            metadata = parse_file_metadata(entry)
            workbook: xw.Book | None = None
            try:
                workbook = app.books.open(str(entry), update_links=False)
                extracted_empirical = extract_empirical_rows(workbook, metadata, entry.name)
                extracted_regression = extract_regression_rows(workbook, metadata, entry.name)
                empirical_rows.extend(extracted_empirical)
                regression_rows.extend(extracted_regression)
                files_processed += 1
                print(
                    f"PROCESSED {entry.name}: "
                    f"empirical={len(extracted_empirical)}, "
                    f"regression={len(extracted_regression)}"
                )
            except Exception as exc:
                print(f"SKIP {entry.name}: {exc}")
            finally:
                if workbook is not None:
                    safe_close_workbook(workbook)

    finally:
        if app is not None:
            try:
                app.quit()
            except Exception:
                pass

    build_output_workbook(output_path, empirical_rows, regression_rows)
    print(f"OUTPUT {output_path}")
    print(f"FILES PROCESSED {files_processed}")
    print(f"EMPIRICAL ROWS {len(empirical_rows)}")
    print(f"REGRESSION ROWS {len(regression_rows)}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
