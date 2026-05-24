#!/usr/bin/env python3
"""Extract empirical and regression candidate rows from Excel model files.

The script opens each source workbook only once, processes both model sheets
while open, and writes one combined output workbook.
"""

from __future__ import annotations

import re
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Any

try:
    import xlwings as xw
except ImportError as exc:  # pragma: no cover - import guard
    raise SystemExit("xlwings is required. Install with: pip install xlwings") from exc

try:
    from openpyxl import Workbook
    from openpyxl.styles import Font
    from openpyxl.utils import get_column_letter
except ImportError as exc:  # pragma: no cover - import guard
    raise SystemExit("openpyxl is required. Install with: pip install openpyxl") from exc


# Configure these two paths before running.
input_dir = Path("./input")
output_dir = Path("./output")


EMPIRICAL_SHEET_NAME = "empirical_candidates"
REGRESSION_SHEET_NAME = "regression_candidates"
EMPIRICAL_MODEL_SHEET = "Empirical Model"
REGRESSION_MODEL_SHEET = "Regression Model"

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

EMPIRICAL_ALIASES = {
    "num_quarters_used": ["num quarters used", "num qtrs used", "quarters used"],
    "last_quarter_used": ["last quarter used", "last qtr used", "last quarter"],
    "forecast_value": ["estimated total sold", "est total sold", "forecast value", "tot fcst"],
    "actual_value": ["reported sales", "actual sales", "actual value", "actual"],
    "forecast_max": ["max", "forecast max"],
    "forecast_min": ["min", "forecast min"],
    "quarterly_sales": ["quarterly sales", "quarter sales"],
    "reported_sales": ["reported sales", "reported total sales"],
    "growth_rate_pct": ["growth rate pct", "growth rate", "growth pct"],
    "sales_captured_in_db_pct": [
        "sales captured in db pct",
        "sales captured in db",
        "captured in db pct",
    ],
    "avg_penetration_pct": ["avg penetration pct", "average penetration pct", "avg penetration"],
}

REGRESSION_ALIASES = {
    "num_quarters_used": ["num quarters used", "num qtrs used", "quarters used"],
    "forecast_value": [
        "tot fcst w/o sa",
        "tot fcst wo sa",
        "tot fcst without sa",
        "forecast total without sa",
        "forecast value",
    ],
    "actual_value": ["actual", "actual value", "reported sales", "actual sales"],
    "forecast_max": ["max", "forecast max"],
    "forecast_min": ["min", "forecast min"],
}

# Fallback offsets are relative to the 'max' anchor column.
EMPIRICAL_FALLBACK_OFFSETS = {
    "num_quarters_used": -8,
    "last_quarter_used": -7,
    "forecast_value": -1,
    "actual_value": 2,
    "forecast_max": 0,
    "forecast_min": 1,
    "quarterly_sales": -6,
    "reported_sales": -5,
    "growth_rate_pct": -4,
    "sales_captured_in_db_pct": -3,
    "avg_penetration_pct": -2,
}

REGRESSION_FALLBACK_OFFSETS = {
    "num_quarters_used": -2,
    "forecast_value": -1,
    "actual_value": 2,
    "forecast_max": 0,
    "forecast_min": 1,
}


@dataclass
class SheetSnapshot:
    top_row: int
    left_col: int
    values: list[list[Any]]

    @property
    def n_rows(self) -> int:
        return len(self.values)

    @property
    def n_cols(self) -> int:
        return len(self.values[0]) if self.values else 0

    def value_at(self, abs_row: int, abs_col: int) -> Any:
        row_idx = abs_row - self.top_row
        col_idx = abs_col - self.left_col
        if row_idx < 0 or col_idx < 0:
            return None
        if row_idx >= self.n_rows or col_idx >= self.n_cols:
            return None
        return self.values[row_idx][col_idx]


def normalize_text(value: Any) -> str:
    if value is None:
        return ""
    text = str(value).strip().lower()
    text = re.sub(r"[^a-z0-9]+", " ", text)
    return re.sub(r"\s+", " ", text).strip()


def to_2d(value: Any) -> list[list[Any]]:
    if value is None:
        return []
    if isinstance(value, (list, tuple)):
        if not value:
            return []
        first_item = value[0]
        if isinstance(first_item, (list, tuple)):
            return [list(row) if isinstance(row, tuple) else row for row in value]
        return [list(value) if isinstance(value, tuple) else value]
    return [[value]]


def to_float(value: Any) -> float | None:
    if value in (None, ""):
        return None
    try:
        return float(value)
    except (TypeError, ValueError):
        return None


def is_blank(value: Any) -> bool:
    return value is None or value == ""


def range_width(max_value: Any, min_value: Any) -> float | None:
    max_float = to_float(max_value)
    min_float = to_float(min_value)
    if max_float is None or min_float is None:
        return None
    return max_float - min_float


def snapshot_sheet(sheet: xw.Sheet) -> SheetSnapshot:
    used = sheet.used_range
    values = to_2d(used.value)
    return SheetSnapshot(top_row=used.row, left_col=used.column, values=values)


def find_max_anchor(snapshot: SheetSnapshot) -> tuple[int, int] | None:
    for row_idx, row_values in enumerate(snapshot.values):
        for col_idx, cell_value in enumerate(row_values):
            if normalize_text(cell_value) == "max":
                return snapshot.top_row + row_idx, snapshot.left_col + col_idx
    return None


def build_header_map(snapshot: SheetSnapshot, header_row: int) -> dict[str, int]:
    row_idx = header_row - snapshot.top_row
    if row_idx < 0 or row_idx >= snapshot.n_rows:
        return {}

    mapping: dict[str, int] = {}
    for col_idx, raw_value in enumerate(snapshot.values[row_idx]):
        normalized = normalize_text(raw_value)
        if normalized and normalized not in mapping:
            mapping[normalized] = snapshot.left_col + col_idx
    return mapping


def resolve_column(
    header_map: dict[str, int],
    aliases: list[str],
    anchor_col: int,
    fallback_offset: int,
) -> int:
    normalized_aliases = [normalize_text(alias) for alias in aliases]

    for alias in normalized_aliases:
        if alias in header_map:
            return header_map[alias]

    for alias in normalized_aliases:
        for header_text, col_idx in header_map.items():
            if alias and alias in header_text:
                return col_idx

    return max(1, anchor_col + fallback_offset)


def get_sheet_if_exists(workbook: xw.Book, sheet_name: str) -> xw.Sheet | None:
    try:
        return workbook.sheets[sheet_name]
    except Exception:
        return None


def set_formula_r1c1(cell: xw.Range, formula: str) -> None:
    try:
        cell.formula2 = formula
    except Exception:
        cell.formula = formula


def safe_close_workbook(workbook: xw.Book) -> None:
    try:
        workbook.close(save=False)
        return
    except TypeError:
        pass
    except Exception:
        pass

    try:
        workbook.api.Close(SaveChanges=False)
        return
    except TypeError:
        try:
            workbook.api.Close(False)
            return
        except Exception:
            pass
    except Exception:
        pass

    try:
        workbook.api.Saved = True
    except Exception:
        pass
    workbook.close()


def safe_cell_value(sheet: xw.Sheet, row: int, col: int) -> Any:
    if row < 1 or col < 1:
        return None
    return sheet.cells(row, col).value


def parse_filename_metadata(file_name: str) -> dict[str, str]:
    stem = Path(file_name).stem
    parts = [part.strip() for part in stem.split(" - ")]
    ticker = parts[1].upper() if len(parts) >= 2 and parts[1] else "UNKNOWN"

    period_match = re.search(r"(Early|Mid|Late)\s*([A-Za-z]{3,9})\s*(\d{4})", stem, flags=re.IGNORECASE)
    model_period = "Unknown"
    model_date = ""

    if period_match:
        phase = period_match.group(1).title()
        month_token = period_match.group(2)[:3].title()
        year = int(period_match.group(3))
        day = {"Early": 5, "Mid": 15, "Late": 25}[phase]
        try:
            month_num = datetime.strptime(month_token, "%b").month
        except ValueError:
            month_num = 1
        model_period = f"{phase}{month_token}_{year}"
        model_date = f"{year:04d}-{month_num:02d}-{day:02d}"

    model = f"{ticker}_{model_period}" if model_period != "Unknown" else ticker
    return {
        "model": model,
        "ticker": ticker,
        "model_period": model_period,
        "model_date": model_date,
    }


def create_output_path(input_path: Path, output_path: Path) -> Path:
    base_name = f"{input_path.name}_PARAM"
    candidate = output_path / f"{base_name}.xlsx"
    suffix = 1
    while candidate.exists():
        candidate = output_path / f"{base_name}.{suffix}.xlsx"
        suffix += 1
    return candidate


def write_sheet(
    workbook: Workbook,
    sheet_name: str,
    columns: list[str],
    rows: list[dict[str, Any]],
) -> None:
    if sheet_name in workbook.sheetnames:
        sheet = workbook[sheet_name]
    else:
        sheet = workbook.create_sheet(sheet_name)

    sheet.append(columns)
    for row in rows:
        sheet.append([row.get(column) for column in columns])

    bold_font = Font(bold=True)
    for cell in sheet[1]:
        cell.font = bold_font

    sheet.freeze_panes = "A2"
    last_column = get_column_letter(len(columns))
    last_row = max(1, sheet.max_row)
    sheet.auto_filter.ref = f"A1:{last_column}{last_row}"

    for col_idx, column_name in enumerate(columns, start=1):
        max_len = len(column_name)
        for row_idx in range(2, sheet.max_row + 1):
            value = sheet.cell(row=row_idx, column=col_idx).value
            if value is None:
                continue
            max_len = max(max_len, len(str(value)))
        sheet.column_dimensions[get_column_letter(col_idx)].width = min(max(max_len + 2, 12), 48)


def rows_are_equal(previous_row: dict[str, Any], current_row: dict[str, Any], keys: list[str]) -> bool:
    for key in keys:
        left = previous_row.get(key)
        right = current_row.get(key)

        left_float = to_float(left)
        right_float = to_float(right)
        if left_float is not None and right_float is not None:
            if abs(left_float - right_float) > 1e-9:
                return False
            continue

        if is_blank(left) and is_blank(right):
            continue
        if str(left) != str(right):
            return False
    return True


def extract_empirical_rows(
    workbook: xw.Book,
    metadata: dict[str, str],
    source_file: str,
) -> list[dict[str, Any]]:
    sheet = get_sheet_if_exists(workbook, EMPIRICAL_MODEL_SHEET)
    if sheet is None:
        print(f"skipped empirical in {source_file}: missing sheet '{EMPIRICAL_MODEL_SHEET}'")
        return []

    snapshot = snapshot_sheet(sheet)
    anchor = find_max_anchor(snapshot)
    if anchor is None:
        print(f"skipped empirical in {source_file}: cannot find 'max' anchor")
        return []

    anchor_row, anchor_col = anchor
    header_map = build_header_map(snapshot, anchor_row)

    column_map = {
        key: resolve_column(
            header_map=header_map,
            aliases=EMPIRICAL_ALIASES[key],
            anchor_col=anchor_col,
            fallback_offset=EMPIRICAL_FALLBACK_OFFSETS[key],
        )
        for key in EMPIRICAL_ALIASES
    }

    helper_col = snapshot.left_col + snapshot.n_cols + 2
    helper_cell = sheet.cells(anchor_row, helper_col)
    data_end_row = anchor_row - 1
    rows: list[dict[str, Any]] = []

    for n_quarters in range(1, N_QUARTERS + 1):
        output_row = anchor_row + n_quarters

        avg_penetration = None
        data_start_row = data_end_row - n_quarters + 1
        if data_start_row >= snapshot.top_row:
            q_col = column_map["quarterly_sales"]
            r_col = column_map["reported_sales"]
            avg_formula = (
                f'=IFERROR(AVERAGE(IFERROR(R{data_start_row}C{q_col}:R{data_end_row}C{q_col}/'
                f'R{data_start_row}C{r_col}:R{data_end_row}C{r_col},"")),"")'
            )
            set_formula_r1c1(helper_cell, avg_formula)
            workbook.app.calculate()
            avg_penetration = helper_cell.value

        if is_blank(avg_penetration):
            avg_penetration = safe_cell_value(sheet, output_row, column_map["avg_penetration_pct"])

        num_quarters_used = safe_cell_value(sheet, output_row, column_map["num_quarters_used"])
        if is_blank(num_quarters_used):
            num_quarters_used = n_quarters

        last_quarter_used = safe_cell_value(sheet, output_row, column_map["last_quarter_used"])
        forecast_value = safe_cell_value(sheet, output_row, column_map["forecast_value"])
        actual_value = safe_cell_value(sheet, output_row, column_map["actual_value"])
        forecast_max = safe_cell_value(sheet, output_row, column_map["forecast_max"])
        forecast_min = safe_cell_value(sheet, output_row, column_map["forecast_min"])
        quarterly_sales = safe_cell_value(sheet, output_row, column_map["quarterly_sales"])
        reported_sales = safe_cell_value(sheet, output_row, column_map["reported_sales"])
        growth_rate_pct = safe_cell_value(sheet, output_row, column_map["growth_rate_pct"])
        sales_captured_in_db_pct = safe_cell_value(sheet, output_row, column_map["sales_captured_in_db_pct"])

        if all(
            is_blank(value)
            for value in [
                forecast_value,
                actual_value,
                forecast_max,
                forecast_min,
                avg_penetration,
            ]
        ):
            continue

        rows.append(
            {
                "model": metadata["model"],
                "ticker": metadata["ticker"],
                "model_period": metadata["model_period"],
                "model_date": metadata["model_date"],
                "method": "empirical",
                "parameter_name": "avg_penetration_pct",
                "parameter_value": avg_penetration,
                "num_quarters_used": num_quarters_used,
                "last_quarter_used": last_quarter_used,
                "forecast_value": forecast_value,
                "actual_value": actual_value,
                "forecast_max": forecast_max,
                "forecast_min": forecast_min,
                "range_width": range_width(forecast_max, forecast_min),
                "avg_penetration_pct": avg_penetration,
                "quarterly_sales": quarterly_sales,
                "reported_sales": reported_sales,
                "growth_rate_pct": growth_rate_pct,
                "sales_captured_in_db_pct": sales_captured_in_db_pct,
                "source_file": source_file,
            }
        )

    helper_cell.value = None
    return rows


def extract_regression_rows(
    workbook: xw.Book,
    metadata: dict[str, str],
    source_file: str,
) -> list[dict[str, Any]]:
    sheet = get_sheet_if_exists(workbook, REGRESSION_MODEL_SHEET)
    if sheet is None:
        print(f"skipped regression in {source_file}: missing sheet '{REGRESSION_MODEL_SHEET}'")
        return []

    snapshot = snapshot_sheet(sheet)
    anchor = find_max_anchor(snapshot)
    if anchor is None:
        print(f"skipped regression in {source_file}: cannot find 'max' anchor")
        return []

    anchor_row, anchor_col = anchor
    header_map = build_header_map(snapshot, anchor_row)

    column_map = {
        key: resolve_column(
            header_map=header_map,
            aliases=REGRESSION_ALIASES[key],
            anchor_col=anchor_col,
            fallback_offset=REGRESSION_FALLBACK_OFFSETS[key],
        )
        for key in REGRESSION_ALIASES
    }

    x_col = anchor_col - 11
    y_col = anchor_col - 7

    helper_base_col = snapshot.left_col + snapshot.n_cols + 2
    intercept_cell = sheet.cells(anchor_row, helper_base_col)
    slope_cell = sheet.cells(anchor_row, helper_base_col + 1)

    rows: list[dict[str, Any]] = []
    data_end_row = anchor_row - 1

    for n_quarters in range(1, N_QUARTERS + 1):
        data_start_row = data_end_row - n_quarters + 1
        if data_start_row < snapshot.top_row:
            break

        intercept_formula = (
            f'=IFERROR(INTERCEPT(R{data_start_row}C{y_col}:R{data_end_row}C{y_col},'
            f'R{data_start_row}C{x_col}:R{data_end_row}C{x_col}),"")'
        )
        slope_formula = (
            f'=IFERROR(SLOPE(R{data_start_row}C{y_col}:R{data_end_row}C{y_col},'
            f'R{data_start_row}C{x_col}:R{data_end_row}C{x_col}),"")'
        )
        set_formula_r1c1(intercept_cell, intercept_formula)
        set_formula_r1c1(slope_cell, slope_formula)
        workbook.app.calculate()

        intercept = intercept_cell.value
        slope = slope_cell.value

        output_row = anchor_row + n_quarters
        num_quarters_used = safe_cell_value(sheet, output_row, column_map["num_quarters_used"])
        if is_blank(num_quarters_used):
            num_quarters_used = n_quarters

        forecast_value = safe_cell_value(sheet, output_row, column_map["forecast_value"])
        actual_value = safe_cell_value(sheet, output_row, column_map["actual_value"])
        forecast_max = safe_cell_value(sheet, output_row, column_map["forecast_max"])
        forecast_min = safe_cell_value(sheet, output_row, column_map["forecast_min"])

        if is_blank(forecast_value):
            x_value = safe_cell_value(sheet, output_row, x_col)
            if to_float(intercept) is not None and to_float(slope) is not None and to_float(x_value) is not None:
                forecast_value = float(intercept) + float(slope) * float(x_value)

        if all(
            is_blank(value)
            for value in [forecast_value, forecast_max, forecast_min, intercept, slope]
        ):
            continue

        current_row = {
            "model": metadata["model"],
            "ticker": metadata["ticker"],
            "model_period": metadata["model_period"],
            "model_date": metadata["model_date"],
            "method": "regression",
            "parameter_name": "num_quarters_used",
            "parameter_value": num_quarters_used,
            "num_quarters_used": num_quarters_used,
            "forecast_value": forecast_value,
            "actual_value": actual_value,
            "forecast_max": forecast_max,
            "forecast_min": forecast_min,
            "range_width": range_width(forecast_max, forecast_min),
            "intercept": intercept,
            "slope": slope,
            "source_file": source_file,
        }

        if rows and rows_are_equal(
            rows[-1],
            current_row,
            keys=["forecast_value", "forecast_max", "forecast_min", "intercept", "slope"],
        ):
            continue

        rows.append(current_row)

    intercept_cell.value = None
    slope_cell.value = None
    return rows


def process_files(source_dir: Path) -> tuple[list[dict[str, Any]], list[dict[str, Any]], int]:
    empirical_rows: list[dict[str, Any]] = []
    regression_rows: list[dict[str, Any]] = []
    processed_files = 0

    excel_app = xw.App(visible=False, add_book=False)
    excel_app.display_alerts = False
    excel_app.screen_updating = False
    try:
        excel_app.calculation = "manual"
    except Exception:
        pass

    try:
        for file_path in sorted(source_dir.iterdir(), key=lambda path: path.name.lower()):
            if not file_path.is_file():
                print(f"skipped {file_path.name}: not a file")
                continue
            if file_path.name.startswith("~"):
                print(f"skipped {file_path.name}: temporary file")
                continue
            if file_path.suffix.lower() != ".xlsx":
                print(f"skipped {file_path.name}: not an .xlsx file")
                continue

            print(f"processing file: {file_path.name}")
            workbook = None
            try:
                workbook = excel_app.books.open(str(file_path), update_links=False)
                metadata = parse_filename_metadata(file_path.name)
                empirical_rows.extend(extract_empirical_rows(workbook, metadata, file_path.name))
                regression_rows.extend(extract_regression_rows(workbook, metadata, file_path.name))
                processed_files += 1
                print(f"processed file: {file_path.name}")
            except Exception as exc:
                print(f"skipped {file_path.name}: failed to process ({exc})")
            finally:
                if workbook is not None:
                    safe_close_workbook(workbook)
    finally:
        excel_app.quit()

    return empirical_rows, regression_rows, processed_files


def main() -> None:
    source_dir = input_dir.expanduser().resolve()
    target_dir = output_dir.expanduser().resolve()

    if not source_dir.exists():
        raise SystemExit(f"input_dir does not exist: {source_dir}")

    target_dir.mkdir(parents=True, exist_ok=True)

    empirical_rows, regression_rows, processed_files = process_files(source_dir)
    output_path = create_output_path(source_dir, target_dir)

    output_workbook = Workbook()
    default_sheet = output_workbook.active
    output_workbook.remove(default_sheet)

    write_sheet(output_workbook, EMPIRICAL_SHEET_NAME, EMPIRICAL_COLUMNS, empirical_rows)
    write_sheet(output_workbook, REGRESSION_SHEET_NAME, REGRESSION_COLUMNS, regression_rows)
    output_workbook.save(output_path)

    print(f"output path: {output_path}")
    print(f"number of files processed: {processed_files}")
    print(f"number of empirical rows: {len(empirical_rows)}")
    print(f"number of regression rows: {len(regression_rows)}")


if __name__ == "__main__":
    main()
