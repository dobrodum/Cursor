#!/usr/bin/env python3
"""Extract empirical/regression model candidates from .xlsx workbooks.

This script opens each source workbook once, processes both model sheets while
the workbook is open, and writes one consolidated output workbook with:
  - empirical_candidates
  - regression_candidates
"""

from __future__ import annotations

import re
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Iterable

import xlwings as xw
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

# ----------------------------
# User-configurable directories
# ----------------------------
input_dir = "/path/to/input"
output_dir = "/path/to/output"


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


# Fallback offsets if headers are not available/clean.
EMPIRICAL_OFFSETS = {
    "num_quarters_used": -6,
    "last_quarter_used": -5,
    "avg_penetration_pct": -4,
    "forecast_value": -3,
    "actual_value": -2,
    "forecast_max": 0,
    "forecast_min": 1,
    "quarterly_sales": -8,
    "reported_sales": -2,
    "growth_rate_pct": -7,
    "sales_captured_in_db_pct": -1,
}

REGRESSION_OFFSETS = {
    "num_quarters_used": -12,
    "forecast_value": -1,  # TOT FCST w/o SA
    "forecast_max": 0,
    "forecast_min": 1,
    "actual_value": -2,
}

N_QUARTERS = 10


@dataclass
class FileLabels:
    model: str
    ticker: str
    model_period: str
    model_date: str


def normalize_text(value: Any) -> str:
    if value is None:
        return ""
    text = str(value).strip().lower()
    text = re.sub(r"[^a-z0-9]+", " ", text)
    return " ".join(text.split())


def to_2d(values: Any) -> list[list[Any]]:
    if values is None:
        return []
    if not isinstance(values, list):
        return [[values]]
    if not values:
        return []
    if isinstance(values[0], list):
        return values
    return [values]


def parse_model_labels(file_path: Path) -> FileLabels:
    stem = file_path.stem
    parts = [segment.strip() for segment in stem.split(" - ")]
    ticker = ""
    if len(parts) >= 2:
        ticker = re.sub(r"\s+", "", parts[1]).upper()

    period_token = ""
    if len(parts) >= 3:
        period_token = parts[2].split("_")[0].strip()

    # Example: MidJan2026 -> MidJan_2026 and 2026-01-15
    match = re.match(r"^(Early|Mid|Late)([A-Za-z]{3,9})(\d{4})$", period_token, re.IGNORECASE)
    month_lookup = {
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
    day_lookup = {"early": 5, "mid": 15, "late": 25}

    model_period = period_token or "unknown_period"
    model_date = ""

    if match:
        timing = match.group(1).title()
        month_text = match.group(2)
        year = match.group(3)
        month_key = month_text[:3].lower()
        month_num = month_lookup.get(month_key)
        if month_num is not None:
            model_period = f"{timing}{month_text[:3].title()}_{year}"
            model_date = f"{year}-{month_num:02d}-{day_lookup[timing.lower()]:02d}"

    if not ticker:
        ticker = "UNKNOWN"
    model = f"{ticker}_{model_period}" if model_period else ticker
    return FileLabels(model=model, ticker=ticker, model_period=model_period, model_date=model_date)


def get_output_path(input_path: Path, output_path: Path) -> Path:
    output_path.mkdir(parents=True, exist_ok=True)
    base_name = f"{input_path.name}_PARAM.xlsx"
    candidate = output_path / base_name
    if not candidate.exists():
        return candidate

    i = 1
    while True:
        candidate = output_path / f"{input_path.name}_PARAM.{i}.xlsx"
        if not candidate.exists():
            return candidate
        i += 1


def list_source_files(input_path: Path) -> list[Path]:
    files: list[Path] = []
    for path in sorted(input_path.iterdir(), key=lambda p: p.name.lower()):
        if not path.is_file():
            continue
        if path.name.startswith("~"):
            print(f"SKIP {path.name}: temp file")
            continue
        if path.suffix.lower() != ".xlsx":
            print(f"SKIP {path.name}: not .xlsx")
            continue
        files.append(path)
    return files


def find_sheet_case_insensitive(wb: xw.Book, target_name: str) -> xw.Sheet | None:
    target_normalized = normalize_text(target_name)
    for sheet in wb.sheets:
        if normalize_text(sheet.name) == target_normalized:
            return sheet
    return None


def find_anchor(sheet: xw.Sheet, anchor_text: str = "max") -> tuple[int, int] | None:
    used = sheet.used_range
    values = to_2d(used.value)
    if not values:
        return None

    for r_offset, row in enumerate(values):
        for c_offset, cell_value in enumerate(row):
            if normalize_text(cell_value) == anchor_text:
                return used.row + r_offset, used.column + c_offset
    return None


def build_header_lookup(
    sheet: xw.Sheet, header_row: int, anchor_col: int, window: int = 30
) -> list[tuple[int, str]]:
    start_col = max(1, anchor_col - window)
    end_col = anchor_col + window
    row_values = sheet.range((header_row, start_col), (header_row, end_col)).value
    if not isinstance(row_values, list):
        row_values = [row_values]

    headers: list[tuple[int, str]] = []
    for i, value in enumerate(row_values):
        headers.append((start_col + i, normalize_text(value)))
    return headers


def resolve_column(
    headers: list[tuple[int, str]],
    anchor_col: int,
    fallback_offset: int,
    keyword_sets: Iterable[tuple[str, ...]],
) -> int:
    matches: list[int] = []
    for col, header in headers:
        for keyword_set in keyword_sets:
            if all(word in header for word in keyword_set):
                matches.append(col)
                break

    if matches:
        return min(matches, key=lambda col: abs(col - anchor_col))
    return max(1, anchor_col + fallback_offset)


def safe_close_workbook_no_save(wb: xw.Book) -> None:
    try:
        wb.close(save=False)
        return
    except TypeError:
        pass
    except Exception:
        pass

    # Fallback for engines that do not support keyword args.
    try:
        wb.api.Close(SaveChanges=False)
    except Exception:
        try:
            wb.close()
        except Exception:
            pass


def numeric(value: Any) -> float | None:
    if value is None or value == "":
        return None
    if isinstance(value, (int, float)):
        return float(value)
    text = str(value).replace(",", "").strip()
    if not text:
        return None
    try:
        return float(text)
    except ValueError:
        return None


def extract_empirical_rows(wb: xw.Book, sheet: xw.Sheet, labels: FileLabels, source_file: str) -> list[dict[str, Any]]:
    anchor = find_anchor(sheet, "max")
    if anchor is None:
        print(f"SKIP {source_file} / Empirical Model: max anchor not found")
        return []

    anchor_row, anchor_col = anchor
    data_start_row = anchor_row + 1
    headers = build_header_lookup(sheet, anchor_row, anchor_col)

    col_map = {
        "num_quarters_used": resolve_column(
            headers,
            anchor_col,
            EMPIRICAL_OFFSETS["num_quarters_used"],
            (("num", "quarter"), ("n", "quarter"), ("quarters", "used")),
        ),
        "last_quarter_used": resolve_column(
            headers,
            anchor_col,
            EMPIRICAL_OFFSETS["last_quarter_used"],
            (("last", "quarter"),),
        ),
        "avg_penetration_pct": resolve_column(
            headers,
            anchor_col,
            EMPIRICAL_OFFSETS["avg_penetration_pct"],
            (("avg", "penetration"),),
        ),
        "forecast_value": resolve_column(
            headers,
            anchor_col,
            EMPIRICAL_OFFSETS["forecast_value"],
            (
                ("estimated", "total", "sold"),
                ("forecast", "value"),
                ("total", "forecast"),
            ),
        ),
        "actual_value": resolve_column(
            headers,
            anchor_col,
            EMPIRICAL_OFFSETS["actual_value"],
            (("reported", "sales"), ("actual", "value"), ("actual", "sales")),
        ),
        "forecast_max": anchor_col,
        "forecast_min": resolve_column(
            headers,
            anchor_col,
            EMPIRICAL_OFFSETS["forecast_min"],
            (("min",),),
        ),
        "quarterly_sales": resolve_column(
            headers,
            anchor_col,
            EMPIRICAL_OFFSETS["quarterly_sales"],
            (("quarterly", "sales"), ("qtr", "sales")),
        ),
        "reported_sales": resolve_column(
            headers,
            anchor_col,
            EMPIRICAL_OFFSETS["reported_sales"],
            (("reported", "sales"),),
        ),
        "growth_rate_pct": resolve_column(
            headers,
            anchor_col,
            EMPIRICAL_OFFSETS["growth_rate_pct"],
            (("growth", "rate"),),
        ),
        "sales_captured_in_db_pct": resolve_column(
            headers,
            anchor_col,
            EMPIRICAL_OFFSETS["sales_captured_in_db_pct"],
            (("captured", "db"), ("sales", "captured"), ("captured", "in", "db")),
        ),
    }

    used_last_col = sheet.used_range.last_cell.column
    temp_avg_col = used_last_col + 2
    temp_rows: list[int] = []

    raw_rows: list[dict[str, Any]] = []
    source_penetration_col = col_map["sales_captured_in_db_pct"]

    for i in range(N_QUARTERS):
        row_idx = data_start_row + i
        num_quarters_used = sheet.cells(row_idx, col_map["num_quarters_used"]).value
        num_quarters_num = int(numeric(num_quarters_used) or (i + 1))

        # Use temporary R1C1 formula to compute average penetration.
        avg_pen_formula = (
            f'=IFERROR(AVERAGE(R{max(1, row_idx - num_quarters_num + 1)}C{source_penetration_col}'
            f":R{row_idx}C{source_penetration_col}),\"\")"
        )
        sheet.cells(row_idx, temp_avg_col).formula2 = avg_pen_formula
        temp_rows.append(row_idx)

        raw_rows.append(
            {
                "row_idx": row_idx,
                "num_quarters_used": num_quarters_used,
                "last_quarter_used": sheet.cells(row_idx, col_map["last_quarter_used"]).value,
                "avg_penetration_existing": sheet.cells(row_idx, col_map["avg_penetration_pct"]).value,
                "forecast_value": sheet.cells(row_idx, col_map["forecast_value"]).value,
                "actual_value": sheet.cells(row_idx, col_map["actual_value"]).value,
                "forecast_max": sheet.cells(row_idx, col_map["forecast_max"]).value,
                "forecast_min": sheet.cells(row_idx, col_map["forecast_min"]).value,
                "quarterly_sales": sheet.cells(row_idx, col_map["quarterly_sales"]).value,
                "reported_sales": sheet.cells(row_idx, col_map["reported_sales"]).value,
                "growth_rate_pct": sheet.cells(row_idx, col_map["growth_rate_pct"]).value,
                "sales_captured_in_db_pct": sheet.cells(row_idx, col_map["sales_captured_in_db_pct"]).value,
            }
        )

    if temp_rows:
        wb.app.calculate()

    rows: list[dict[str, Any]] = []
    for raw in raw_rows:
        row_idx = raw["row_idx"]
        avg_penetration_formula_value = sheet.cells(row_idx, temp_avg_col).value
        avg_penetration_pct = (
            avg_penetration_formula_value
            if avg_penetration_formula_value not in (None, "")
            else raw["avg_penetration_existing"]
        )

        forecast_max_num = numeric(raw["forecast_max"])
        forecast_min_num = numeric(raw["forecast_min"])
        range_width = None
        if forecast_max_num is not None and forecast_min_num is not None:
            range_width = forecast_max_num - forecast_min_num

        row_payload = {
            "model": labels.model,
            "ticker": labels.ticker,
            "model_period": labels.model_period,
            "model_date": labels.model_date,
            "method": "empirical",
            "parameter_name": "avg_penetration_pct",
            "parameter_value": avg_penetration_pct,
            "num_quarters_used": raw["num_quarters_used"],
            "last_quarter_used": raw["last_quarter_used"],
            "forecast_value": raw["forecast_value"],
            "actual_value": raw["actual_value"],
            "forecast_max": raw["forecast_max"],
            "forecast_min": raw["forecast_min"],
            "range_width": range_width,
            "avg_penetration_pct": avg_penetration_pct,
            "quarterly_sales": raw["quarterly_sales"],
            "reported_sales": raw["reported_sales"],
            "growth_rate_pct": raw["growth_rate_pct"],
            "sales_captured_in_db_pct": raw["sales_captured_in_db_pct"],
            "source_file": source_file,
        }

        key_values = (
            row_payload["num_quarters_used"],
            row_payload["forecast_value"],
            row_payload["forecast_max"],
            row_payload["forecast_min"],
            row_payload["avg_penetration_pct"],
        )
        if any(v not in (None, "") for v in key_values):
            rows.append(row_payload)

    # Clear temporary formula cells.
    if temp_rows:
        sheet.range((temp_rows[0], temp_avg_col), (temp_rows[-1], temp_avg_col)).clear_contents()

    return rows


def extract_regression_rows(wb: xw.Book, sheet: xw.Sheet, labels: FileLabels, source_file: str) -> list[dict[str, Any]]:
    anchor = find_anchor(sheet, "max")
    if anchor is None:
        print(f"SKIP {source_file} / Regression Model: max anchor not found")
        return []

    anchor_row, anchor_col = anchor
    data_start_row = anchor_row + 1
    headers = build_header_lookup(sheet, anchor_row, anchor_col)

    x_col = max(1, anchor_col - 11)
    y_col = max(1, anchor_col - 7)

    col_map = {
        "num_quarters_used": resolve_column(
            headers,
            anchor_col,
            REGRESSION_OFFSETS["num_quarters_used"],
            (("num", "quarter"), ("n", "quarter"), ("quarters", "used")),
        ),
        "forecast_value": resolve_column(
            headers,
            anchor_col,
            REGRESSION_OFFSETS["forecast_value"],
            (
                ("tot", "fcst", "w", "o", "sa"),
                ("total", "forecast", "without", "sa"),
                ("forecast", "without", "sa"),
            ),
        ),
        "forecast_max": anchor_col,
        "forecast_min": resolve_column(
            headers,
            anchor_col,
            REGRESSION_OFFSETS["forecast_min"],
            (("min",),),
        ),
        "actual_value": resolve_column(
            headers,
            anchor_col,
            REGRESSION_OFFSETS["actual_value"],
            (("actual", "value"), ("actual", "sales"), ("reported", "sales")),
        ),
    }

    used_last_col = sheet.used_range.last_cell.column
    temp_intercept_col = used_last_col + 2
    temp_slope_col = used_last_col + 3

    raw_rows: list[dict[str, Any]] = []
    formula_rows: list[int] = []

    for i in range(N_QUARTERS):
        row_idx = data_start_row + i
        num_quarters_raw = sheet.cells(row_idx, col_map["num_quarters_used"]).value
        num_quarters = int(numeric(num_quarters_raw) or (i + 1))

        start_row = max(1, row_idx - num_quarters + 1)
        intercept_formula = (
            f'=IFERROR(INTERCEPT(R{start_row}C{y_col}:R{row_idx}C{y_col},'
            f'R{start_row}C{x_col}:R{row_idx}C{x_col}),"")'
        )
        slope_formula = (
            f'=IFERROR(SLOPE(R{start_row}C{y_col}:R{row_idx}C{y_col},'
            f'R{start_row}C{x_col}:R{row_idx}C{x_col}),"")'
        )
        sheet.cells(row_idx, temp_intercept_col).formula2 = intercept_formula
        sheet.cells(row_idx, temp_slope_col).formula2 = slope_formula
        formula_rows.append(row_idx)

        raw_rows.append(
            {
                "row_idx": row_idx,
                "num_quarters_used": num_quarters_raw,
                "forecast_value": sheet.cells(row_idx, col_map["forecast_value"]).value,
                "actual_value": sheet.cells(row_idx, col_map["actual_value"]).value,
                "forecast_max": sheet.cells(row_idx, col_map["forecast_max"]).value,
                "forecast_min": sheet.cells(row_idx, col_map["forecast_min"]).value,
            }
        )

    if formula_rows:
        wb.app.calculate()

    rows: list[dict[str, Any]] = []
    prev_signature: tuple[Any, ...] | None = None

    for raw in raw_rows:
        row_idx = raw["row_idx"]
        intercept = sheet.cells(row_idx, temp_intercept_col).value
        slope = sheet.cells(row_idx, temp_slope_col).value

        forecast_max_num = numeric(raw["forecast_max"])
        forecast_min_num = numeric(raw["forecast_min"])
        range_width = None
        if forecast_max_num is not None and forecast_min_num is not None:
            range_width = forecast_max_num - forecast_min_num

        payload = {
            "model": labels.model,
            "ticker": labels.ticker,
            "model_period": labels.model_period,
            "model_date": labels.model_date,
            "method": "regression",
            "parameter_name": "num_quarters_used",
            "parameter_value": raw["num_quarters_used"],
            "num_quarters_used": raw["num_quarters_used"],
            "forecast_value": raw["forecast_value"],
            "actual_value": raw["actual_value"] if raw["actual_value"] not in (None, "") else "",
            "forecast_max": raw["forecast_max"],
            "forecast_min": raw["forecast_min"],
            "range_width": range_width,
            "intercept": intercept,
            "slope": slope,
            "source_file": source_file,
        }

        signature = (
            payload["num_quarters_used"],
            payload["forecast_value"],
            payload["forecast_max"],
            payload["forecast_min"],
            payload["intercept"],
            payload["slope"],
        )
        if signature == prev_signature:
            continue
        prev_signature = signature

        if any(value not in (None, "") for value in signature):
            rows.append(payload)

    if formula_rows:
        sheet.range(
            (formula_rows[0], temp_intercept_col), (formula_rows[-1], temp_slope_col)
        ).clear_contents()

    return rows


def write_output_sheet(wb: Workbook, sheet_name: str, columns: list[str], rows: list[dict[str, Any]]) -> None:
    ws = wb.create_sheet(sheet_name)
    ws.append(columns)
    ws.freeze_panes = "A2"

    for row in rows:
        ws.append([row.get(column) for column in columns])

    for cell in ws[1]:
        cell.font = Font(bold=True)

    ws.auto_filter.ref = ws.dimensions

    for idx, col_name in enumerate(columns, start=1):
        max_len = len(col_name)
        for row_idx in range(2, ws.max_row + 1):
            value = ws.cell(row=row_idx, column=idx).value
            if value is None:
                continue
            max_len = max(max_len, len(str(value)))
        ws.column_dimensions[get_column_letter(idx)].width = min(60, max(12, max_len + 2))


def main() -> None:
    input_path = Path(input_dir)
    output_path = Path(output_dir)

    if not input_path.exists() or not input_path.is_dir():
        raise FileNotFoundError(f"Input directory does not exist or is not a directory: {input_path}")

    source_files = list_source_files(input_path)
    if not source_files:
        print("No source .xlsx files found.")
        return

    out_file = get_output_path(input_path, output_path)

    empirical_rows: list[dict[str, Any]] = []
    regression_rows: list[dict[str, Any]] = []
    processed_files = 0

    app = xw.App(visible=False, add_book=False)
    app.display_alerts = False
    app.screen_updating = False

    try:
        try:
            app.calculation = "manual"
        except Exception:
            pass

        for file_path in source_files:
            wb: xw.Book | None = None
            try:
                wb = app.books.open(str(file_path), update_links=False)
                labels = parse_model_labels(file_path)

                empirical_sheet = find_sheet_case_insensitive(wb, "Empirical Model")
                if empirical_sheet is None:
                    print(f"SKIP {file_path.name} / Empirical Model: sheet not found")
                else:
                    empirical_rows.extend(
                        extract_empirical_rows(
                            wb=wb,
                            sheet=empirical_sheet,
                            labels=labels,
                            source_file=file_path.name,
                        )
                    )

                regression_sheet = find_sheet_case_insensitive(wb, "Regression Model")
                if regression_sheet is None:
                    print(f"SKIP {file_path.name} / Regression Model: sheet not found")
                else:
                    regression_rows.extend(
                        extract_regression_rows(
                            wb=wb,
                            sheet=regression_sheet,
                            labels=labels,
                            source_file=file_path.name,
                        )
                    )

                processed_files += 1
                print(f"PROCESSED {file_path.name}")
            except Exception as exc:
                print(f"SKIP {file_path.name}: {exc}")
            finally:
                if wb is not None:
                    safe_close_workbook_no_save(wb)
    finally:
        try:
            app.quit()
        except Exception:
            pass

    out_wb = Workbook()
    default_sheet = out_wb.active
    out_wb.remove(default_sheet)

    write_output_sheet(out_wb, "empirical_candidates", EMPIRICAL_COLUMNS, empirical_rows)
    write_output_sheet(out_wb, "regression_candidates", REGRESSION_COLUMNS, regression_rows)
    out_wb.save(out_file)

    print(f"OUTPUT {out_file}")
    print(f"FILES_PROCESSED {processed_files}")
    print(f"EMPIRICAL_ROWS {len(empirical_rows)}")
    print(f"REGRESSION_ROWS {len(regression_rows)}")


if __name__ == "__main__":
    main()
