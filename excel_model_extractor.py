#!/usr/bin/env python3
from __future__ import annotations

import math
import re
import sys
from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path
from typing import Any

from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

try:
    import xlwings as xw
except ImportError as exc:  # pragma: no cover - runtime dependency guard
    raise SystemExit(
        "Missing dependency: xlwings. Install with `pip install xlwings`."
    ) from exc


# -----------------------------
# User-configurable paths
# -----------------------------
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

N_QUARTERS = 10
MONTHS = {
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
PERIOD_DAY = {"early": 5, "mid": 15, "late": 25}


@dataclass
class HistoryPoint:
    row: int
    x: float
    y: float
    label: str


def normalize_matrix(values: Any) -> list[list[Any]]:
    if values is None:
        return []
    if not isinstance(values, list):
        return [[values]]
    if values and not isinstance(values[0], list):
        return [values]
    return values


def is_number(value: Any) -> bool:
    if value is None or isinstance(value, bool):
        return False
    if isinstance(value, (int, float)):
        return not (isinstance(value, float) and math.isnan(value))
    return False


def as_float(value: Any) -> float | None:
    if is_number(value):
        return float(value)
    return None


def parse_file_labels(file_path: Path) -> dict[str, str]:
    stem = file_path.stem
    ticker_match = re.search(r"\s-\s([A-Za-z0-9]+)\s-\s", stem)
    period_match = re.search(
        r"\b(Early|Mid|Late)(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)(\d{4})\b",
        stem,
        flags=re.IGNORECASE,
    )
    if not ticker_match:
        raise ValueError("ticker pattern not found")
    if not period_match:
        raise ValueError("model period pattern not found")

    ticker = ticker_match.group(1).upper()
    period_word = period_match.group(1).title()
    month_abbr = period_match.group(2).title()
    year = int(period_match.group(3))

    month_num = MONTHS[month_abbr.lower()]
    day = PERIOD_DAY[period_word.lower()]
    model_period = f"{period_word}{month_abbr}_{year}"
    model_date = date(year, month_num, day).isoformat()
    model = f"{ticker}_{model_period}"

    return {
        "model": model,
        "ticker": ticker,
        "model_period": model_period,
        "model_date": model_date,
    }


def next_output_path(input_path: Path, out_path: Path) -> Path:
    stem = f"{input_path.name}_PARAM"
    candidate = out_path / f"{stem}.xlsx"
    counter = 1
    while candidate.exists():
        candidate = out_path / f"{stem}.{counter}.xlsx"
        counter += 1
    return candidate


def safe_close_workbook(wb: xw.Book) -> None:
    try:
        wb.close(save=False)
        return
    except TypeError:
        pass
    except Exception:
        pass

    try:
        wb.api.Close(SaveChanges=False)
        return
    except Exception:
        pass

    try:
        wb.api.Saved = True
    except Exception:
        pass

    try:
        wb.close()
    except Exception:
        pass


def find_max_anchor(sheet: xw.Sheet) -> tuple[int, int]:
    used = sheet.used_range
    matrix = normalize_matrix(used.value)
    if not matrix:
        raise ValueError("sheet is empty")

    for r_idx, row_vals in enumerate(matrix):
        row_values = row_vals if isinstance(row_vals, list) else [row_vals]
        for c_idx, value in enumerate(row_values):
            if isinstance(value, str) and value.strip().lower() == "max":
                return used.row + r_idx, used.column + c_idx
    raise ValueError("could not find 'max' anchor")


def get_row_label(row_slice: list[Any], first_col: int, x_col: int, row_number: int) -> str:
    for col in range(x_col - 1, max(first_col - 1, x_col - 5), -1):
        idx = col - first_col
        if idx < 0 or idx >= len(row_slice):
            continue
        value = row_slice[idx]
        if isinstance(value, datetime):
            return value.date().isoformat()
        if isinstance(value, date):
            return value.isoformat()
        if isinstance(value, str) and value.strip():
            return value.strip()
    return f"row_{row_number}"


def collect_history_points(
    sheet: xw.Sheet,
    anchor_row: int,
    x_col: int,
    y_col: int,
    scan_rows: int = 260,
) -> list[HistoryPoint]:
    if anchor_row <= 2:
        return []

    start_row = max(1, anchor_row - scan_rows)
    end_row = anchor_row - 1
    first_col = max(1, min(x_col, y_col) - 4)
    last_col = max(x_col, y_col)

    matrix = normalize_matrix(
        sheet.range((start_row, first_col), (end_row, last_col)).value
    )
    if not matrix:
        return []

    raw_points: list[HistoryPoint] = []
    for idx, row_slice in enumerate(matrix):
        row_vals = row_slice if isinstance(row_slice, list) else [row_slice]
        row_number = start_row + idx
        x_idx = x_col - first_col
        y_idx = y_col - first_col
        if x_idx < 0 or y_idx < 0 or x_idx >= len(row_vals) or y_idx >= len(row_vals):
            continue

        x_val = as_float(row_vals[x_idx])
        y_val = as_float(row_vals[y_idx])
        if x_val is None or y_val is None:
            continue

        label = get_row_label(row_vals, first_col, x_col, row_number)
        raw_points.append(HistoryPoint(row=row_number, x=x_val, y=y_val, label=label))

    if not raw_points:
        return []

    groups: list[list[HistoryPoint]] = [[raw_points[0]]]
    for point in raw_points[1:]:
        if point.row == groups[-1][-1].row + 1:
            groups[-1].append(point)
        else:
            groups.append([point])

    for group in reversed(groups):
        if len(group) >= 3:
            return group
    return groups[-1]


def approx_equal(a: float | None, b: float | None, tol: float = 1e-9) -> bool:
    if a is None and b is None:
        return True
    if a is None or b is None:
        return False
    return abs(a - b) <= tol * max(1.0, abs(a), abs(b))


def extract_empirical_rows(
    wb: xw.Book,
    sheet: xw.Sheet,
    meta: dict[str, str],
    source_file: str,
) -> list[dict[str, Any]]:
    anchor_row, anchor_col = find_max_anchor(sheet)
    quarterly_col = anchor_col - 7
    reported_col = anchor_col - 11
    if quarterly_col < 1 or reported_col < 1:
        raise ValueError("invalid anchor offsets for empirical model")

    history = collect_history_points(
        sheet=sheet,
        anchor_row=anchor_row,
        x_col=reported_col,
        y_col=quarterly_col,
    )
    if len(history) < 3:
        return []

    target = history[-1]
    previous = history[-2]
    max_n = min(N_QUARTERS, len(history) - 1)

    temp_row = min(max(anchor_row + 2, 2), 1048570)
    temp_col = 16370

    rows: list[dict[str, Any]] = []
    for n in range(1, max_n + 1):
        window = history[-(n + 1) : -1]
        start_row = window[0].row
        end_row = window[-1].row

        avg_cell = sheet.range((temp_row, temp_col))
        max_pen_cell = sheet.range((temp_row, temp_col + 1))
        min_pen_cell = sheet.range((temp_row, temp_col + 2))
        fcst_cell = sheet.range((temp_row, temp_col + 3))
        fcst_max_cell = sheet.range((temp_row, temp_col + 4))
        fcst_min_cell = sheet.range((temp_row, temp_col + 5))

        ratio_let = (
            f"LET(rep,R{start_row}C{reported_col}:R{end_row}C{reported_col},"
            f"tot,R{start_row}C{quarterly_col}:R{end_row}C{quarterly_col},"
            "ratio,IFERROR(rep/tot,\"\"),"
            "FILTER(ratio,ratio<>\"\"))"
        )
        avg_cell.formula2 = f"=AVERAGE({ratio_let})"
        max_pen_cell.formula2 = f"=MAX({ratio_let})"
        min_pen_cell.formula2 = f"=MIN({ratio_let})"
        fcst_cell.formula2 = (
            f"=IFERROR(R{target.row}C{reported_col}/R{temp_row}C{temp_col},\"\")"
        )
        fcst_max_cell.formula2 = (
            f"=IFERROR(R{target.row}C{reported_col}/R{temp_row}C{temp_col + 2},\"\")"
        )
        fcst_min_cell.formula2 = (
            f"=IFERROR(R{target.row}C{reported_col}/R{temp_row}C{temp_col + 1},\"\")"
        )

        wb.app.calculate()
        calc = normalize_matrix(
            sheet.range((temp_row, temp_col), (temp_row, temp_col + 5)).value
        )[0]

        avg_pen = as_float(calc[0] if len(calc) > 0 else None)
        forecast_value = as_float(calc[3] if len(calc) > 3 else None)
        forecast_max = as_float(calc[4] if len(calc) > 4 else None)
        forecast_min = as_float(calc[5] if len(calc) > 5 else None)

        quarterly_sales = target.y
        reported_sales = target.x
        actual_value = quarterly_sales
        growth_rate_pct = (
            ((target.y - previous.y) / previous.y) if previous.y not in (0, None) else None
        )
        sales_captured = (
            (reported_sales / quarterly_sales) if quarterly_sales not in (0, None) else None
        )
        range_width = (
            (forecast_max - forecast_min)
            if forecast_max is not None and forecast_min is not None
            else None
        )

        rows.append(
            {
                "model": meta["model"],
                "ticker": meta["ticker"],
                "model_period": meta["model_period"],
                "model_date": meta["model_date"],
                "method": "empirical",
                "parameter_name": "avg_penetration_pct",
                "parameter_value": avg_pen,
                "num_quarters_used": n,
                "last_quarter_used": window[-1].label,
                "forecast_value": forecast_value,
                "actual_value": actual_value,
                "forecast_max": forecast_max,
                "forecast_min": forecast_min,
                "range_width": range_width,
                "avg_penetration_pct": avg_pen,
                "quarterly_sales": quarterly_sales,
                "reported_sales": reported_sales,
                "growth_rate_pct": growth_rate_pct,
                "sales_captured_in_db_pct": sales_captured,
                "source_file": source_file,
            }
        )

    sheet.range((temp_row, temp_col), (temp_row, temp_col + 5)).clear_contents()
    return rows


def extract_regression_rows(
    wb: xw.Book,
    sheet: xw.Sheet,
    meta: dict[str, str],
    source_file: str,
) -> list[dict[str, Any]]:
    anchor_row, anchor_col = find_max_anchor(sheet)
    y_col = anchor_col - 7
    x_col = anchor_col - 11
    if x_col < 1 or y_col < 1:
        raise ValueError("invalid anchor offsets for regression model")

    history = collect_history_points(sheet=sheet, anchor_row=anchor_row, x_col=x_col, y_col=y_col)
    if len(history) < 3:
        return []

    target = history[-1]
    max_n = min(N_QUARTERS, len(history) - 1)
    temp_row = min(max(anchor_row + 3, 3), 1048570)
    temp_col = 16370

    rows: list[dict[str, Any]] = []
    for n in range(1, max_n + 1):
        window = history[-(n + 1) : -1]
        start_row = window[0].row
        end_row = window[-1].row

        intercept_cell = sheet.range((temp_row, temp_col))
        slope_cell = sheet.range((temp_row, temp_col + 1))
        fcst_cell = sheet.range((temp_row, temp_col + 2))
        err_cell = sheet.range((temp_row, temp_col + 3))
        fcst_max_cell = sheet.range((temp_row, temp_col + 4))
        fcst_min_cell = sheet.range((temp_row, temp_col + 5))

        y_range = f"R{start_row}C{y_col}:R{end_row}C{y_col}"
        x_range = f"R{start_row}C{x_col}:R{end_row}C{x_col}"

        intercept_cell.formula2 = f"=INTERCEPT({y_range},{x_range})"
        slope_cell.formula2 = f"=SLOPE({y_range},{x_range})"
        fcst_cell.formula2 = (
            f"=R{temp_row}C{temp_col}+R{temp_row}C{temp_col + 1}*R{target.row}C{x_col}"
        )
        err_cell.formula2 = (
            "=LET("
            f"y,{y_range},"
            f"x,{x_range},"
            f"b0,R{temp_row}C{temp_col},"
            f"b1,R{temp_row}C{temp_col + 1},"
            "pred,b0+b1*x,"
            "MAX(ABS(IFERROR((y-pred)/pred,0)))"
            ")"
        )
        fcst_max_cell.formula2 = f"=R{temp_row}C{temp_col + 2}*(1+R{temp_row}C{temp_col + 3})"
        fcst_min_cell.formula2 = f"=R{temp_row}C{temp_col + 2}*(1-R{temp_row}C{temp_col + 3})"

        wb.app.calculate()
        calc = normalize_matrix(
            sheet.range((temp_row, temp_col), (temp_row, temp_col + 5)).value
        )[0]

        intercept = as_float(calc[0] if len(calc) > 0 else None)
        slope = as_float(calc[1] if len(calc) > 1 else None)
        forecast_value = as_float(calc[2] if len(calc) > 2 else None)
        forecast_max = as_float(calc[4] if len(calc) > 4 else None)
        forecast_min = as_float(calc[5] if len(calc) > 5 else None)
        range_width = (
            (forecast_max - forecast_min)
            if forecast_max is not None and forecast_min is not None
            else None
        )

        if n == max_n and rows:
            prev = rows[-1]
            if (
                approx_equal(intercept, prev.get("intercept"))
                and approx_equal(slope, prev.get("slope"))
                and approx_equal(forecast_value, prev.get("forecast_value"))
                and approx_equal(forecast_max, prev.get("forecast_max"))
                and approx_equal(forecast_min, prev.get("forecast_min"))
            ):
                continue

        rows.append(
            {
                "model": meta["model"],
                "ticker": meta["ticker"],
                "model_period": meta["model_period"],
                "model_date": meta["model_date"],
                "method": "regression",
                "parameter_name": "num_quarters_used",
                "parameter_value": n,
                "num_quarters_used": n,
                "forecast_value": forecast_value,
                "actual_value": target.y,
                "forecast_max": forecast_max,
                "forecast_min": forecast_min,
                "range_width": range_width,
                "intercept": intercept,
                "slope": slope,
                "source_file": source_file,
            }
        )

    sheet.range((temp_row, temp_col), (temp_row, temp_col + 5)).clear_contents()
    return rows


def set_reasonable_column_widths(ws: Any) -> None:
    max_width = 48
    min_width = 10
    for col_idx in range(1, ws.max_column + 1):
        letter = get_column_letter(col_idx)
        header = ws.cell(row=1, column=col_idx).value
        header_len = len(str(header)) if header is not None else 0
        content_len = header_len
        for row_idx in range(2, min(ws.max_row, 250) + 1):
            value = ws.cell(row=row_idx, column=col_idx).value
            if value is None:
                continue
            content_len = max(content_len, len(str(value)))
        ws.column_dimensions[letter].width = max(min_width, min(max_width, content_len + 2))


def write_output_workbook(
    output_path: Path,
    empirical_rows: list[dict[str, Any]],
    regression_rows: list[dict[str, Any]],
) -> None:
    wb = Workbook()
    default_sheet = wb.active
    wb.remove(default_sheet)

    sheets = [
        ("empirical_candidates", EMPIRICAL_COLUMNS, empirical_rows),
        ("regression_candidates", REGRESSION_COLUMNS, regression_rows),
    ]
    for title, columns, rows in sheets:
        ws = wb.create_sheet(title=title)
        ws.append(columns)
        for row in rows:
            ws.append([row.get(col) for col in columns])

        for cell in ws[1]:
            cell.font = Font(bold=True)
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions
        set_reasonable_column_widths(ws)

    wb.save(output_path)


def main() -> int:
    in_path = Path(input_dir).expanduser().resolve()
    out_path = Path(output_dir).expanduser().resolve()

    if not in_path.exists() or not in_path.is_dir():
        print(f"Input folder not found: {in_path}")
        return 1

    out_path.mkdir(parents=True, exist_ok=True)
    output_file = next_output_path(in_path, out_path)

    empirical_rows: list[dict[str, Any]] = []
    regression_rows: list[dict[str, Any]] = []
    files_processed = 0

    app = xw.App(visible=False, add_book=False)
    app.display_alerts = False
    app.screen_updating = False

    try:
        for file_path in sorted(in_path.iterdir()):
            if not file_path.is_file():
                continue
            if file_path.suffix.lower() != ".xlsx":
                print(f"skipped: {file_path.name} (not .xlsx)")
                continue
            if file_path.name.startswith("~"):
                print(f"skipped: {file_path.name} (temporary file)")
                continue

            try:
                meta = parse_file_labels(file_path)
            except ValueError as exc:
                print(f"skipped: {file_path.name} ({exc})")
                continue

            wb: xw.Book | None = None
            file_had_output = False
            try:
                wb = app.books.open(str(file_path), update_links=False)

                try:
                    empirical_sheet = wb.sheets["Empirical Model"]
                    extracted_empirical = extract_empirical_rows(
                        wb=wb,
                        sheet=empirical_sheet,
                        meta=meta,
                        source_file=file_path.name,
                    )
                    if extracted_empirical:
                        empirical_rows.extend(extracted_empirical)
                        file_had_output = True
                except Exception as exc:
                    print(f"skipped: {file_path.name} (empirical failed: {exc})")

                try:
                    regression_sheet = wb.sheets["Regression Model"]
                    extracted_regression = extract_regression_rows(
                        wb=wb,
                        sheet=regression_sheet,
                        meta=meta,
                        source_file=file_path.name,
                    )
                    if extracted_regression:
                        regression_rows.extend(extracted_regression)
                        file_had_output = True
                except Exception as exc:
                    print(f"skipped: {file_path.name} (regression failed: {exc})")

                if file_had_output:
                    files_processed += 1
                    print(f"processed: {file_path.name}")
                else:
                    print(f"skipped: {file_path.name} (no candidate rows generated)")
            except Exception as exc:
                print(f"skipped: {file_path.name} (open failed: {exc})")
            finally:
                if wb is not None:
                    safe_close_workbook(wb)
    finally:
        app.quit()

    write_output_workbook(
        output_path=output_file,
        empirical_rows=empirical_rows,
        regression_rows=regression_rows,
    )

    print(f"output_path: {output_file}")
    print(f"files_processed: {files_processed}")
    print(f"empirical_rows: {len(empirical_rows)}")
    print(f"regression_rows: {len(regression_rows)}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
