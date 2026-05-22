#!/usr/bin/env python3
from __future__ import annotations

import argparse
import re
import sys
from dataclasses import dataclass
from datetime import date
from pathlib import Path
from typing import Any

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font

DEFAULT_N_QUARTERS = 10
DEFAULT_ROWS_PER_QUARTER = 3

EMPIRICAL_HEADERS = [
    "candidate_id",
    "model",
    "ticker",
    "model_date",
    "method",
    "num_quarters_used",
    "quarter_window_start_date",
    "quarter_window_end_date",
    "quarter_window_label",
    "avg_penetration_pct",
    "forecast_value",
    "forecast_max",
    "forecast_min",
    "range_width",
    "last_quarter_used",
    "quarterly_sales",
    "reported_sales",
    "growth_rate_pct",
    "sales_captured_in_db_pct",
]

REGRESSION_HEADERS = [
    "candidate_id",
    "model",
    "ticker",
    "model_date",
    "method",
    "num_quarters_used",
    "quarter_window_start_date",
    "quarter_window_end_date",
    "quarter_window",
    "intercept",
    "slope",
    "forecast_value",
    "forecast_max",
    "forecast_min",
    "range_width",
]


@dataclass
class SheetSnapshot:
    start_row: int
    start_col: int
    values: list[list[Any]]

    @property
    def end_row(self) -> int:
        return self.start_row + len(self.values) - 1

    @property
    def end_col(self) -> int:
        if not self.values:
            return self.start_col
        return self.start_col + len(self.values[0]) - 1

    def get(self, row: int, col: int) -> Any:
        row_idx = row - self.start_row
        col_idx = col - self.start_col
        if row_idx < 0 or col_idx < 0:
            return None
        if row_idx >= len(self.values):
            return None
        current_row = self.values[row_idx]
        if col_idx >= len(current_row):
            return None
        return current_row[col_idx]


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Batch-extract parameters from MedMiner Excel workbooks."
    )
    parser.add_argument(
        "--input-dir",
        default="/Users/mariadobrodum/Desktop/MedMine/MedMiner_Models",
    )
    parser.add_argument(
        "--output-dir",
        default="/Users/mariadobrodum/Desktop/MedMine/outputs",
    )
    parser.add_argument("--n-quarters", type=int, default=DEFAULT_N_QUARTERS)
    parser.add_argument("--rows-per-quarter", type=int, default=DEFAULT_ROWS_PER_QUARTER)
    parser.add_argument("--visible", action="store_true")
    return parser.parse_args()


def normalize(value: Any) -> str:
    if value is None:
        return ""
    return str(value).strip().lower()


def safe_float(value: Any) -> float | None:
    if value in (None, ""):
        return None
    try:
        return float(value)
    except (TypeError, ValueError):
        return None


def rounded_signature(*values: Any) -> tuple[Any, ...]:
    out: list[Any] = []
    for value in values:
        if isinstance(value, (int, float)):
            out.append(round(float(value), 10))
        else:
            out.append(value)
    return tuple(out)


def normalize_2d(values: Any) -> list[list[Any]]:
    if values is None:
        return [[]]
    if not isinstance(values, list):
        return [[values]]
    if values and not isinstance(values[0], list):
        return [values]

    matrix = values
    width = max((len(row) for row in matrix), default=0)
    if width == 0:
        return [[]]

    out: list[list[Any]] = []
    for row in matrix:
        out.append(list(row) + [None] * (width - len(row)))
    return out


def parse_model_period_to_date(model_period: str) -> date | None:
    month_map = {
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

    day_map = {"Early": 5, "Mid": 15, "Late": 25}
    match = re.match(r"^(Early|Mid|Late)([A-Za-z]{3})_(\d{4})$", model_period)
    if not match:
        return None

    part, month_abbr, year = match.groups()
    month_key = month_abbr.title()
    if month_key not in month_map:
        return None

    return date(int(year), month_map[month_key], day_map[part])


def parse_file_labels(file_path: Path) -> tuple[str, str, date | None]:
    parts = file_path.stem.split(" - ")
    if len(parts) < 3:
        raise ValueError("expected 'MedMiner_Model - TICKER - DATE_Send'")

    ticker = parts[1].strip()
    raw_period = parts[-1].split("_")[0].replace("-", "_")
    model_period = f"{raw_period[:-4]}_{raw_period[-4:]}"
    model_date = parse_model_period_to_date(model_period)
    model = f"{ticker}_{model_period}"
    return model, ticker, model_date


def format_quarter_label(value: Any) -> str | None:
    if value is None:
        return None

    text = str(value).strip()
    match = re.match(r"^Q([1-4])[\s_-]?(\d{2}|\d{4})$", text, re.IGNORECASE)
    if not match:
        return text.lower().replace(" ", "_")

    quarter, year = match.groups()
    if len(year) == 2:
        year = f"20{year}"
    return f"q{quarter}_{year}"


def quarter_label_to_date(quarter_label: str | None) -> date | None:
    if not quarter_label:
        return None
    match = re.match(r"^q([1-4])_(\d{4})$", quarter_label)
    if not match:
        return None

    quarter, year = match.groups()
    month = (int(quarter) - 1) * 3 + 1
    return date(int(year), month, 1)


def make_quarter_window_label(start_label: str | None, end_label: str | None) -> str | None:
    if start_label is None and end_label is None:
        return None
    if start_label == end_label:
        return start_label
    return f"{start_label}_to_{end_label}"


def make_candidate_id(model: str, method: str, num_quarters_used: int) -> str:
    method_code = "EMP" if method == "empirical" else "REG"
    return f"{model}_{method_code}_{int(num_quarters_used):03d}"


def set_formula2(cell: Any, formula: str) -> None:
    try:
        cell.formula2 = formula
    except Exception:
        cell.formula = formula


def take_snapshot(sheet: Any) -> SheetSnapshot:
    used = sheet.used_range
    matrix = normalize_2d(used.value)
    return SheetSnapshot(start_row=used.row, start_col=used.column, values=matrix)


def find_anchor(snapshot: SheetSnapshot, target: str = "max") -> tuple[int, int] | None:
    needle = target.strip().lower()
    for row_offset, row in enumerate(snapshot.values):
        for col_offset, value in enumerate(row):
            if normalize(value) == needle:
                return snapshot.start_row + row_offset, snapshot.start_col + col_offset
    return None


def find_last_non_empty_row(snapshot: SheetSnapshot, col: int) -> int | None:
    for row in range(snapshot.end_row, snapshot.start_row - 1, -1):
        if snapshot.get(row, col) not in (None, ""):
            return row
    return None


def get_sheet_by_name(workbook: Any, name: str) -> Any | None:
    key = name.strip().lower()
    for sheet in workbook.sheets:
        if sheet.name.strip().lower() == key:
            return sheet
    return None


def process_empirical(
    wb: Any,
    file_path: Path,
    n_quarters: int,
    rows_per_quarter: int,
) -> tuple[list[list[Any]], list[list[Any]]]:
    calc_pulls: list[list[Any]] = []
    raw_pulls: list[list[Any]] = []

    sheet = get_sheet_by_name(wb, "Empirical Model")
    if sheet is None:
        print(f"Empirical Model sheet not found — skipping {file_path.name}")
        return calc_pulls, raw_pulls

    snapshot = take_snapshot(sheet)
    anchor = find_anchor(snapshot, "max")
    if anchor is None:
        print(f"Empirical MAX anchor not found — skipping {file_path.name}")
        return calc_pulls, raw_pulls

    anchor_row, anchor_col = anchor
    max_row = anchor_row
    min_row = anchor_row + 1
    est_row = anchor_row - 1
    pen_row = anchor_row + 5

    est_col = anchor_col + 1
    sales_cap_col = anchor_col + 4
    param_col = anchor_col + 5
    b_col = anchor_col - 3
    e_col = anchor_col
    g_col = anchor_col + 2
    h_col = anchor_col + 3

    last_b_row = find_last_non_empty_row(snapshot, b_col)
    if last_b_row is None:
        print(f"Empirical source rows not found — skipping {file_path.name}")
        return calc_pulls, raw_pulls

    last_row = last_b_row - rows_per_quarter
    if last_row < snapshot.start_row:
        print(f"Empirical source rows invalid — skipping {file_path.name}")
        return calc_pulls, raw_pulls

    avg_pen_cell = sheet.cells(pen_row, param_col)

    for i in range(1, n_quarters + 1):
        first_row = last_row - (i * rows_per_quarter) + 1
        if first_row < snapshot.start_row:
            break

        set_formula2(
            avg_pen_cell,
            f"=AVERAGE(R{first_row}C{sales_cap_col}:R{last_row}C{sales_cap_col})",
        )

        wb.app.calculate()

        avg_pen = safe_float(avg_pen_cell.value)
        if avg_pen is not None:
            avg_pen = round(avg_pen * 100, 5)

        forecast_range = normalize_2d(
            sheet.range((est_row, est_col), (min_row, est_col)).value
        )
        forecast_value = safe_float(forecast_range[0][0]) if forecast_range else None
        forecast_max = safe_float(forecast_range[1][0]) if len(forecast_range) > 1 else None
        forecast_min = safe_float(forecast_range[2][0]) if len(forecast_range) > 2 else None

        range_width = None
        if forecast_max is not None and forecast_min is not None:
            range_width = forecast_max - forecast_min

        calc_pulls.append(
            [i, avg_pen, forecast_value, forecast_max, forecast_min, range_width]
        )

    raw_rows: list[int] = []
    for k in range(1, n_quarters + 1):
        row_number = last_row - ((k - 1) * rows_per_quarter)
        if row_number < snapshot.start_row:
            break
        raw_rows.append(row_number)

    if raw_rows:
        row_start = raw_rows[-1]
        row_end = raw_rows[0]
        block = normalize_2d(sheet.range((row_start, b_col), (row_end, sales_cap_col)).value)

        e_idx = e_col - b_col
        g_idx = g_col - b_col
        h_idx = h_col - b_col
        cap_idx = sales_cap_col - b_col

        for row_number in raw_rows:
            row_values = block[row_number - row_start]
            last_quarter_used = format_quarter_label(row_values[0] if row_values else None)
            quarterly_sales = row_values[e_idx] if e_idx < len(row_values) else None
            reported_sales = row_values[g_idx] if g_idx < len(row_values) else None
            growth_rate_pct = row_values[h_idx] if h_idx < len(row_values) else None
            captured_pct = row_values[cap_idx] if cap_idx < len(row_values) else None

            if isinstance(growth_rate_pct, (int, float)):
                growth_rate_pct = growth_rate_pct * 100
            if isinstance(captured_pct, (int, float)):
                captured_pct = captured_pct * 100

            raw_pulls.append(
                [
                    last_quarter_used,
                    quarterly_sales,
                    reported_sales,
                    growth_rate_pct,
                    captured_pct,
                ]
            )

    return calc_pulls, raw_pulls


def process_regression(wb: Any, file_path: Path) -> list[list[Any]]:
    calc_pulls: list[list[Any]] = []

    sheet = get_sheet_by_name(wb, "Regression Model")
    if sheet is None:
        print(f"Regression Model sheet not found — skipping {file_path.name}")
        return calc_pulls

    snapshot = take_snapshot(sheet)
    anchor = find_anchor(snapshot, "max")
    if anchor is None:
        print(f"Regression MAX anchor not found — skipping {file_path.name}")
        return calc_pulls

    anchor_row, anchor_col = anchor
    last_data_row = anchor_row - 1
    fixed_end = last_data_row - 1
    start_row = fixed_end - 3

    slope_col = anchor_col - 3
    intercept_col = anchor_col - 2
    forecast_col = anchor_col + 1
    y_col = anchor_col - 7
    x_col = anchor_col - 11
    quarter_col = anchor_col - 14

    quarter_end_label = format_quarter_label(snapshot.get(fixed_end, quarter_col))
    quarter_end_date = quarter_label_to_date(quarter_end_label)

    intercept_cell = sheet.cells(last_data_row, intercept_col)
    slope_cell = sheet.cells(last_data_row, slope_col)
    previous_signature: tuple[Any, ...] | None = None

    i = 0
    while True:
        current_start = start_row - i
        if current_start < snapshot.start_row:
            break

        y_value = safe_float(snapshot.get(current_start, y_col))
        x_value = safe_float(snapshot.get(current_start, x_col))
        if y_value is None or x_value is None:
            break

        set_formula2(
            intercept_cell,
            (
                "=INTERCEPT("
                f"R{current_start}C{y_col}:R{fixed_end}C{y_col},"
                f"R{current_start}C{x_col}:R{fixed_end}C{x_col})"
            ),
        )
        set_formula2(
            slope_cell,
            (
                "=SLOPE("
                f"R{current_start}C{y_col}:R{fixed_end}C{y_col},"
                f"R{current_start}C{x_col}:R{fixed_end}C{x_col})"
            ),
        )

        wb.app.calculate()

        output_block = normalize_2d(
            sheet.range((last_data_row, slope_col), (last_data_row, forecast_col)).value
        )
        output_row = output_block[0] if output_block else []

        slope = safe_float(output_row[0]) if output_row else None
        intercept_idx = intercept_col - slope_col
        forecast_idx = forecast_col - slope_col
        intercept = safe_float(output_row[intercept_idx]) if intercept_idx < len(output_row) else None
        forecast_value = safe_float(output_row[forecast_idx]) if forecast_idx < len(output_row) else None

        max_min = normalize_2d(sheet.range((anchor_row, forecast_col), (anchor_row + 1, forecast_col)).value)
        forecast_max = safe_float(max_min[0][0]) if max_min else None
        forecast_min = safe_float(max_min[1][0]) if len(max_min) > 1 else None

        range_width = None
        if forecast_max is not None and forecast_min is not None:
            range_width = forecast_max - forecast_min

        num_quarters_used = i + 4
        quarter_start_label = format_quarter_label(snapshot.get(current_start, quarter_col))
        quarter_start_date = quarter_label_to_date(quarter_start_label)
        quarter_window = make_quarter_window_label(quarter_start_label, quarter_end_label)

        signature = rounded_signature(
            intercept,
            slope,
            forecast_value,
            forecast_max,
            forecast_min,
            range_width,
        )
        if previous_signature == signature:
            break
        previous_signature = signature

        calc_pulls.append(
            [
                num_quarters_used,
                quarter_start_date,
                quarter_end_date,
                quarter_window,
                intercept,
                slope,
                forecast_value,
                forecast_max,
                forecast_min,
                range_width,
            ]
        )
        i += 1

    return calc_pulls


def format_sheet(ws: Any) -> None:
    for cell in ws[1]:
        cell.font = Font(size=12, bold=True)
        cell.alignment = Alignment(horizontal="center")

    for col_idx in range(1, ws.max_column + 1):
        col_letter = ws.cell(1, col_idx).column_letter
        ws.column_dimensions[col_letter].width = 18

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions


def write_output_workbook(
    output_file: Path,
    empirical_rows: list[list[Any]],
    regression_rows: list[list[Any]],
) -> None:
    workbook = Workbook()
    empirical_sheet = workbook.active
    empirical_sheet.title = "Empirical Param"
    regression_sheet = workbook.create_sheet("Regression Param")

    empirical_sheet.append(EMPIRICAL_HEADERS)
    for row in empirical_rows:
        empirical_sheet.append(row)

    regression_sheet.append(REGRESSION_HEADERS)
    for row in regression_rows:
        regression_sheet.append(row)

    format_sheet(empirical_sheet)
    format_sheet(regression_sheet)
    workbook.save(output_file)


def build_output_path(in_path: Path, out_dir: Path) -> Path:
    output_file = out_dir / f"{in_path.name}_PARAM.xlsx"
    if output_file.exists():
        n = 1
        while True:
            candidate = out_dir / f"{in_path.name}_PARAM.{n}.xlsx"
            if not candidate.exists():
                output_file = candidate
                break
            n += 1
    return output_file


def iter_excel_files(in_path: Path, output_file: Path) -> list[Path]:
    files: list[Path] = []
    for path in sorted(in_path.iterdir()):
        if not path.is_file():
            continue
        if path.name.startswith("~"):
            print(f"Skipped {path.name}: temp file")
            continue
        if path.suffix.lower() != ".xlsx":
            continue
        if output_file.exists() and path.resolve() == output_file.resolve():
            continue
        files.append(path)
    return files


def safe_close_workbook(wb: Any, file_name: str) -> None:
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
    except Exception as exc:
        print(f"Warning: could not close {file_name}: {exc}")


def build_empirical_rows(
    model: str,
    ticker: str,
    model_date: date | None,
    emp_calc: list[list[Any]],
    emp_raw: list[list[Any]],
) -> list[list[Any]]:
    if not emp_calc:
        return []

    quarter_end_label = emp_raw[0][0] if emp_raw else None
    quarter_end_date = quarter_label_to_date(quarter_end_label)

    rows: list[list[Any]] = []
    for idx, calc in enumerate(emp_calc):
        num_quarters_used, avg_pen, forecast_value, forecast_max, forecast_min, range_width = calc
        raw = emp_raw[idx] if idx < len(emp_raw) else [None, None, None, None, None]
        last_quarter_used, quarterly_sales, reported_sales, growth_rate_pct, captured_pct = raw

        quarter_start_date = quarter_label_to_date(last_quarter_used)
        quarter_window_label = make_quarter_window_label(last_quarter_used, quarter_end_label)

        rows.append(
            [
                make_candidate_id(model, "empirical", int(num_quarters_used)),
                model,
                ticker,
                model_date,
                "empirical",
                num_quarters_used,
                quarter_start_date,
                quarter_end_date,
                quarter_window_label,
                avg_pen,
                forecast_value,
                forecast_max,
                forecast_min,
                range_width,
                last_quarter_used,
                quarterly_sales,
                reported_sales,
                growth_rate_pct,
                captured_pct,
            ]
        )
    return rows


def build_regression_rows(
    model: str,
    ticker: str,
    model_date: date | None,
    reg_calc: list[list[Any]],
) -> list[list[Any]]:
    rows: list[list[Any]] = []
    for calc in reg_calc:
        (
            num_quarters_used,
            quarter_start_date,
            quarter_end_date,
            quarter_window,
            intercept,
            slope,
            forecast_value,
            forecast_max,
            forecast_min,
            range_width,
        ) = calc

        rows.append(
            [
                make_candidate_id(model, "regression", int(num_quarters_used)),
                model,
                ticker,
                model_date,
                "regression",
                num_quarters_used,
                quarter_start_date,
                quarter_end_date,
                quarter_window,
                intercept,
                slope,
                forecast_value,
                forecast_max,
                forecast_min,
                range_width,
            ]
        )
    return rows


def main() -> None:
    args = parse_args()

    if sys.platform.startswith("linux"):
        raise RuntimeError("xlwings needs desktop Excel.")

    try:
        import xlwings as xw
    except ImportError as exc:
        raise RuntimeError("xlwings is not installed.") from exc

    in_path = Path(args.input_dir).expanduser().resolve()
    out_dir = Path(args.output_dir).expanduser().resolve()
    out_dir.mkdir(parents=True, exist_ok=True)
    output_file = build_output_path(in_path, out_dir)

    excel_files = iter_excel_files(in_path, output_file)
    if not excel_files:
        raise FileNotFoundError(f"No .xlsx files found in {in_path}")

    empirical_rows: list[list[Any]] = []
    regression_rows: list[list[Any]] = []
    processed_files = 0

    with xw.App(visible=args.visible, add_book=False) as app:
        app.display_alerts = False
        app.screen_updating = False
        try:
            app.calculation = "manual"
        except Exception:
            pass
        try:
            app.api.EnableEvents = False
        except Exception:
            pass

        for file_path in excel_files:
            workbook = None
            try:
                model, ticker, model_date = parse_file_labels(file_path)
                workbook = app.books.open(
                    str(file_path),
                    read_only=True,
                    update_links=False,
                )

                emp_calc, emp_raw = process_empirical(
                    workbook,
                    file_path,
                    args.n_quarters,
                    args.rows_per_quarter,
                )
                reg_calc = process_regression(workbook, file_path)

                empirical_rows.extend(
                    build_empirical_rows(model, ticker, model_date, emp_calc, emp_raw)
                )
                regression_rows.extend(
                    build_regression_rows(model, ticker, model_date, reg_calc)
                )

                processed_files += 1
                print(f"Processed: {file_path.name}")
            except Exception as exc:
                print(f"Skipped {file_path.name}: {exc}")
            finally:
                if workbook is not None:
                    safe_close_workbook(workbook, file_path.name)

    if not empirical_rows and not regression_rows:
        raise ValueError("No files were processed successfully.")

    write_output_workbook(output_file, empirical_rows, regression_rows)

    print(f"\nSaved: {output_file}")
    print(f"Files processed: {processed_files}")
    print(f"Empirical rows: {len(empirical_rows)}")
    print(f"Regression rows: {len(regression_rows)}")


if __name__ == "__main__":
    main()
