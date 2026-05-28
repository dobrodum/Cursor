#!/usr/bin/env python3
from __future__ import annotations

import re
from dataclasses import dataclass
from datetime import date
from pathlib import Path
from typing import Any

import xlwings as xw
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter


# Configure these paths before running.
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

EMPIRICAL_OFFSETS = {
    "forecast_max": 0,
    "forecast_min": -1,
    "forecast_value": -2,
    "actual_value": -3,
    "quarterly_sales": -4,
    "growth_rate_pct": -5,
    "sales_captured_in_db_pct": -6,
    "last_quarter_used": -7,
    "avg_penetration_pct": -8,
}

REGRESSION_OFFSETS = {
    "forecast_max": 0,
    "forecast_min": -1,
    "forecast_value": -2,  # TOT FCST w/o SA
    "actual_value": -3,
}

DAY_MAP = {"Early": 5, "Mid": 15, "Late": 25}
MONTH_MAP = {
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
PERIOD_RE = re.compile(r"(Early|Mid|Late)\s*([A-Za-z]{3,12})\s*(\d{4})", re.IGNORECASE)


@dataclass
class FileLabel:
    model: str
    ticker: str
    model_period: str
    model_date: str


def parse_file_label(file_name: str) -> FileLabel:
    stem = Path(file_name).stem
    parts = [part.strip() for part in stem.split("-")]

    ticker = "UNKNOWN"
    if len(parts) >= 2:
        ticker_candidate = re.sub(r"[^A-Za-z0-9]", "", parts[1]).upper()
        if ticker_candidate:
            ticker = ticker_candidate

    period_match = PERIOD_RE.search(stem)
    if not period_match:
        raise ValueError(
            f"could not parse model period from filename '{file_name}' "
            "(expected Early/Mid/Late + month + year)"
        )

    phase = period_match.group(1).title()
    month_token = period_match.group(2).title()[:3]
    year = int(period_match.group(3))

    if month_token not in MONTH_MAP:
        raise ValueError(f"unrecognized month token '{month_token}' in filename '{file_name}'")

    month_number = MONTH_MAP[month_token]
    day = DAY_MAP[phase]
    model_period = f"{phase}{month_token}_{year}"
    model_date = date(year, month_number, day).isoformat()
    model = f"{ticker}_{model_period}"

    return FileLabel(model=model, ticker=ticker, model_period=model_period, model_date=model_date)


def build_output_path(input_path: Path, output_path: Path) -> Path:
    output_path.mkdir(parents=True, exist_ok=True)
    base_name = f"{input_path.name}_PARAM"
    candidate = output_path / f"{base_name}.xlsx"
    idx = 1
    while candidate.exists():
        candidate = output_path / f"{base_name}.{idx}.xlsx"
        idx += 1
    return candidate


def collect_source_files(source_dir: Path) -> list[Path]:
    if not source_dir.exists():
        raise FileNotFoundError(f"input directory does not exist: {source_dir}")
    if not source_dir.is_dir():
        raise NotADirectoryError(f"input_dir must be a directory: {source_dir}")

    source_files: list[Path] = []
    for path in sorted(source_dir.iterdir()):
        if not path.is_file():
            print(f"Skipped {path.name}: not a file")
            continue
        if path.name.startswith("~"):
            print(f"Skipped {path.name}: temporary Excel file")
            continue
        if path.suffix.lower() != ".xlsx":
            print(f"Skipped {path.name}: not an .xlsx file")
            continue
        source_files.append(path)

    return source_files


def ensure_2d(values: Any) -> list[list[Any]]:
    if values is None:
        return []
    if not isinstance(values, list):
        return [[values]]
    if values and not isinstance(values[0], list):
        return [values]
    return values


def find_anchor(sheet: xw.Sheet, label: str = "max") -> tuple[int, int]:
    used = sheet.used_range
    values = ensure_2d(used.value)
    wanted = label.strip().lower()

    for row_idx, row_values in enumerate(values):
        for col_idx, cell_value in enumerate(row_values):
            if isinstance(cell_value, str) and cell_value.strip().lower() == wanted:
                return used.row + row_idx, used.column + col_idx

    raise ValueError(f'anchor "{label}" not found on sheet "{sheet.name}"')


def set_formula2(cell: xw.Range, formula_r1c1: str) -> None:
    try:
        cell.formula2 = formula_r1c1
    except Exception:
        # Fallback for Excel versions without formula2 support.
        cell.formula = formula_r1c1


def safe_close_workbook(wb: xw.Book) -> None:
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
        print(f"Warning: failed to close workbook safely ({wb.name}): {exc}")


def as_float(value: Any) -> float | None:
    if value in (None, ""):
        return None
    try:
        return float(value)
    except (TypeError, ValueError):
        return None


def subtract_if_numeric(left: Any, right: Any) -> float | str:
    left_float = as_float(left)
    right_float = as_float(right)
    if left_float is None or right_float is None:
        return ""
    return left_float - right_float


def extract_empirical_rows(
    wb: xw.Book,
    meta: FileLabel,
    source_file: str,
) -> list[dict[str, Any]]:
    sheet = wb.sheets["Empirical Model"]
    anchor_row, anchor_col = find_anchor(sheet, "max")
    rows: list[dict[str, Any]] = []

    # Temporary formula writes trigger calc results we need for extraction.
    for n_quarters in range(1, N_QUARTERS + 1):
        row_idx = anchor_row + n_quarters
        avg_cell = sheet.cells(row_idx, anchor_col + EMPIRICAL_OFFSETS["avg_penetration_pct"])
        set_formula2(avg_cell, f'=IFERROR(AVERAGE(RC[-{n_quarters}]:RC[-1]),"")')

    wb.app.calculate()

    for n_quarters in range(1, N_QUARTERS + 1):
        row_idx = anchor_row + n_quarters
        forecast_max = sheet.cells(row_idx, anchor_col + EMPIRICAL_OFFSETS["forecast_max"]).value
        forecast_min = sheet.cells(row_idx, anchor_col + EMPIRICAL_OFFSETS["forecast_min"]).value
        forecast_value = sheet.cells(row_idx, anchor_col + EMPIRICAL_OFFSETS["forecast_value"]).value
        actual_value = sheet.cells(row_idx, anchor_col + EMPIRICAL_OFFSETS["actual_value"]).value
        quarterly_sales = sheet.cells(row_idx, anchor_col + EMPIRICAL_OFFSETS["quarterly_sales"]).value
        growth_rate_pct = sheet.cells(row_idx, anchor_col + EMPIRICAL_OFFSETS["growth_rate_pct"]).value
        sales_captured = sheet.cells(
            row_idx,
            anchor_col + EMPIRICAL_OFFSETS["sales_captured_in_db_pct"],
        ).value
        last_quarter_used = sheet.cells(
            row_idx,
            anchor_col + EMPIRICAL_OFFSETS["last_quarter_used"],
        ).value
        avg_penetration = sheet.cells(
            row_idx,
            anchor_col + EMPIRICAL_OFFSETS["avg_penetration_pct"],
        ).value

        if all(
            val in (None, "")
            for val in [forecast_max, forecast_min, forecast_value, actual_value, avg_penetration]
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
                "forecast_value": forecast_value,  # Estimated total sold.
                "actual_value": actual_value,  # Reported sales.
                "forecast_max": forecast_max,
                "forecast_min": forecast_min,
                "range_width": subtract_if_numeric(forecast_max, forecast_min),
                "avg_penetration_pct": avg_penetration,
                "quarterly_sales": quarterly_sales,
                "reported_sales": actual_value,
                "growth_rate_pct": growth_rate_pct,
                "sales_captured_in_db_pct": sales_captured,
                "source_file": source_file,
            }
        )

    return rows


def extract_regression_rows(
    wb: xw.Book,
    meta: FileLabel,
    source_file: str,
) -> list[dict[str, Any]]:
    sheet = wb.sheets["Regression Model"]
    anchor_row, anchor_col = find_anchor(sheet, "max")
    y_col = anchor_col - 7
    x_col = anchor_col - 11

    used_last_col = sheet.used_range.last_cell.column
    helper_intercept_col = max(anchor_col + 2, used_last_col + 2)
    helper_slope_col = helper_intercept_col + 1
    data_start_row = anchor_row + 1

    rows: list[dict[str, Any]] = []

    for n_quarters in range(1, N_QUARTERS + 1):
        helper_row = anchor_row + n_quarters
        data_end_row = data_start_row + n_quarters - 1
        intercept_cell = sheet.cells(helper_row, helper_intercept_col)
        slope_cell = sheet.cells(helper_row, helper_slope_col)

        set_formula2(
            intercept_cell,
            (
                f'=IFERROR(INTERCEPT(R{data_start_row}C{y_col}:R{data_end_row}C{y_col},'
                f'R{data_start_row}C{x_col}:R{data_end_row}C{x_col}),"")'
            ),
        )
        set_formula2(
            slope_cell,
            (
                f'=IFERROR(SLOPE(R{data_start_row}C{y_col}:R{data_end_row}C{y_col},'
                f'R{data_start_row}C{x_col}:R{data_end_row}C{x_col}),"")'
            ),
        )

    wb.app.calculate()

    previous_signature: tuple[Any, ...] | None = None
    for n_quarters in range(1, N_QUARTERS + 1):
        row_idx = anchor_row + n_quarters
        forecast_max = sheet.cells(row_idx, anchor_col + REGRESSION_OFFSETS["forecast_max"]).value
        forecast_min = sheet.cells(row_idx, anchor_col + REGRESSION_OFFSETS["forecast_min"]).value
        forecast_total = sheet.cells(row_idx, anchor_col + REGRESSION_OFFSETS["forecast_value"]).value
        actual_candidate = sheet.cells(row_idx, anchor_col + REGRESSION_OFFSETS["actual_value"]).value
        intercept = sheet.cells(row_idx, helper_intercept_col).value
        slope = sheet.cells(row_idx, helper_slope_col).value

        if all(val in (None, "") for val in [forecast_max, forecast_min, forecast_total, intercept, slope]):
            continue

        signature = (forecast_total, forecast_max, forecast_min, intercept, slope)
        if previous_signature is not None and signature == previous_signature:
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
                "forecast_value": forecast_total,
                "actual_value": actual_candidate if actual_candidate not in (None, "") else "",
                "forecast_max": forecast_max,
                "forecast_min": forecast_min,
                "range_width": subtract_if_numeric(forecast_max, forecast_min),
                "intercept": intercept,
                "slope": slope,
                "source_file": source_file,
            }
        )

    return rows


def write_output_sheet(
    wb: Workbook,
    sheet_name: str,
    headers: list[str],
    rows: list[dict[str, Any]],
) -> None:
    ws = wb.create_sheet(title=sheet_name)
    ws.append(headers)

    for cell in ws[1]:
        cell.font = Font(bold=True)

    for row_data in rows:
        ws.append([row_data.get(header, "") for header in headers])

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(headers))}{max(ws.max_row, 1)}"

    for idx, header in enumerate(headers, start=1):
        max_width = len(header)
        for col_cells in ws.iter_cols(min_col=idx, max_col=idx, min_row=2, max_row=ws.max_row):
            for value_cell in col_cells:
                if value_cell.value in (None, ""):
                    continue
                max_width = max(max_width, len(str(value_cell.value)))
        ws.column_dimensions[get_column_letter(idx)].width = min(max_width + 2, 40)


def write_output_workbook(
    out_path: Path,
    empirical_rows: list[dict[str, Any]],
    regression_rows: list[dict[str, Any]],
) -> None:
    output_wb = Workbook()
    output_wb.remove(output_wb.active)

    write_output_sheet(output_wb, "empirical_candidates", EMPIRICAL_HEADERS, empirical_rows)
    write_output_sheet(output_wb, "regression_candidates", REGRESSION_HEADERS, regression_rows)

    output_wb.save(out_path)


def process_all_workbooks(source_files: list[Path]) -> tuple[int, list[dict[str, Any]], list[dict[str, Any]]]:
    processed_files = 0
    empirical_rows: list[dict[str, Any]] = []
    regression_rows: list[dict[str, Any]] = []

    app = xw.App(visible=False, add_book=False)
    app.display_alerts = False
    app.screen_updating = False
    try:
        app.calculation = "manual"
    except Exception:
        pass

    try:
        for file_path in source_files:
            print(f"Processing {file_path.name}")
            wb = None
            try:
                wb = app.books.open(str(file_path), update_links=False)
                meta = parse_file_label(file_path.name)

                try:
                    empirical_rows.extend(extract_empirical_rows(wb, meta, file_path.name))
                except Exception as exc:
                    print(f"Skipped {file_path.name}: empirical extraction failed ({exc})")

                try:
                    regression_rows.extend(extract_regression_rows(wb, meta, file_path.name))
                except Exception as exc:
                    print(f"Skipped {file_path.name}: regression extraction failed ({exc})")

                processed_files += 1
            except Exception as exc:
                print(f"Skipped {file_path.name}: open/parse failed ({exc})")
            finally:
                if wb is not None:
                    safe_close_workbook(wb)
    finally:
        try:
            app.quit()
        except Exception:
            pass

    return processed_files, empirical_rows, regression_rows


def main() -> None:
    input_path = Path(input_dir)
    output_path = Path(output_dir)

    source_files = collect_source_files(input_path)
    output_file = build_output_path(input_path, output_path)

    processed_files, empirical_rows, regression_rows = process_all_workbooks(source_files)
    write_output_workbook(output_file, empirical_rows, regression_rows)

    print(f"Output path: {output_file}")
    print(f"Number of files processed: {processed_files}")
    print(f"Number of empirical rows: {len(empirical_rows)}")
    print(f"Number of regression rows: {len(regression_rows)}")


if __name__ == "__main__":
    main()
