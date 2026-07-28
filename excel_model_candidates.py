from __future__ import annotations

from dataclasses import dataclass
from datetime import date
from pathlib import Path
import re
from typing import Any

import xlwings as xw
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter


# Update these paths before running.
input_dir = Path("./input")
output_dir = Path("./output")


EMPIRICAL_SHEET_NAME = "Empirical Model"
REGRESSION_SHEET_NAME = "Regression Model"
OUTPUT_EMPIRICAL_SHEET = "empirical_candidates"
OUTPUT_REGRESSION_SHEET = "regression_candidates"
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

PHASE_DAY = {"early": 5, "mid": 15, "late": 25}


EMPIRICAL_FALLBACK_OFFSETS = {
    "num_quarters_used": -8,
    "last_quarter_used": -7,
    "forecast_value": -6,  # estimated total sold
    "actual_value": -5,  # reported sales
    "avg_penetration_pct": -4,
    "quarterly_sales": -3,
    "reported_sales": -2,
    "growth_rate_pct": -1,
    "forecast_max": 0,
    "forecast_min": 1,
    "sales_captured_in_db_pct": 2,
}

REGRESSION_FALLBACK_OFFSETS = {
    "num_quarters_used": -8,
    "forecast_value": -6,  # TOT FCST w/o SA
    "actual_value": -5,
    "forecast_max": 0,
    "forecast_min": 1,
}


@dataclass
class FileMeta:
    model: str
    ticker: str
    model_period: str
    model_date: str | None


def normalize_label(value: Any) -> str:
    if value is None:
        return ""
    text = str(value).strip().lower()
    text = re.sub(r"[%/]", " ", text)
    text = re.sub(r"[^a-z0-9]+", "_", text)
    return text.strip("_")


def is_blank(value: Any) -> bool:
    return value is None or (isinstance(value, str) and value.strip() == "")


def to_1d(values: Any) -> list[Any]:
    if values is None:
        return []
    if isinstance(values, list):
        if values and isinstance(values[0], list):
            return [row[0] if row else None for row in values]
        return values
    return [values]


def to_2d(values: Any) -> list[list[Any]]:
    if values is None:
        return []
    if isinstance(values, list):
        if values and isinstance(values[0], list):
            return values
        return [values]
    return [[values]]


def numeric(value: Any) -> float | None:
    if isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        return float(value)
    if isinstance(value, str):
        text = value.strip().replace(",", "")
        if not text:
            return None
        try:
            return float(text)
        except ValueError:
            return None
    return None


def safe_diff(a: Any, b: Any) -> float | None:
    a_num = numeric(a)
    b_num = numeric(b)
    if a_num is None or b_num is None:
        return None
    return a_num - b_num


def parse_file_meta(file_name: str) -> FileMeta:
    stem = Path(file_name).stem
    parts = [part.strip() for part in stem.split(" - ")]

    ticker = ""
    if len(parts) >= 2:
        ticker = parts[1].upper()
    else:
        ticker_match = re.search(r"\b([A-Z]{2,6})\b", stem)
        ticker = ticker_match.group(1).upper() if ticker_match else "UNKNOWN"

    period_source = parts[2] if len(parts) >= 3 else stem
    period_source = period_source.split("_")[0]
    period_match = re.search(r"(Early|Mid|Late)([A-Za-z]{3,9})(20\d{2})", period_source, re.IGNORECASE)

    if period_match:
        phase = period_match.group(1).title()
        month_token = period_match.group(2).title()
        month_key = month_token[:3].lower()
        year = int(period_match.group(3))
        model_period = f"{phase}{month_token[:3]}_{year}"
        month_number = MONTHS.get(month_key)
        day = PHASE_DAY.get(phase.lower())
        model_date = date(year, month_number, day).isoformat() if month_number and day else None
    else:
        model_period = "UNKNOWN_PERIOD"
        model_date = None

    model = f"{ticker}_{model_period}"
    return FileMeta(model=model, ticker=ticker, model_period=model_period, model_date=model_date)


def unique_output_path(in_dir: Path, out_dir: Path) -> Path:
    folder_name = in_dir.name or "input"
    stem = f"{folder_name}_PARAM"
    candidate = out_dir / f"{stem}.xlsx"
    suffix = 1
    while candidate.exists():
        candidate = out_dir / f"{stem}.{suffix}.xlsx"
        suffix += 1
    return candidate


def get_sheet(wb: xw.Book, name: str) -> xw.Sheet | None:
    try:
        return wb.sheets[name]
    except Exception:
        return None


def close_workbook_safely(wb: xw.Book) -> None:
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
        print(f"warning: failed to close workbook safely: {wb.name} ({exc})")


def find_max_anchor(sheet: xw.Sheet) -> tuple[int, int] | None:
    used = sheet.used_range
    values = to_2d(used.value)
    if not values:
        return None

    for row_index, row_values in enumerate(values, start=used.row):
        for col_index, value in enumerate(row_values, start=used.column):
            if isinstance(value, str) and value.strip().lower() == "max":
                return row_index, col_index
    return None


def build_header_map(sheet: xw.Sheet, header_row: int, anchor_col: int) -> dict[str, int]:
    min_col = max(1, anchor_col - 25)
    max_col = anchor_col + 25
    row_values = to_1d(sheet.range((header_row, min_col), (header_row, max_col)).value)
    header_map: dict[str, int] = {}
    for offset, value in enumerate(row_values):
        key = normalize_label(value)
        if key:
            header_map[key] = min_col + offset
    return header_map


def resolve_col(
    header_map: dict[str, int],
    aliases: list[str],
    anchor_col: int,
    fallback_offsets: dict[str, int],
    fallback_key: str,
) -> int:
    for alias in aliases:
        key = normalize_label(alias)
        if key in header_map:
            return header_map[key]
    return anchor_col + fallback_offsets[fallback_key]


def read_cell(sheet: xw.Sheet, row: int, col: int) -> Any:
    if row < 1 or col < 1:
        return None
    return sheet.range((row, col)).value


def collect_candidate_rows(values: list[Any], start_row: int, max_rows: int) -> list[int]:
    rows: list[int] = []
    seen_data = False
    for idx, value in enumerate(values, start=start_row):
        if is_blank(value):
            if seen_data:
                break
            continue
        seen_data = True
        rows.append(idx)
        if len(rows) >= max_rows:
            break
    return rows


def detect_candidate_rows(sheet: xw.Sheet, anchor_row: int, anchor_col: int, n_rows: int) -> list[int]:
    down_end = anchor_row + n_rows + 25
    down_values = to_1d(sheet.range((anchor_row + 1, anchor_col), (down_end, anchor_col)).value)
    down_rows = collect_candidate_rows(down_values, anchor_row + 1, n_rows)
    if down_rows:
        return down_rows

    if anchor_row <= 1:
        return [anchor_row + i for i in range(1, n_rows + 1)]

    up_start = max(1, anchor_row - (n_rows + 25))
    up_values = to_1d(sheet.range((up_start, anchor_col), (anchor_row - 1, anchor_col)).value)
    up_rows = collect_candidate_rows(up_values, up_start, n_rows)
    if up_rows:
        return up_rows[-n_rows:]

    return [anchor_row + i for i in range(1, n_rows + 1)]


def get_empirical_rows(sheet: xw.Sheet, wb: xw.Book, meta: FileMeta, source_file: str) -> list[dict[str, Any]]:
    anchor = find_max_anchor(sheet)
    if anchor is None:
        print(f"skipped empirical extraction for {source_file}: 'max' anchor not found")
        return []

    anchor_row, anchor_col = anchor
    header_map = build_header_map(sheet, anchor_row, anchor_col)
    row_ids = detect_candidate_rows(sheet, anchor_row, anchor_col, N_QUARTERS)

    num_col = resolve_col(
        header_map,
        ["num_quarters_used", "num_quarters", "quarters_used", "n_quarters"],
        anchor_col,
        EMPIRICAL_FALLBACK_OFFSETS,
        "num_quarters_used",
    )
    last_q_col = resolve_col(
        header_map,
        ["last_quarter_used", "last_quarter", "quarter_used"],
        anchor_col,
        EMPIRICAL_FALLBACK_OFFSETS,
        "last_quarter_used",
    )
    forecast_col = resolve_col(
        header_map,
        ["estimated_total_sold", "forecast_value", "forecast", "tot_fcst", "total_forecast"],
        anchor_col,
        EMPIRICAL_FALLBACK_OFFSETS,
        "forecast_value",
    )
    actual_col = resolve_col(
        header_map,
        ["reported_sales", "actual_value", "actual", "sales_reported"],
        anchor_col,
        EMPIRICAL_FALLBACK_OFFSETS,
        "actual_value",
    )
    avg_pen_col = resolve_col(
        header_map,
        ["avg_penetration_pct", "avg_penetration", "penetration_pct"],
        anchor_col,
        EMPIRICAL_FALLBACK_OFFSETS,
        "avg_penetration_pct",
    )
    q_sales_col = resolve_col(
        header_map,
        ["quarterly_sales", "qtr_sales", "quarter_sales"],
        anchor_col,
        EMPIRICAL_FALLBACK_OFFSETS,
        "quarterly_sales",
    )
    reported_sales_col = resolve_col(
        header_map,
        ["reported_sales", "sales_reported", "actual_sales"],
        anchor_col,
        EMPIRICAL_FALLBACK_OFFSETS,
        "reported_sales",
    )
    growth_col = resolve_col(
        header_map,
        ["growth_rate_pct", "growth_rate", "growth_pct"],
        anchor_col,
        EMPIRICAL_FALLBACK_OFFSETS,
        "growth_rate_pct",
    )
    max_col = resolve_col(
        header_map,
        ["max", "forecast_max"],
        anchor_col,
        EMPIRICAL_FALLBACK_OFFSETS,
        "forecast_max",
    )
    min_col = resolve_col(
        header_map,
        ["min", "forecast_min"],
        anchor_col,
        EMPIRICAL_FALLBACK_OFFSETS,
        "forecast_min",
    )
    captured_col = resolve_col(
        header_map,
        ["sales_captured_in_db_pct", "captured_in_db_pct", "db_capture_pct"],
        anchor_col,
        EMPIRICAL_FALLBACK_OFFSETS,
        "sales_captured_in_db_pct",
    )

    avg_from_formula: dict[int, Any] = {}
    if row_ids and q_sales_col > 0 and reported_sales_col > 0:
        helper_col = max(sheet.used_range.last_cell.column + 2, anchor_col + 12)
        for row in row_ids:
            # R1C1 avoids letter conversion and keeps formulas relative.
            formula = f'=IFERROR(RC[{q_sales_col - helper_col}]/RC[{reported_sales_col - helper_col}], "")'
            sheet.range((row, helper_col)).formula2 = formula

        wb.app.calculate()
        helper_values = to_1d(sheet.range((row_ids[0], helper_col), (row_ids[-1], helper_col)).value)
        for row, value in zip(row_ids, helper_values):
            avg_from_formula[row] = value
        sheet.range((row_ids[0], helper_col), (row_ids[-1], helper_col)).clear_contents()

    rows: list[dict[str, Any]] = []
    for idx, row in enumerate(row_ids, start=1):
        avg_pen = avg_from_formula.get(row, read_cell(sheet, row, avg_pen_col))
        max_value = read_cell(sheet, row, max_col)
        min_value = read_cell(sheet, row, min_col)

        num_quarters = read_cell(sheet, row, num_col)
        if is_blank(num_quarters):
            num_quarters = idx

        rows.append(
            {
                "model": meta.model,
                "ticker": meta.ticker,
                "model_period": meta.model_period,
                "model_date": meta.model_date,
                "method": "empirical",
                "parameter_name": "avg_penetration_pct",
                "parameter_value": avg_pen,
                "num_quarters_used": num_quarters,
                "last_quarter_used": read_cell(sheet, row, last_q_col),
                "forecast_value": read_cell(sheet, row, forecast_col),
                "actual_value": read_cell(sheet, row, actual_col),
                "forecast_max": max_value,
                "forecast_min": min_value,
                "range_width": safe_diff(max_value, min_value),
                "avg_penetration_pct": avg_pen,
                "quarterly_sales": read_cell(sheet, row, q_sales_col),
                "reported_sales": read_cell(sheet, row, reported_sales_col),
                "growth_rate_pct": read_cell(sheet, row, growth_col),
                "sales_captured_in_db_pct": read_cell(sheet, row, captured_col),
                "source_file": source_file,
            }
        )
    return rows


def regression_series_rows(sheet: xw.Sheet, x_col: int, y_col: int, anchor_row: int) -> list[int]:
    used = sheet.used_range
    start_row = used.row
    end_row = used.last_cell.row
    x_values = to_1d(sheet.range((start_row, x_col), (end_row, x_col)).value)
    y_values = to_1d(sheet.range((start_row, y_col), (end_row, y_col)).value)

    numeric_rows: list[int] = []
    for offset, (x_val, y_val) in enumerate(zip(x_values, y_values)):
        if numeric(x_val) is not None and numeric(y_val) is not None:
            numeric_rows.append(start_row + offset)

    if not numeric_rows:
        return []

    prior_rows = [row for row in numeric_rows if row <= anchor_row]
    if len(prior_rows) >= 2:
        return prior_rows
    return numeric_rows


def rounded_signature(values: tuple[Any, ...]) -> tuple[Any, ...]:
    out: list[Any] = []
    for value in values:
        number = numeric(value)
        if number is None:
            out.append(value)
        else:
            out.append(round(number, 10))
    return tuple(out)


def get_regression_rows(sheet: xw.Sheet, wb: xw.Book, meta: FileMeta, source_file: str) -> list[dict[str, Any]]:
    anchor = find_max_anchor(sheet)
    if anchor is None:
        print(f"skipped regression extraction for {source_file}: 'max' anchor not found")
        return []

    anchor_row, anchor_col = anchor
    header_map = build_header_map(sheet, anchor_row, anchor_col)
    candidate_rows = detect_candidate_rows(sheet, anchor_row, anchor_col, N_QUARTERS)

    y_col = anchor_col - 7
    x_col = anchor_col - 11

    num_col = resolve_col(
        header_map,
        ["num_quarters_used", "num_quarters", "quarters_used", "n_quarters"],
        anchor_col,
        REGRESSION_FALLBACK_OFFSETS,
        "num_quarters_used",
    )
    forecast_col = resolve_col(
        header_map,
        ["tot_fcst_wo_sa", "tot_fcst_w_o_sa", "tot_fcst_without_sa", "forecast_value", "forecast"],
        anchor_col,
        REGRESSION_FALLBACK_OFFSETS,
        "forecast_value",
    )
    actual_col = resolve_col(
        header_map,
        ["actual_value", "actual", "reported_sales"],
        anchor_col,
        REGRESSION_FALLBACK_OFFSETS,
        "actual_value",
    )
    max_col = resolve_col(
        header_map,
        ["max", "forecast_max"],
        anchor_col,
        REGRESSION_FALLBACK_OFFSETS,
        "forecast_max",
    )
    min_col = resolve_col(
        header_map,
        ["min", "forecast_min"],
        anchor_col,
        REGRESSION_FALLBACK_OFFSETS,
        "forecast_min",
    )

    series_rows = regression_series_rows(sheet, x_col, y_col, anchor_row)
    if len(series_rows) < 2:
        print(f"skipped regression extraction for {source_file}: insufficient x/y points")
        return []

    quarter_steps = list(range(2, min(N_QUARTERS, len(series_rows)) + 1))
    helper_col = max(sheet.used_range.last_cell.column + 2, anchor_col + 12)
    helper_row_start = anchor_row + 2

    for idx, n_quarters in enumerate(quarter_steps):
        calc_row = helper_row_start + idx
        start = series_rows[-n_quarters]
        end = series_rows[-1]

        intercept_formula = (
            f'=IFERROR(INTERCEPT(R{start}C{y_col}:R{end}C{y_col},R{start}C{x_col}:R{end}C{x_col}),"")'
        )
        slope_formula = f'=IFERROR(SLOPE(R{start}C{y_col}:R{end}C{y_col},R{start}C{x_col}:R{end}C{x_col}),"")'
        forecast_formula = f'=IF(OR(RC[-2]="",RC[-1]=""),"",RC[-2]+RC[-1]*R{end}C{x_col})'

        sheet.range((calc_row, helper_col)).formula2 = intercept_formula
        sheet.range((calc_row, helper_col + 1)).formula2 = slope_formula
        sheet.range((calc_row, helper_col + 2)).formula2 = forecast_formula

    wb.app.calculate()

    calc_matrix = to_2d(
        sheet.range(
            (helper_row_start, helper_col),
            (helper_row_start + len(quarter_steps) - 1, helper_col + 2),
        ).value
    )
    sheet.range(
        (helper_row_start, helper_col),
        (helper_row_start + len(quarter_steps) - 1, helper_col + 2),
    ).clear_contents()

    rows: list[dict[str, Any]] = []
    previous_signature: tuple[Any, ...] | None = None

    for idx, n_quarters in enumerate(quarter_steps):
        data_row = candidate_rows[idx] if idx < len(candidate_rows) else None
        intercept_value = calc_matrix[idx][0]
        slope_value = calc_matrix[idx][1]
        forecast_calc = calc_matrix[idx][2]

        sheet_forecast = read_cell(sheet, data_row, forecast_col) if data_row else None
        forecast_value = sheet_forecast if not is_blank(sheet_forecast) else forecast_calc

        max_value = read_cell(sheet, data_row, max_col) if data_row else None
        min_value = read_cell(sheet, data_row, min_col) if data_row else None
        actual_value = read_cell(sheet, data_row, actual_col) if data_row else None
        num_quarters_used = read_cell(sheet, data_row, num_col) if data_row else n_quarters
        if is_blank(num_quarters_used):
            num_quarters_used = n_quarters

        signature = rounded_signature((intercept_value, slope_value, forecast_value, max_value, min_value))
        if previous_signature == signature:
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
                "parameter_value": num_quarters_used,
                "num_quarters_used": num_quarters_used,
                "forecast_value": forecast_value,
                "actual_value": actual_value,
                "forecast_max": max_value,
                "forecast_min": min_value,
                "range_width": safe_diff(max_value, min_value),
                "intercept": intercept_value,
                "slope": slope_value,
                "source_file": source_file,
            }
        )
    return rows


def write_sheet(sheet, columns: list[str], rows: list[dict[str, Any]]) -> None:
    sheet.append(columns)
    for row in rows:
        sheet.append([row.get(column) for column in columns])

    for header_cell in sheet[1]:
        header_cell.font = Font(bold=True)

    sheet.freeze_panes = "A2"
    sheet.auto_filter.ref = sheet.dimensions

    for col_idx, col_name in enumerate(columns, start=1):
        max_len = len(col_name)
        for row_idx in range(2, sheet.max_row + 1):
            value = sheet.cell(row=row_idx, column=col_idx).value
            if value is None:
                continue
            max_len = max(max_len, len(str(value)))
        sheet.column_dimensions[get_column_letter(col_idx)].width = min(max_len + 2, 48)


def write_output(path: Path, empirical_rows: list[dict[str, Any]], regression_rows: list[dict[str, Any]]) -> None:
    wb = Workbook()
    default_sheet = wb.active
    wb.remove(default_sheet)

    empirical_sheet = wb.create_sheet(OUTPUT_EMPIRICAL_SHEET)
    regression_sheet = wb.create_sheet(OUTPUT_REGRESSION_SHEET)

    write_sheet(empirical_sheet, EMPIRICAL_COLUMNS, empirical_rows)
    write_sheet(regression_sheet, REGRESSION_COLUMNS, regression_rows)
    wb.save(path)


def should_skip(file_path: Path) -> str | None:
    name = file_path.name
    if name.startswith("~"):
        return "temp file"
    if file_path.suffix.lower() != ".xlsx":
        return "not .xlsx"
    if re.search(r"_param(?:\.\d+)?\.xlsx$", name, flags=re.IGNORECASE):
        return "generated output file"
    return None


def main() -> None:
    in_dir = input_dir.expanduser().resolve()
    out_dir = output_dir.expanduser().resolve()

    if not in_dir.exists():
        raise FileNotFoundError(f"input_dir does not exist: {in_dir}")

    out_dir.mkdir(parents=True, exist_ok=True)
    out_path = unique_output_path(in_dir, out_dir)

    empirical_rows: list[dict[str, Any]] = []
    regression_rows: list[dict[str, Any]] = []
    processed_files = 0

    app = xw.App(visible=False, add_book=False)
    app.display_alerts = False
    app.screen_updating = False

    try:
        for file_path in sorted(in_dir.iterdir()):
            if not file_path.is_file():
                continue

            skip_reason = should_skip(file_path)
            if skip_reason:
                print(f"skipped: {file_path.name} ({skip_reason})")
                continue

            print(f"processing: {file_path.name}")
            wb = None
            try:
                wb = app.books.open(str(file_path), update_links=False)
                meta = parse_file_meta(file_path.name)

                empirical_sheet = get_sheet(wb, EMPIRICAL_SHEET_NAME)
                if empirical_sheet is None:
                    print(f"skipped empirical for {file_path.name}: missing sheet '{EMPIRICAL_SHEET_NAME}'")
                else:
                    empirical_rows.extend(get_empirical_rows(empirical_sheet, wb, meta, file_path.name))

                regression_sheet = get_sheet(wb, REGRESSION_SHEET_NAME)
                if regression_sheet is None:
                    print(f"skipped regression for {file_path.name}: missing sheet '{REGRESSION_SHEET_NAME}'")
                else:
                    regression_rows.extend(get_regression_rows(regression_sheet, wb, meta, file_path.name))

                processed_files += 1
            except Exception as exc:
                print(f"skipped: {file_path.name} (error: {exc})")
            finally:
                if wb is not None:
                    close_workbook_safely(wb)
    finally:
        app.quit()

    write_output(out_path, empirical_rows, regression_rows)

    print(f"output_path: {out_path}")
    print(f"files_processed: {processed_files}")
    print(f"empirical_rows: {len(empirical_rows)}")
    print(f"regression_rows: {len(regression_rows)}")


if __name__ == "__main__":
    main()
