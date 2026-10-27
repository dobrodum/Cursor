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
# User-configurable paths
# -----------------------------
input_dir = r"./input"
output_dir = r"./output"


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


MONTH_TO_NUM = {
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

PERIOD_DAY = {
    "early": 5,
    "mid": 15,
    "late": 25,
}


@dataclass
class ModelMeta:
    model: str
    ticker: str
    model_period: str
    model_date: Any


@dataclass
class SheetSnapshot:
    start_row: int
    start_col: int
    end_row: int
    end_col: int
    values: List[List[Any]]

    def get(self, row: int, col: int) -> Any:
        if row < self.start_row or row > self.end_row:
            return None
        if col < self.start_col or col > self.end_col:
            return None
        return self.values[row - self.start_row][col - self.start_col]


def normalize_text(value: Any) -> str:
    if value is None:
        return ""
    return str(value).strip().lower()


def ensure_2d(values: Any) -> List[List[Any]]:
    if not isinstance(values, list):
        return [[values]]
    if not values:
        return [[]]
    if isinstance(values[0], list):
        return values
    return [values]


def to_float(value: Any) -> Optional[float]:
    if value is None:
        return None
    if isinstance(value, (int, float)):
        return float(value)
    text = str(value).strip().replace(",", "")
    if not text:
        return None
    if text.endswith("%"):
        try:
            return float(text[:-1]) / 100.0
        except ValueError:
            return None
    try:
        return float(text)
    except ValueError:
        return None


def numeric_or_default(value: Any, default: Any) -> Any:
    number = to_float(value)
    if number is None:
        return default
    if float(number).is_integer():
        return int(number)
    return number


def subtract(a: Any, b: Any) -> Any:
    a_val = to_float(a)
    b_val = to_float(b)
    if a_val is None or b_val is None:
        return ""
    return a_val - b_val


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
    except Exception:
        wb.api.Close(False)


def set_formula2(cell: xw.Range, formula_r1c1: str) -> None:
    try:
        cell.formula2 = formula_r1c1
    except Exception:
        # Fallback keeps the workflow alive on older Excel builds.
        cell.formula = formula_r1c1


def parse_model_meta(file_name: str) -> ModelMeta:
    stem = Path(file_name).stem
    parts = [part.strip() for part in stem.split(" - ") if part.strip()]

    ticker = parts[1] if len(parts) >= 2 else ""
    period_source = parts[2] if len(parts) >= 3 else stem
    period_source = period_source.split("_")[0].replace(" ", "")

    match = re.search(
        r"(Early|Mid|Late)([A-Za-z]{3,9})(\d{4})",
        period_source,
        re.IGNORECASE,
    )

    if not match:
        model_period = period_source or "unknown_period"
        model_date: Any = ""
        model = f"{ticker}_{model_period}" if ticker else model_period
        return ModelMeta(model=model, ticker=ticker, model_period=model_period, model_date=model_date)

    timing = match.group(1).lower()
    month_abbr = match.group(2)[:3].title()
    year = int(match.group(3))
    month_num = MONTH_TO_NUM.get(month_abbr.lower())
    day = PERIOD_DAY.get(timing, 15)

    model_period = f"{match.group(1).title()}{month_abbr}_{year}"
    if month_num:
        model_date = date(year, month_num, day)
    else:
        model_date = ""
    model = f"{ticker}_{model_period}" if ticker else model_period

    return ModelMeta(model=model, ticker=ticker, model_period=model_period, model_date=model_date)


def get_sheet_case_insensitive(wb: xw.Book, sheet_name: str) -> Optional[xw.Sheet]:
    target = sheet_name.strip().lower()
    for sheet in wb.sheets:
        if sheet.name.strip().lower() == target:
            return sheet
    return None


def build_sheet_snapshot(sheet: xw.Sheet) -> SheetSnapshot:
    used = sheet.used_range
    start_row = used.row
    start_col = used.column
    values = ensure_2d(used.value)
    row_count = len(values)
    col_count = max((len(row) for row in values), default=1)

    normalized: List[List[Any]] = []
    for row in values:
        if len(row) < col_count:
            row = row + [None] * (col_count - len(row))
        normalized.append(row)

    end_row = start_row + row_count - 1
    end_col = start_col + col_count - 1

    return SheetSnapshot(
        start_row=start_row,
        start_col=start_col,
        end_row=end_row,
        end_col=end_col,
        values=normalized,
    )


def find_max_anchor(snapshot: SheetSnapshot) -> Tuple[int, int]:
    candidates: List[Tuple[int, int, int]] = []
    for r_idx, row in enumerate(snapshot.values):
        for c_idx, value in enumerate(row):
            if normalize_text(value) != "max":
                continue
            row_num = snapshot.start_row + r_idx
            col_num = snapshot.start_col + c_idx
            right = normalize_text(snapshot.get(row_num, col_num + 1))
            # Prefer "max" headers where the next column is "min".
            score = 0 if right == "min" else 1
            candidates.append((score, row_num, col_num))

    if not candidates:
        raise ValueError("could not find 'max' anchor")

    candidates.sort(key=lambda item: (item[0], item[1], item[2]))
    _, row_num, col_num = candidates[0]
    return row_num, col_num


def build_header_maps(snapshot: SheetSnapshot, anchor_row: int) -> List[Dict[str, int]]:
    maps: List[Dict[str, int]] = []
    for row_num in (anchor_row - 1, anchor_row, anchor_row + 1):
        if row_num < snapshot.start_row or row_num > snapshot.end_row:
            continue
        row_values = snapshot.values[row_num - snapshot.start_row]
        row_map: Dict[str, int] = {}
        for idx, value in enumerate(row_values):
            key = normalize_text(value)
            if key:
                row_map[key] = snapshot.start_col + idx
        maps.append(row_map)
    return maps


def resolve_col(
    header_maps: Sequence[Dict[str, int]],
    synonyms: Sequence[str],
    anchor_col: int,
    fallback_offset: Optional[int],
) -> Optional[int]:
    normalized_synonyms = [normalize_text(s) for s in synonyms]
    for row_map in header_maps:
        for header, col_idx in row_map.items():
            if any(syn and (header == syn or syn in header) for syn in normalized_synonyms):
                return col_idx
    if fallback_offset is None:
        return None
    return anchor_col + fallback_offset


def read_cell(sheet: xw.Sheet, row: int, col: Optional[int]) -> Any:
    if col is None or col <= 0:
        return None
    return sheet.cells(row, col).value


def is_empty_candidate_row(values: Sequence[Any]) -> bool:
    return all(value in (None, "") for value in values)


def extract_empirical_rows(wb: xw.Book, meta: ModelMeta, source_file: str) -> List[Dict[str, Any]]:
    sheet = get_sheet_case_insensitive(wb, "Empirical Model")
    if sheet is None:
        return []

    snapshot = build_sheet_snapshot(sheet)
    anchor_row, anchor_col = find_max_anchor(snapshot)
    header_maps = build_header_maps(snapshot, anchor_row)

    col_num_quarters = resolve_col(
        header_maps,
        ["num quarters", "# quarters", "quarters used", "n quarters"],
        anchor_col,
        -9,
    )
    col_last_quarter = resolve_col(
        header_maps,
        ["last quarter", "last qtr"],
        anchor_col,
        -8,
    )
    col_quarterly_sales = resolve_col(
        header_maps,
        ["quarterly sales", "qtr sales"],
        anchor_col,
        -7,
    )
    col_avg_penetration = resolve_col(
        header_maps,
        ["avg penetration", "average penetration", "penetration %"],
        anchor_col,
        -6,
    )
    col_forecast = resolve_col(
        header_maps,
        ["estimated total sold", "est total sold", "forecast", "tot fcst"],
        anchor_col,
        -5,
    )
    col_reported_sales = resolve_col(
        header_maps,
        ["reported sales", "actual sales", "actual"],
        anchor_col,
        -4,
    )
    col_growth = resolve_col(
        header_maps,
        ["growth rate", "growth rate %"],
        anchor_col,
        -3,
    )
    col_sales_captured = resolve_col(
        header_maps,
        ["sales captured in db", "captured in db", "sales captured"],
        anchor_col,
        -2,
    )
    col_forecast_max = anchor_col
    col_forecast_min = resolve_col(header_maps, ["min"], anchor_col, 1)

    temp_avg_col = snapshot.end_col + 2
    helper_rows: List[int] = []
    source_for_avg = col_sales_captured or col_avg_penetration

    # Write all helper formulas first, then calculate once.
    if source_for_avg is not None:
        for step in range(1, N_QUARTERS + 1):
            row_num = anchor_row + step
            range_start = max(anchor_row + 1, row_num - step + 1)
            formula = (
                f'=IFERROR(AVERAGE(R{range_start}C{source_for_avg}:R{row_num}C{source_for_avg}),"")'
            )
            set_formula2(sheet.cells(row_num, temp_avg_col), formula)
            helper_rows.append(row_num)
        wb.app.calculate()

    rows: List[Dict[str, Any]] = []
    empty_streak = 0
    for step in range(1, N_QUARTERS + 1):
        row_num = anchor_row + step

        num_quarters_used = numeric_or_default(read_cell(sheet, row_num, col_num_quarters), step)
        last_quarter_used = read_cell(sheet, row_num, col_last_quarter)
        quarterly_sales = read_cell(sheet, row_num, col_quarterly_sales)
        reported_sales = read_cell(sheet, row_num, col_reported_sales)
        growth_rate_pct = read_cell(sheet, row_num, col_growth)
        sales_captured_in_db_pct = read_cell(sheet, row_num, col_sales_captured)
        forecast_value = read_cell(sheet, row_num, col_forecast)
        forecast_max = read_cell(sheet, row_num, col_forecast_max)
        forecast_min = read_cell(sheet, row_num, col_forecast_min)
        avg_penetration_pct = read_cell(sheet, row_num, col_avg_penetration)

        if avg_penetration_pct in (None, "") and row_num in helper_rows:
            avg_penetration_pct = sheet.cells(row_num, temp_avg_col).value

        actual_value = reported_sales
        range_width = subtract(forecast_max, forecast_min)

        if is_empty_candidate_row(
            [
                forecast_value,
                forecast_max,
                forecast_min,
                avg_penetration_pct,
                reported_sales,
                quarterly_sales,
            ]
        ):
            empty_streak += 1
            if empty_streak >= 2:
                break
            continue
        empty_streak = 0

        rows.append(
            {
                "model": meta.model,
                "ticker": meta.ticker,
                "model_period": meta.model_period,
                "model_date": meta.model_date,
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

    if helper_rows:
        sheet.range((helper_rows[0], temp_avg_col), (helper_rows[-1], temp_avg_col)).value = None

    return rows


def rounded_signature_value(value: Any) -> Any:
    number = to_float(value)
    if number is None:
        return value
    return round(number, 10)


def extract_regression_rows(wb: xw.Book, meta: ModelMeta, source_file: str) -> List[Dict[str, Any]]:
    sheet = get_sheet_case_insensitive(wb, "Regression Model")
    if sheet is None:
        return []

    snapshot = build_sheet_snapshot(sheet)
    anchor_row, anchor_col = find_max_anchor(snapshot)
    header_maps = build_header_maps(snapshot, anchor_row)

    # Required anchor-based source columns from existing logic.
    y_col = anchor_col - 7
    x_col = anchor_col - 11

    col_num_quarters = resolve_col(
        header_maps,
        ["num quarters", "# quarters", "quarters used", "n quarters"],
        anchor_col,
        -9,
    )
    col_forecast = resolve_col(
        header_maps,
        ["tot fcst w/o sa", "tot fcst without sa", "forecast w/o sa", "forecast"],
        anchor_col,
        -2,
    )
    col_actual = resolve_col(
        header_maps,
        ["actual", "actual sales", "reported sales"],
        anchor_col,
        None,
    )
    col_forecast_max = anchor_col
    col_forecast_min = resolve_col(header_maps, ["min"], anchor_col, 1)
    col_intercept_existing = resolve_col(header_maps, ["intercept"], anchor_col, -4)
    col_slope_existing = resolve_col(header_maps, ["slope"], anchor_col, -3)

    temp_intercept_col = snapshot.end_col + 2
    temp_slope_col = snapshot.end_col + 3
    helper_rows: List[int] = []

    for step in range(1, N_QUARTERS + 1):
        row_num = anchor_row + step
        num_quarters_used = numeric_or_default(read_cell(sheet, row_num, col_num_quarters), step)
        n_value = max(1, int(to_float(num_quarters_used) or step))
        range_start = max(anchor_row + 1, row_num - n_value + 1)

        intercept_formula = (
            f'=IFERROR(INTERCEPT(R{range_start}C{y_col}:R{row_num}C{y_col},'
            f'R{range_start}C{x_col}:R{row_num}C{x_col}),"")'
        )
        slope_formula = (
            f'=IFERROR(SLOPE(R{range_start}C{y_col}:R{row_num}C{y_col},'
            f'R{range_start}C{x_col}:R{row_num}C{x_col}),"")'
        )
        set_formula2(sheet.cells(row_num, temp_intercept_col), intercept_formula)
        set_formula2(sheet.cells(row_num, temp_slope_col), slope_formula)
        helper_rows.append(row_num)

    if helper_rows:
        wb.app.calculate()

    rows: List[Dict[str, Any]] = []
    prev_signature: Optional[Tuple[Any, ...]] = None
    empty_streak = 0

    for step in range(1, N_QUARTERS + 1):
        row_num = anchor_row + step
        num_quarters_used = numeric_or_default(read_cell(sheet, row_num, col_num_quarters), step)
        forecast_value = read_cell(sheet, row_num, col_forecast)
        actual_value = read_cell(sheet, row_num, col_actual) if col_actual else ""
        forecast_max = read_cell(sheet, row_num, col_forecast_max)
        forecast_min = read_cell(sheet, row_num, col_forecast_min)

        intercept_value = sheet.cells(row_num, temp_intercept_col).value
        slope_value = sheet.cells(row_num, temp_slope_col).value

        if intercept_value in (None, ""):
            intercept_value = read_cell(sheet, row_num, col_intercept_existing)
        if slope_value in (None, ""):
            slope_value = read_cell(sheet, row_num, col_slope_existing)

        if is_empty_candidate_row(
            [forecast_value, forecast_max, forecast_min, intercept_value, slope_value]
        ):
            empty_streak += 1
            if empty_streak >= 2:
                break
            continue
        empty_streak = 0

        range_width = subtract(forecast_max, forecast_min)

        signature = (
            rounded_signature_value(num_quarters_used),
            rounded_signature_value(forecast_value),
            rounded_signature_value(forecast_max),
            rounded_signature_value(forecast_min),
            rounded_signature_value(intercept_value),
            rounded_signature_value(slope_value),
        )
        if prev_signature is not None and signature == prev_signature:
            # Prevent duplicate final rows generated by repeated formulas.
            continue
        prev_signature = signature

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
                "forecast_max": forecast_max,
                "forecast_min": forecast_min,
                "range_width": range_width,
                "intercept": intercept_value,
                "slope": slope_value,
                "source_file": source_file,
            }
        )

    if helper_rows:
        sheet.range((helper_rows[0], temp_intercept_col), (helper_rows[-1], temp_slope_col)).value = None

    return rows


def build_output_path(input_path: Path, out_path: Path) -> Path:
    base_name = f"{input_path.name}_PARAM"
    candidate = out_path / f"{base_name}.xlsx"
    if not candidate.exists():
        return candidate

    suffix = 1
    while True:
        candidate = out_path / f"{base_name}.{suffix}.xlsx"
        if not candidate.exists():
            return candidate
        suffix += 1


def write_rows_to_sheet(ws, columns: Sequence[str], rows: Sequence[Dict[str, Any]]) -> None:
    ws.append(list(columns))
    for row in rows:
        ws.append([row.get(col, "") for col in columns])

    for cell in ws[1]:
        cell.font = Font(bold=True)

    ws.freeze_panes = "A2"

    if ws.max_row >= 1:
        end_col = get_column_letter(len(columns))
        ws.auto_filter.ref = f"A1:{end_col}{ws.max_row}"

    for col_idx, column_name in enumerate(columns, start=1):
        max_len = len(column_name)
        for row_idx in range(2, ws.max_row + 1):
            value = ws.cell(row=row_idx, column=col_idx).value
            if value is None:
                continue
            max_len = max(max_len, len(str(value)))
        ws.column_dimensions[get_column_letter(col_idx)].width = min(max_len + 2, 48)


def build_output_workbook(
    output_path: Path,
    empirical_rows: Sequence[Dict[str, Any]],
    regression_rows: Sequence[Dict[str, Any]],
) -> None:
    output_wb = Workbook()
    output_wb.remove(output_wb.active)

    ws_empirical = output_wb.create_sheet("empirical_candidates")
    ws_regression = output_wb.create_sheet("regression_candidates")

    write_rows_to_sheet(ws_empirical, EMPIRICAL_COLUMNS, empirical_rows)
    write_rows_to_sheet(ws_regression, REGRESSION_COLUMNS, regression_rows)

    output_wb.save(output_path)


def process_input_files(input_path: Path) -> Tuple[int, List[Dict[str, Any]], List[Dict[str, Any]]]:
    empirical_rows: List[Dict[str, Any]] = []
    regression_rows: List[Dict[str, Any]] = []
    processed_files = 0

    excel = xw.App(visible=False, add_book=False)
    excel.display_alerts = False
    excel.screen_updating = False
    try:
        excel.calculation = "manual"
    except Exception:
        pass

    try:
        entries = sorted(input_path.iterdir(), key=lambda p: p.name.lower())
        for file_path in entries:
            if not file_path.is_file():
                continue

            if file_path.name.startswith("~"):
                print(f"Skipped {file_path.name}: temp file")
                continue
            if file_path.suffix.lower() != ".xlsx":
                print(f"Skipped {file_path.name}: not an .xlsx file")
                continue

            print(f"Processing {file_path.name}")

            wb: Optional[xw.Book] = None
            try:
                wb = excel.books.open(str(file_path), update_links=False)
                meta = parse_model_meta(file_path.name)

                empirical_part = extract_empirical_rows(wb, meta, file_path.name)
                regression_part = extract_regression_rows(wb, meta, file_path.name)

                empirical_rows.extend(empirical_part)
                regression_rows.extend(regression_part)
                processed_files += 1
            except Exception as exc:
                print(f"Skipped {file_path.name}: {exc}")
            finally:
                if wb is not None:
                    close_workbook_safely(wb)
    finally:
        excel.quit()

    return processed_files, empirical_rows, regression_rows


def main() -> None:
    input_path = Path(input_dir).expanduser().resolve()
    output_path = Path(output_dir).expanduser().resolve()

    if not input_path.exists() or not input_path.is_dir():
        raise FileNotFoundError(f"Input folder does not exist: {input_path}")

    output_path.mkdir(parents=True, exist_ok=True)
    output_file = build_output_path(input_path, output_path)

    processed_files, empirical_rows, regression_rows = process_input_files(input_path)
    build_output_workbook(output_file, empirical_rows, regression_rows)

    print(f"Output written: {output_file}")
    print(f"Files processed: {processed_files}")
    print(f"Empirical rows: {len(empirical_rows)}")
    print(f"Regression rows: {len(regression_rows)}")


if __name__ == "__main__":
    main()
