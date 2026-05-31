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

# Configure these folders before running.
input_dir = Path("./input")
output_dir = Path("./output")

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

EARLY_MID_LATE_DAY = {"early": 5, "mid": 15, "late": 25}
MONTH_NUMBER = {
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


@dataclass
class FileMetadata:
    model: str
    ticker: str
    model_period: str
    model_date: str


def normalize_text(value: Any) -> str:
    if value is None:
        return ""
    text = str(value).strip().lower()
    text = re.sub(r"[\s_\-]+", " ", text)
    return text


def to_2d(values: Any) -> List[List[Any]]:
    if values is None:
        return []
    if not isinstance(values, list):
        return [[values]]
    if values and not isinstance(values[0], list):
        return [values]
    return values


def coerce_float(value: Any) -> Optional[float]:
    if value is None:
        return None
    if isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        return float(value)
    text = str(value).strip()
    if not text:
        return None
    text = text.replace(",", "")
    if text.endswith("%"):
        try:
            return float(text[:-1]) / 100.0
        except ValueError:
            return None
    try:
        return float(text)
    except ValueError:
        return None


def coerce_int(value: Any) -> Optional[int]:
    as_float = coerce_float(value)
    if as_float is None:
        return None
    return int(round(as_float))


def clean_output_value(value: Any) -> Any:
    if isinstance(value, str):
        value = value.strip()
        return value if value else None
    return value


def get_unique_output_path(input_path: Path, output_path: Path) -> Path:
    output_path.mkdir(parents=True, exist_ok=True)
    stem = f"{input_path.name}_PARAM"
    candidate = output_path / f"{stem}.xlsx"
    suffix = 1
    while candidate.exists():
        candidate = output_path / f"{stem}.{suffix}.xlsx"
        suffix += 1
    return candidate


def parse_file_metadata(file_name: str) -> FileMetadata:
    stem = Path(file_name).stem
    parts = [part.strip() for part in stem.split(" - ")]
    ticker = parts[1].strip().upper() if len(parts) >= 2 else "UNKNOWN"
    period_part = parts[2].strip() if len(parts) >= 3 else stem
    period_token = re.split(r"[_\s]+", period_part)[0]
    period_match = re.search(
        r"(Early|Mid|Late)\s*([A-Za-z]+)\s*(\d{4})", period_token, flags=re.IGNORECASE
    )

    if not period_match:
        period_match = re.search(
            r"(Early|Mid|Late)\s*([A-Za-z]+)\s*(\d{4})", stem, flags=re.IGNORECASE
        )

    if period_match:
        day_key = period_match.group(1).lower()
        month_token = period_match.group(2)[:3].lower()
        year = int(period_match.group(3))
        month_num = MONTH_NUMBER.get(month_token, 1)
        month_label = month_token.title()
        day = EARLY_MID_LATE_DAY[day_key]
        model_period = f"{day_key.title()}{month_label}_{year}"
        model_date = date(year, month_num, day).isoformat()
    else:
        model_period = "Unknown_0000"
        model_date = ""

    model = f"{ticker}_{model_period}"
    return FileMetadata(
        model=model, ticker=ticker, model_period=model_period, model_date=model_date
    )


def safe_close_source_workbook(wb: xw.Book) -> None:
    # Never save source files. Try several close signatures for compatibility.
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
        try:
            wb.close()
        except Exception:
            pass


def get_sheet_by_name(wb: xw.Book, target_name: str) -> Optional[xw.Sheet]:
    normalized_target = normalize_text(target_name)
    for sheet in wb.sheets:
        if normalize_text(sheet.name) == normalized_target:
            return sheet
    return None


def find_anchor_cell(sheet: xw.Sheet, anchor_text: str = "max") -> Optional[Tuple[int, int]]:
    used = sheet.used_range
    values = to_2d(used.value)
    if not values:
        return None
    top_row, left_col = used.row, used.column
    anchor_norm = normalize_text(anchor_text)
    for r_idx, row in enumerate(values):
        for c_idx, value in enumerate(row):
            if normalize_text(value) == anchor_norm:
                return top_row + r_idx, left_col + c_idx
    return None


def infer_header_map(
    sheet: xw.Sheet, anchor_row: int, anchor_col: int, col_window: int = 20
) -> Tuple[int, Dict[int, str]]:
    best_row = anchor_row
    best_score = -1
    start_col = max(1, anchor_col - col_window)
    end_col = anchor_col + col_window

    for row in range(max(1, anchor_row - 3), anchor_row + 4):
        row_values = to_2d(sheet.range((row, start_col), (row, end_col)).value)[0]
        normalized = [normalize_text(v) for v in row_values]
        score = sum(1 for item in normalized if item) + (25 if "max" in normalized else 0)
        if score > best_score:
            best_score = score
            best_row = row

    header_values = to_2d(sheet.range((best_row, start_col), (best_row, end_col)).value)[0]
    header_map: Dict[int, str] = {}
    for idx, value in enumerate(header_values):
        text = normalize_text(value)
        if text:
            header_map[start_col + idx] = text
    return best_row, header_map


def pick_column(
    header_map: Dict[int, str], patterns: Sequence[Sequence[str]]
) -> Optional[int]:
    for terms in patterns:
        for col, text in header_map.items():
            if all(term in text for term in terms):
                return col
    return None


def fallback_column(
    preferred: Optional[int], anchor_col: int, offset: int, min_col: int = 1
) -> int:
    if preferred is not None:
        return preferred
    return max(min_col, anchor_col + offset)


def read_rows_block(
    sheet: xw.Sheet,
    start_row: int,
    end_row: int,
    min_col: int,
    max_col: int,
) -> List[List[Any]]:
    if end_row < start_row or max_col < min_col:
        return []
    return to_2d(sheet.range((start_row, min_col), (end_row, max_col)).value)


def detect_data_rows(
    sheet: xw.Sheet,
    start_row: int,
    stop_row: int,
    key_cols: Sequence[int],
) -> List[int]:
    if not key_cols:
        return []
    min_col = min(key_cols)
    max_col = max(key_cols)
    block = read_rows_block(sheet, start_row, stop_row, min_col, max_col)
    rows: List[int] = []
    blank_run = 0

    for idx, row_values in enumerate(block):
        row_number = start_row + idx
        has_data = False
        for col in key_cols:
            value = row_values[col - min_col]
            if value not in (None, ""):
                has_data = True
                break

        if has_data:
            rows.append(row_number)
            blank_run = 0
        else:
            blank_run += 1
            if rows and blank_run >= 6:
                break
    return rows


def set_formula2(cell: xw.Range, formula: str) -> None:
    try:
        cell.formula2 = formula
    except Exception:
        cell.formula = formula


def get_cell_value(sheet: xw.Sheet, row: int, col: int) -> Any:
    if row < 1 or col < 1:
        return None
    return sheet.range((row, col)).value


def extract_empirical_rows(
    wb: xw.Book, metadata: FileMetadata, source_file: str
) -> List[Dict[str, Any]]:
    sheet = get_sheet_by_name(wb, "Empirical Model")
    if sheet is None:
        print(f"Skipped file {source_file}: missing sheet 'Empirical Model'")
        return []

    anchor = find_anchor_cell(sheet, "max")
    if anchor is None:
        print(f"Skipped file {source_file}: could not find 'max' anchor in Empirical Model")
        return []

    anchor_row, anchor_col = anchor
    header_row, header_map = infer_header_map(sheet, anchor_row, anchor_col)

    col_num_q = fallback_column(
        pick_column(
            header_map,
            [
                ("num", "quarter"),
                ("quarter", "used"),
                ("n", "quarter"),
            ],
        ),
        anchor_col,
        -10,
    )
    col_last_q = fallback_column(
        pick_column(header_map, [(("last", "quarter")), (("last", "qtr"))]),
        anchor_col,
        -9,
    )
    col_forecast = fallback_column(
        pick_column(
            header_map,
            [
                ("estimated", "total", "sold"),
                ("forecast", "value"),
                ("forecast", "total"),
                ("tot", "fcst"),
            ],
        ),
        anchor_col,
        -2,
    )
    col_actual = fallback_column(
        pick_column(
            header_map,
            [
                ("reported", "sales"),
                ("actual", "sales"),
                ("actual", "value"),
            ],
        ),
        anchor_col,
        -3,
    )
    col_max = fallback_column(
        pick_column(header_map, [("max",)]), anchor_col, 0, min_col=anchor_col
    )
    col_min = fallback_column(
        pick_column(header_map, [("min",)]), anchor_col, 1
    )
    col_avg_pen = fallback_column(
        pick_column(
            header_map,
            [("avg", "penetration"), ("average", "penetration"), ("penetration",)],
        ),
        anchor_col,
        -5,
    )
    col_quarterly_sales = fallback_column(
        pick_column(header_map, [("quarterly", "sales"), ("quarter", "sales")]),
        anchor_col,
        -7,
    )
    col_reported_sales = fallback_column(
        pick_column(header_map, [("reported", "sales"), ("report", "sales")]),
        anchor_col,
        -3,
    )
    col_growth = fallback_column(
        pick_column(header_map, [("growth", "rate"), ("growth",)]),
        anchor_col,
        -6,
    )
    col_captured = fallback_column(
        pick_column(
            header_map,
            [
                ("captured", "db"),
                ("sales", "captured"),
                ("captured",),
            ],
        ),
        anchor_col,
        -4,
    )

    key_cols = [col_num_q, col_forecast, col_actual, col_max, col_min, col_avg_pen]
    data_rows = detect_data_rows(
        sheet=sheet,
        start_row=header_row + 1,
        stop_row=header_row + 160,
        key_cols=key_cols,
    )
    if not data_rows:
        data_rows = [header_row + offset for offset in range(1, 11)]

    # Existing empirical logic iterates up to 10 quarter windows.
    data_rows = data_rows[:10]
    if not data_rows:
        return []

    helper_col = max(anchor_col + 2, col_max + 2, col_min + 2)
    helper_start_row = anchor_row + 220
    formula_cells: List[Tuple[int, xw.Range]] = []
    first_data_row = data_rows[0]

    for idx, row in enumerate(data_rows):
        n_quarters = coerce_int(get_cell_value(sheet, row, col_num_q)) or (idx + 1)
        n_quarters = max(1, min(10, n_quarters))
        start_row = max(first_data_row, row - n_quarters + 1)
        formula = f"=AVERAGE(R{start_row}C{col_avg_pen}:R{row}C{col_avg_pen})"
        helper_cell = sheet.range((helper_start_row + idx, helper_col))
        set_formula2(helper_cell, formula)
        formula_cells.append((row, helper_cell))

    if formula_cells:
        wb.app.calculate()

    rows: List[Dict[str, Any]] = []
    for idx, (row, helper_cell) in enumerate(formula_cells):
        num_quarters = coerce_int(get_cell_value(sheet, row, col_num_q)) or (idx + 1)
        last_quarter = clean_output_value(get_cell_value(sheet, row, col_last_q))
        forecast_value = coerce_float(get_cell_value(sheet, row, col_forecast))
        actual_value = coerce_float(get_cell_value(sheet, row, col_actual))
        forecast_max = coerce_float(get_cell_value(sheet, row, col_max))
        forecast_min = coerce_float(get_cell_value(sheet, row, col_min))
        avg_penetration = coerce_float(helper_cell.value)
        if avg_penetration is None:
            avg_penetration = coerce_float(get_cell_value(sheet, row, col_avg_pen))
        quarterly_sales = coerce_float(get_cell_value(sheet, row, col_quarterly_sales))
        reported_sales = coerce_float(get_cell_value(sheet, row, col_reported_sales))
        growth_rate = coerce_float(get_cell_value(sheet, row, col_growth))
        captured_pct = coerce_float(get_cell_value(sheet, row, col_captured))
        range_width = (
            forecast_max - forecast_min
            if forecast_max is not None and forecast_min is not None
            else None
        )

        if all(
            value is None
            for value in (
                forecast_value,
                actual_value,
                forecast_max,
                forecast_min,
                avg_penetration,
            )
        ):
            continue

        rows.append(
            {
                "model": metadata.model,
                "ticker": metadata.ticker,
                "model_period": metadata.model_period,
                "model_date": metadata.model_date,
                "method": "empirical",
                "parameter_name": "avg_penetration_pct",
                "parameter_value": avg_penetration,
                "num_quarters_used": num_quarters,
                "last_quarter_used": last_quarter,
                "forecast_value": forecast_value,
                "actual_value": actual_value,
                "forecast_max": forecast_max,
                "forecast_min": forecast_min,
                "range_width": range_width,
                "avg_penetration_pct": avg_penetration,
                "quarterly_sales": quarterly_sales,
                "reported_sales": reported_sales,
                "growth_rate_pct": growth_rate,
                "sales_captured_in_db_pct": captured_pct,
                "source_file": source_file,
            }
        )

    return rows


def get_xy_history(sheet: xw.Sheet, anchor_row: int, x_col: int, y_col: int) -> List[Tuple[int, float, float]]:
    if x_col < 1 or y_col < 1 or anchor_row <= 1:
        return []
    min_col = min(x_col, y_col)
    max_col = max(x_col, y_col)
    block = read_rows_block(sheet, 1, anchor_row - 1, min_col, max_col)
    history: List[Tuple[int, float, float]] = []
    for idx, row_values in enumerate(block):
        row_num = idx + 1
        x_value = coerce_float(row_values[x_col - min_col])
        y_value = coerce_float(row_values[y_col - min_col])
        if x_value is None or y_value is None:
            continue
        history.append((row_num, x_value, y_value))
    return history


def extract_regression_rows(
    wb: xw.Book, metadata: FileMetadata, source_file: str
) -> List[Dict[str, Any]]:
    sheet = get_sheet_by_name(wb, "Regression Model")
    if sheet is None:
        print(f"Skipped file {source_file}: missing sheet 'Regression Model'")
        return []

    anchor = find_anchor_cell(sheet, "max")
    if anchor is None:
        print(f"Skipped file {source_file}: could not find 'max' anchor in Regression Model")
        return []

    anchor_row, anchor_col = anchor
    header_row, header_map = infer_header_map(sheet, anchor_row, anchor_col)

    y_col = anchor_col - 7
    x_col = anchor_col - 11
    col_num_q = fallback_column(
        pick_column(
            header_map,
            [
                ("num", "quarter"),
                ("quarter", "used"),
                ("n", "quarter"),
            ],
        ),
        anchor_col,
        -12,
    )
    col_forecast = fallback_column(
        pick_column(
            header_map,
            [
                ("tot", "fcst", "w/o", "sa"),
                ("tot", "fcst"),
                ("forecast", "total"),
                ("forecast", "value"),
            ],
        ),
        anchor_col,
        -1,
    )
    col_actual = fallback_column(
        pick_column(header_map, [("actual", "sales"), ("actual", "value"), ("reported",)]),
        anchor_col,
        -2,
    )
    col_max = fallback_column(
        pick_column(header_map, [("max",)]), anchor_col, 0, min_col=anchor_col
    )
    col_min = fallback_column(
        pick_column(header_map, [("min",)]), anchor_col, 1
    )

    key_cols = [col_num_q, col_forecast, col_max, col_min]
    data_rows = detect_data_rows(
        sheet=sheet,
        start_row=header_row + 1,
        stop_row=header_row + 160,
        key_cols=key_cols,
    )

    history = get_xy_history(sheet, anchor_row=anchor_row, x_col=x_col, y_col=y_col)
    if not history:
        return []

    if data_rows:
        data_rows = data_rows[:10]
    else:
        data_rows = [header_row + offset for offset in range(1, min(10, len(history)) + 1)]

    if not data_rows:
        return []

    helper_start_row = anchor_row + 260
    helper_col = max(anchor_col + 2, col_max + 3, col_min + 3)
    formula_refs: List[Tuple[int, int, xw.Range, xw.Range]] = []

    for idx, row in enumerate(data_rows):
        n_quarters = coerce_int(get_cell_value(sheet, row, col_num_q)) or (idx + 1)
        n_quarters = max(2, min(10, n_quarters, len(history)))
        start_row = history[-n_quarters][0]
        end_row = history[-1][0]

        intercept_formula = (
            f"=INTERCEPT(R{start_row}C{y_col}:R{end_row}C{y_col},"
            f"R{start_row}C{x_col}:R{end_row}C{x_col})"
        )
        slope_formula = (
            f"=SLOPE(R{start_row}C{y_col}:R{end_row}C{y_col},"
            f"R{start_row}C{x_col}:R{end_row}C{x_col})"
        )

        intercept_cell = sheet.range((helper_start_row + idx, helper_col))
        slope_cell = sheet.range((helper_start_row + idx, helper_col + 1))
        set_formula2(intercept_cell, intercept_formula)
        set_formula2(slope_cell, slope_formula)
        formula_refs.append((row, n_quarters, intercept_cell, slope_cell))

    if formula_refs:
        wb.app.calculate()

    rows: List[Dict[str, Any]] = []
    previous_signature: Optional[Tuple[Any, ...]] = None
    next_x = history[-1][1] + 1

    for row, n_quarters, intercept_cell, slope_cell in formula_refs:
        intercept_value = coerce_float(intercept_cell.value)
        slope_value = coerce_float(slope_cell.value)
        forecast_value = coerce_float(get_cell_value(sheet, row, col_forecast))
        if forecast_value is None and intercept_value is not None and slope_value is not None:
            forecast_value = intercept_value + (slope_value * next_x)

        forecast_max = coerce_float(get_cell_value(sheet, row, col_max))
        forecast_min = coerce_float(get_cell_value(sheet, row, col_min))
        actual_value = coerce_float(get_cell_value(sheet, row, col_actual))
        range_width = (
            forecast_max - forecast_min
            if forecast_max is not None and forecast_min is not None
            else None
        )

        signature = (
            n_quarters,
            round(intercept_value, 10) if intercept_value is not None else None,
            round(slope_value, 10) if slope_value is not None else None,
            round(forecast_value, 10) if forecast_value is not None else None,
            round(forecast_max, 10) if forecast_max is not None else None,
            round(forecast_min, 10) if forecast_min is not None else None,
        )
        if previous_signature is not None and signature == previous_signature:
            continue
        previous_signature = signature

        if all(
            value is None
            for value in (
                forecast_value,
                forecast_max,
                forecast_min,
                intercept_value,
                slope_value,
            )
        ):
            continue

        rows.append(
            {
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
                "intercept": intercept_value,
                "slope": slope_value,
                "source_file": source_file,
            }
        )

    return rows


def autosize_columns(ws) -> None:
    for col_idx in range(1, ws.max_column + 1):
        header = ws.cell(row=1, column=col_idx).value
        max_len = len(str(header)) if header else 0
        for row_idx in range(2, ws.max_row + 1):
            value = ws.cell(row=row_idx, column=col_idx).value
            if value is None:
                continue
            max_len = max(max_len, len(str(value)))
        ws.column_dimensions[get_column_letter(col_idx)].width = min(max(max_len + 2, 12), 42)


def write_sheet(ws, columns: Sequence[str], rows: Sequence[Dict[str, Any]]) -> None:
    ws.append(list(columns))
    for entry in rows:
        ws.append([clean_output_value(entry.get(column)) for column in columns])
    for cell in ws[1]:
        cell.font = Font(bold=True)
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    autosize_columns(ws)


def save_output_workbook(
    output_path: Path,
    empirical_rows: Sequence[Dict[str, Any]],
    regression_rows: Sequence[Dict[str, Any]],
) -> None:
    workbook = Workbook()
    workbook.remove(workbook.active)
    empirical_sheet = workbook.create_sheet("empirical_candidates")
    regression_sheet = workbook.create_sheet("regression_candidates")
    write_sheet(empirical_sheet, EMPIRICAL_COLUMNS, empirical_rows)
    write_sheet(regression_sheet, REGRESSION_COLUMNS, regression_rows)
    workbook.save(output_path)


def process_workbooks() -> None:
    if not input_dir.exists():
        raise FileNotFoundError(f"Input folder does not exist: {input_dir}")

    source_entries = sorted(input_dir.iterdir(), key=lambda path: path.name.lower())
    output_path = get_unique_output_path(input_dir, output_dir)

    processed_files = 0
    empirical_rows: List[Dict[str, Any]] = []
    regression_rows: List[Dict[str, Any]] = []

    app = xw.App(visible=False, add_book=False)
    try:
        try:
            app.display_alerts = False
            app.screen_updating = False
        except Exception:
            pass

        for entry in source_entries:
            if not entry.is_file():
                continue
            if entry.name.startswith("~"):
                print(f"Skipped file {entry.name}: temporary file")
                continue
            if entry.suffix.lower() != ".xlsx":
                print(f"Skipped file {entry.name}: not an .xlsx file")
                continue

            print(f"Processing file: {entry.name}")
            wb: Optional[xw.Book] = None
            try:
                wb = app.books.open(str(entry), update_links=False)
                metadata = parse_file_metadata(entry.name)
                empirical_rows.extend(
                    extract_empirical_rows(
                        wb=wb,
                        metadata=metadata,
                        source_file=entry.name,
                    )
                )
                regression_rows.extend(
                    extract_regression_rows(
                        wb=wb,
                        metadata=metadata,
                        source_file=entry.name,
                    )
                )
                processed_files += 1
            except Exception as exc:
                print(f"Skipped file {entry.name}: {exc}")
            finally:
                if wb is not None:
                    safe_close_source_workbook(wb)
    finally:
        app.quit()

    save_output_workbook(
        output_path=output_path,
        empirical_rows=empirical_rows,
        regression_rows=regression_rows,
    )

    print(f"Output path: {output_path}")
    print(f"Number of files processed: {processed_files}")
    print(f"Number of empirical rows: {len(empirical_rows)}")
    print(f"Number of regression rows: {len(regression_rows)}")


if __name__ == "__main__":
    process_workbooks()
