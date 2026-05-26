from __future__ import annotations

import re
from dataclasses import dataclass
from datetime import date
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Sequence, Tuple

import xlwings as xw
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter


# Configure these two paths before running.
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


EMPIRICAL_HEADER_SYNONYMS = {
    "num_quarters_used": ["numquartersused", "quartersused", "numquarters", "nquarters", "quarters"],
    "last_quarter_used": ["lastquarterused", "lastquarter", "quarterused"],
    "forecast_value": [
        "estimatedtotalsold",
        "totfcstwosa",
        "totfcst",
        "forecastvalue",
        "forecast",
        "estimatedsales",
    ],
    "actual_value": ["reportedsales", "actualsales", "actualvalue", "salesreported", "actual"],
    "forecast_max": ["max", "forecastmax"],
    "forecast_min": ["min", "forecastmin"],
    "avg_penetration_pct": ["avgpenetrationpct", "averagepenetration", "avgpenetration", "penetrationpct"],
    "quarterly_sales": ["quarterlysales", "salesquarterly", "dbsales", "totalsold"],
    "reported_sales": ["reportedsales", "reported", "salesreported", "actualsales"],
    "growth_rate_pct": ["growthratepct", "growthrate", "growthpct", "growth"],
    "sales_captured_in_db_pct": [
        "salescapturedindbpct",
        "capturedindbpct",
        "salescaptured",
        "capturedpct",
    ],
}

REGRESSION_HEADER_SYNONYMS = {
    "num_quarters_used": ["numquartersused", "quartersused", "numquarters", "nquarters", "quarters"],
    "forecast_value": [
        "totfcstwosa",
        "totfcstwosa",
        "forecastvalue",
        "forecast",
        "totfcst",
        "forecasttotalwithoutsa",
    ],
    "actual_value": ["reportedsales", "actualsales", "actualvalue", "actual"],
    "forecast_max": ["max", "forecastmax"],
    "forecast_min": ["min", "forecastmin"],
    "intercept": ["intercept"],
    "slope": ["slope"],
}


EMPIRICAL_DEFAULT_OFFSETS = {
    "num_quarters_used": -7,
    "last_quarter_used": -6,
    "forecast_value": -2,
    "actual_value": -1,
    "forecast_max": 0,
    "forecast_min": 1,
    "avg_penetration_pct": -5,
    "quarterly_sales": -10,
    "reported_sales": -9,
    "growth_rate_pct": -8,
    "sales_captured_in_db_pct": -4,
}

REGRESSION_DEFAULT_OFFSETS = {
    "num_quarters_used": -6,
    "forecast_value": -1,
    "actual_value": -2,
    "forecast_max": 0,
    "forecast_min": 1,
    "intercept": -4,
    "slope": -3,
}

DAY_BY_PERIOD = {"Early": 5, "Mid": 15, "Late": 25}
MONTH_INDEX = {
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
PERIOD_RE = re.compile(r"(Early|Mid|Late)([A-Za-z]{3,9})(\d{4})", re.IGNORECASE)


@dataclass(frozen=True)
class ModelMetadata:
    ticker: str
    model_period: str
    model_date: str
    model: str


def normalize_header(value: Any) -> str:
    if value is None:
        return ""
    text = str(value).strip().lower()
    if not text:
        return ""
    text = text.replace("%", "pct").replace("&", "and")
    return re.sub(r"[^a-z0-9]+", "", text)


def ensure_2d(values: Any) -> List[List[Any]]:
    if values is None:
        return []
    if isinstance(values, (list, tuple)):
        if not values:
            return []
        first = values[0]
        if isinstance(first, (list, tuple)):
            return [list(row) if isinstance(row, (list, tuple)) else [row] for row in values]
        return [list(values)]
    return [[values]]


def to_float(value: Any) -> Optional[float]:
    if value is None or value == "":
        return None
    if isinstance(value, (int, float)):
        return float(value)
    if isinstance(value, str):
        cleaned = value.strip().replace(",", "")
        if cleaned.endswith("%"):
            cleaned = cleaned[:-1]
        if not cleaned:
            return None
        try:
            return float(cleaned)
        except ValueError:
            return None
    return None


def to_int(value: Any) -> Optional[int]:
    numeric = to_float(value)
    if numeric is None:
        return None
    try:
        return int(round(numeric))
    except (TypeError, ValueError):
        return None


def safe_subtract(left: Any, right: Any) -> Optional[float]:
    left_num = to_float(left)
    right_num = to_float(right)
    if left_num is None or right_num is None:
        return None
    return left_num - right_num


def safe_close_workbook(wb: xw.Book) -> None:
    try:
        wb.close(save=False)
        return
    except TypeError:
        pass
    except Exception:
        pass

    try:
        wb.close(SaveChanges=False)
        return
    except Exception:
        pass

    try:
        wb.api.Close(SaveChanges=False)
    except Exception:
        pass


def find_sheet(wb: xw.Book, name: str) -> Optional[xw.Sheet]:
    try:
        return wb.sheets[name]
    except Exception:
        lowered = name.strip().lower()
        for sheet in wb.sheets:
            if sheet.name.strip().lower() == lowered:
                return sheet
    return None


def find_anchor_cell(sheet: xw.Sheet, target: str = "max") -> Optional[Tuple[int, int]]:
    used = sheet.used_range
    values = ensure_2d(used.value)
    if not values:
        return None

    start_row, start_col = used.row, used.column
    target_norm = normalize_header(target)
    best_match: Optional[Tuple[int, int, int]] = None

    for r_idx, row_values in enumerate(values):
        for c_idx, value in enumerate(row_values):
            if normalize_header(value) != target_norm:
                continue

            score = 1
            right_cell = row_values[c_idx + 1] if c_idx + 1 < len(row_values) else None
            if normalize_header(right_cell) == "min":
                score += 2
            left_cell = row_values[c_idx - 1] if c_idx > 0 else None
            if isinstance(left_cell, str) and left_cell.strip():
                score += 1

            candidate = (score, start_row + r_idx, start_col + c_idx)
            if best_match is None or candidate[0] > best_match[0]:
                best_match = candidate

    if best_match is None:
        return None
    return best_match[1], best_match[2]


def collect_header_offsets(sheet: xw.Sheet, anchor_row: int, anchor_col: int, window: int = 40) -> Dict[str, int]:
    start_col = max(1, anchor_col - window)
    end_col = anchor_col + window
    row_deltas = (0, -1, 1)
    weighted_offsets: Dict[str, Tuple[int, Tuple[int, int]]] = {}

    for row_delta in row_deltas:
        row = anchor_row + row_delta
        if row < 1:
            continue

        values = sheet.range((row, start_col), (row, end_col)).value
        row_values = values if isinstance(values, list) else [values]
        priority = 3 if row_delta == 0 else 2 if row_delta == -1 else 1

        for idx, raw in enumerate(row_values):
            key = normalize_header(raw)
            if not key:
                continue
            col = start_col + idx
            offset = col - anchor_col
            score = (priority, -abs(offset))
            existing = weighted_offsets.get(key)
            if existing is None or score > existing[1]:
                weighted_offsets[key] = (offset, score)

    return {key: payload[0] for key, payload in weighted_offsets.items()}


def resolve_column(
    anchor_col: int,
    header_offsets: Dict[str, int],
    synonyms: Sequence[str],
    fallback_offset: Optional[int] = None,
) -> Optional[int]:
    for synonym in synonyms:
        if synonym in header_offsets:
            return anchor_col + header_offsets[synonym]
    if fallback_offset is not None:
        return anchor_col + fallback_offset
    return None


def build_columns(
    anchor_col: int,
    header_offsets: Dict[str, int],
    synonym_map: Dict[str, Sequence[str]],
    fallback_offsets: Dict[str, int],
) -> Dict[str, Optional[int]]:
    resolved: Dict[str, Optional[int]] = {}
    for field, synonyms in synonym_map.items():
        resolved[field] = resolve_column(
            anchor_col=anchor_col,
            header_offsets=header_offsets,
            synonyms=synonyms,
            fallback_offset=fallback_offsets.get(field),
        )
    return resolved


def get_block_value(block: List[List[Any]], row_index: int, col: Optional[int], min_col: int) -> Any:
    if col is None:
        return None
    if row_index < 0 or row_index >= len(block):
        return None
    col_index = col - min_col
    row = block[row_index]
    if col_index < 0 or col_index >= len(row):
        return None
    return row[col_index]


def parse_model_metadata(file_name: str) -> ModelMetadata:
    stem = Path(file_name).stem
    parts = [part.strip() for part in stem.split(" - ")]
    ticker = parts[1] if len(parts) > 1 and parts[1] else ""

    if not ticker:
        ticker_match = re.search(r"\b[A-Z]{1,6}\b", stem)
        ticker = ticker_match.group(0) if ticker_match else "UNKNOWN"

    period_chunk = parts[2] if len(parts) > 2 else stem
    period_chunk = period_chunk.split("_")[0].strip()
    period_match = PERIOD_RE.search(period_chunk)

    model_period = "unknown_period"
    model_date = ""

    if period_match:
        period_label = period_match.group(1).capitalize()
        month_token = period_match.group(2)[:3].lower()
        year = int(period_match.group(3))

        month = MONTH_INDEX.get(month_token)
        day = DAY_BY_PERIOD.get(period_label)
        if month is not None and day is not None:
            month_abbrev = period_match.group(2)[:3].title()
            model_period = f"{period_label}{month_abbrev}_{year}"
            model_date = date(year, month, day).isoformat()

    model = f"{ticker}_{model_period}"
    return ModelMetadata(ticker=ticker, model_period=model_period, model_date=model_date, model=model)


def is_row_effectively_empty(values: Sequence[Any]) -> bool:
    for value in values:
        if value is None:
            continue
        if isinstance(value, str) and not value.strip():
            continue
        return False
    return True


def values_match(a: Any, b: Any, tolerance: float = 1e-9) -> bool:
    a_num = to_float(a)
    b_num = to_float(b)
    if a_num is not None and b_num is not None:
        return abs(a_num - b_num) <= tolerance
    return a == b


def rows_look_duplicate(previous: Dict[str, Any], current: Dict[str, Any]) -> bool:
    keys = ("forecast_value", "forecast_max", "forecast_min", "intercept", "slope")
    return all(values_match(previous.get(key), current.get(key)) for key in keys)


def extract_empirical_rows(
    wb: xw.Book,
    sheet: xw.Sheet,
    metadata: ModelMetadata,
    source_file: str,
) -> List[Dict[str, Any]]:
    anchor = find_anchor_cell(sheet, "max")
    if anchor is None:
        print(f"Skipped empirical extraction for {source_file}: could not find 'max' anchor.")
        return []

    anchor_row, anchor_col = anchor
    header_offsets = collect_header_offsets(sheet, anchor_row, anchor_col)
    columns = build_columns(
        anchor_col=anchor_col,
        header_offsets=header_offsets,
        synonym_map=EMPIRICAL_HEADER_SYNONYMS,
        fallback_offsets=EMPIRICAL_DEFAULT_OFFSETS,
    )

    start_row = anchor_row + 1
    end_row = start_row + 9

    # Temporary formula writes are allowed for calculation refreshes.
    avg_col = columns.get("avg_penetration_pct")
    quarterly_col = columns.get("quarterly_sales")
    reported_col = columns.get("reported_sales")
    quarters_col = columns.get("num_quarters_used")
    formula_updates = False

    if avg_col and quarterly_col and reported_col:
        avg_cell_offset_q = quarterly_col - avg_col
        avg_cell_offset_r = reported_col - avg_col
        for row in range(start_row, end_row + 1):
            avg_value = sheet.range((row, avg_col)).value
            if avg_value not in (None, ""):
                continue

            n_quarters = to_int(sheet.range((row, quarters_col)).value) if quarters_col else None
            n_quarters = n_quarters if n_quarters and n_quarters > 0 else (row - start_row + 1)
            if n_quarters <= 1:
                formula = f'=IFERROR(RC[{avg_cell_offset_q}]/RC[{avg_cell_offset_r}], "")'
            else:
                formula = (
                    f'=IFERROR(SUM(R[-{n_quarters - 1}]C[{avg_cell_offset_q}]:RC[{avg_cell_offset_q}])'
                    f'/SUM(R[-{n_quarters - 1}]C[{avg_cell_offset_r}]:RC[{avg_cell_offset_r}]), "")'
                )
            sheet.range((row, avg_col)).formula2 = formula
            formula_updates = True

    if formula_updates:
        wb.app.calculate()

    used_cols = [col for col in columns.values() if col is not None]
    if not used_cols:
        return []

    min_col, max_col = min(used_cols), max(used_cols)
    block = ensure_2d(sheet.range((start_row, min_col), (end_row, max_col)).value)
    if not block:
        return []

    results: List[Dict[str, Any]] = []
    for idx in range(10):
        row_vals = block[idx] if idx < len(block) else []
        if is_row_effectively_empty(row_vals):
            continue

        num_quarters = to_int(get_block_value(block, idx, columns.get("num_quarters_used"), min_col)) or (idx + 1)
        last_quarter_used = get_block_value(block, idx, columns.get("last_quarter_used"), min_col)
        forecast_value = get_block_value(block, idx, columns.get("forecast_value"), min_col)
        actual_value = get_block_value(block, idx, columns.get("actual_value"), min_col)
        forecast_max = get_block_value(block, idx, columns.get("forecast_max"), min_col)
        forecast_min = get_block_value(block, idx, columns.get("forecast_min"), min_col)
        avg_penetration = get_block_value(block, idx, columns.get("avg_penetration_pct"), min_col)
        quarterly_sales = get_block_value(block, idx, columns.get("quarterly_sales"), min_col)
        reported_sales = get_block_value(block, idx, columns.get("reported_sales"), min_col)
        growth_rate = get_block_value(block, idx, columns.get("growth_rate_pct"), min_col)
        sales_captured = get_block_value(block, idx, columns.get("sales_captured_in_db_pct"), min_col)

        results.append(
            {
                "model": metadata.model,
                "ticker": metadata.ticker,
                "model_period": metadata.model_period,
                "model_date": metadata.model_date,
                "method": "empirical",
                "parameter_name": "avg_penetration_pct",
                "parameter_value": avg_penetration,
                "num_quarters_used": num_quarters,
                "last_quarter_used": last_quarter_used,
                "forecast_value": forecast_value,
                "actual_value": actual_value,
                "forecast_max": forecast_max,
                "forecast_min": forecast_min,
                "range_width": safe_subtract(forecast_max, forecast_min),
                "avg_penetration_pct": avg_penetration,
                "quarterly_sales": quarterly_sales,
                "reported_sales": reported_sales,
                "growth_rate_pct": growth_rate,
                "sales_captured_in_db_pct": sales_captured,
                "source_file": source_file,
            }
        )

    return results


def extract_regression_rows(
    wb: xw.Book,
    sheet: xw.Sheet,
    metadata: ModelMetadata,
    source_file: str,
) -> List[Dict[str, Any]]:
    anchor = find_anchor_cell(sheet, "max")
    if anchor is None:
        print(f"Skipped regression extraction for {source_file}: could not find 'max' anchor.")
        return []

    anchor_row, anchor_col = anchor
    y_col = anchor_col - 7
    x_col = anchor_col - 11

    header_offsets = collect_header_offsets(sheet, anchor_row, anchor_col)
    columns = build_columns(
        anchor_col=anchor_col,
        header_offsets=header_offsets,
        synonym_map=REGRESSION_HEADER_SYNONYMS,
        fallback_offsets=REGRESSION_DEFAULT_OFFSETS,
    )

    start_row = anchor_row + 1
    end_row = start_row + 9
    history_end_row = anchor_row - 1
    history_start_floor = max(1, history_end_row - 80)

    num_quarters_col = columns.get("num_quarters_used")
    intercept_col = columns.get("intercept")
    slope_col = columns.get("slope")

    formula_updates = False
    if history_end_row > history_start_floor and intercept_col and slope_col:
        for row in range(start_row, end_row + 1):
            n_quarters = to_int(sheet.range((row, num_quarters_col)).value) if num_quarters_col else None
            n_quarters = n_quarters if n_quarters and n_quarters > 0 else (row - start_row + 1)
            history_start = max(history_start_floor, history_end_row - n_quarters + 1)

            intercept_formula = (
                f'=IFERROR(INTERCEPT(R{history_start}C{y_col}:R{history_end_row}C{y_col},'
                f'R{history_start}C{x_col}:R{history_end_row}C{x_col}), "")'
            )
            slope_formula = (
                f'=IFERROR(SLOPE(R{history_start}C{y_col}:R{history_end_row}C{y_col},'
                f'R{history_start}C{x_col}:R{history_end_row}C{x_col}), "")'
            )

            sheet.range((row, intercept_col)).formula2 = intercept_formula
            sheet.range((row, slope_col)).formula2 = slope_formula
            formula_updates = True

    if formula_updates:
        wb.app.calculate()

    used_cols = [col for col in columns.values() if col is not None]
    if not used_cols:
        return []

    min_col, max_col = min(used_cols), max(used_cols)
    block = ensure_2d(sheet.range((start_row, min_col), (end_row, max_col)).value)
    if not block:
        return []

    results: List[Dict[str, Any]] = []
    for idx in range(10):
        row_vals = block[idx] if idx < len(block) else []
        if is_row_effectively_empty(row_vals):
            continue

        num_quarters = to_int(get_block_value(block, idx, columns.get("num_quarters_used"), min_col)) or (idx + 1)
        forecast_value = get_block_value(block, idx, columns.get("forecast_value"), min_col)
        actual_value = get_block_value(block, idx, columns.get("actual_value"), min_col)
        forecast_max = get_block_value(block, idx, columns.get("forecast_max"), min_col)
        forecast_min = get_block_value(block, idx, columns.get("forecast_min"), min_col)
        intercept = get_block_value(block, idx, columns.get("intercept"), min_col)
        slope = get_block_value(block, idx, columns.get("slope"), min_col)

        row = {
            "model": metadata.model,
            "ticker": metadata.ticker,
            "model_period": metadata.model_period,
            "model_date": metadata.model_date,
            "method": "regression",
            "parameter_name": "num_quarters_used",
            "parameter_value": num_quarters,
            "num_quarters_used": num_quarters,
            "forecast_value": forecast_value,
            "actual_value": actual_value if actual_value not in (None, "") else "",
            "forecast_max": forecast_max,
            "forecast_min": forecast_min,
            "range_width": safe_subtract(forecast_max, forecast_min),
            "intercept": intercept,
            "slope": slope,
            "source_file": source_file,
        }

        if idx == 9 and results and rows_look_duplicate(results[-1], row):
            continue
        results.append(row)

    return results


def write_sheet(workbook: Workbook, sheet_name: str, columns: Sequence[str], rows: Sequence[Dict[str, Any]]) -> None:
    ws = workbook.create_sheet(title=sheet_name)
    ws.append(list(columns))
    for row in rows:
        ws.append([row.get(column, "") for column in columns])

    for cell in ws[1]:
        cell.font = Font(bold=True)

    ws.freeze_panes = "A2"
    if ws.max_row >= 1 and ws.max_column >= 1:
        last_col_letter = get_column_letter(ws.max_column)
        ws.auto_filter.ref = f"A1:{last_col_letter}{ws.max_row}"

    for idx, column in enumerate(columns, start=1):
        max_len = len(column)
        for row_idx in range(2, ws.max_row + 1):
            value = ws.cell(row=row_idx, column=idx).value
            if value is None:
                text = ""
            elif isinstance(value, float):
                text = f"{value:.6g}"
            else:
                text = str(value)
            if len(text) > max_len:
                max_len = len(text)
        ws.column_dimensions[get_column_letter(idx)].width = min(max_len + 2, 60)


def write_output_workbook(
    output_path: Path,
    empirical_rows: Sequence[Dict[str, Any]],
    regression_rows: Sequence[Dict[str, Any]],
) -> None:
    workbook = Workbook()
    default_sheet = workbook.active
    workbook.remove(default_sheet)

    write_sheet(workbook, "empirical_candidates", EMPIRICAL_COLUMNS, empirical_rows)
    write_sheet(workbook, "regression_candidates", REGRESSION_COLUMNS, regression_rows)

    workbook.save(output_path)


def choose_output_path(input_folder: Path, output_folder: Path) -> Path:
    input_folder_name = input_folder.resolve().name
    stem = f"{input_folder_name}_PARAM"
    candidate = output_folder / f"{stem}.xlsx"
    if not candidate.exists():
        return candidate

    index = 1
    while True:
        candidate = output_folder / f"{stem}.{index}.xlsx"
        if not candidate.exists():
            return candidate
        index += 1


def iter_input_files(folder: Path) -> Iterable[Path]:
    for file_path in sorted(folder.iterdir(), key=lambda p: p.name.lower()):
        if not file_path.is_file():
            continue
        if file_path.name.startswith("~"):
            print(f"Skipped {file_path.name}: temporary file.")
            continue
        if file_path.suffix.lower() != ".xlsx":
            print(f"Skipped {file_path.name}: not an .xlsx file.")
            continue
        yield file_path


def main() -> None:
    source_dir = input_dir.expanduser().resolve()
    target_dir = output_dir.expanduser().resolve()

    if not source_dir.exists():
        raise FileNotFoundError(f"Input directory does not exist: {source_dir}")

    target_dir.mkdir(parents=True, exist_ok=True)
    output_path = choose_output_path(source_dir, target_dir)

    empirical_rows: List[Dict[str, Any]] = []
    regression_rows: List[Dict[str, Any]] = []
    files_processed = 0

    app = xw.App(visible=False, add_book=False)
    app.display_alerts = False
    try:
        app.screen_updating = False
    except Exception:
        pass

    try:
        for file_path in iter_input_files(source_dir):
            print(f"Processing {file_path.name}")
            wb: Optional[xw.Book] = None
            try:
                wb = app.books.open(str(file_path), update_links=False)
                metadata = parse_model_metadata(file_path.name)

                empirical_sheet = find_sheet(wb, "Empirical Model")
                if empirical_sheet is None:
                    print(f"Skipped empirical extraction for {file_path.name}: missing 'Empirical Model' sheet.")
                else:
                    empirical_rows.extend(
                        extract_empirical_rows(
                            wb=wb,
                            sheet=empirical_sheet,
                            metadata=metadata,
                            source_file=file_path.name,
                        )
                    )

                regression_sheet = find_sheet(wb, "Regression Model")
                if regression_sheet is None:
                    print(f"Skipped regression extraction for {file_path.name}: missing 'Regression Model' sheet.")
                else:
                    regression_rows.extend(
                        extract_regression_rows(
                            wb=wb,
                            sheet=regression_sheet,
                            metadata=metadata,
                            source_file=file_path.name,
                        )
                    )

                files_processed += 1
            except Exception as exc:
                print(f"Skipped {file_path.name}: {exc}")
            finally:
                if wb is not None:
                    safe_close_workbook(wb)
    finally:
        app.quit()

    write_output_workbook(output_path, empirical_rows, regression_rows)

    print(f"Output path: {output_path}")
    print(f"Number of files processed: {files_processed}")
    print(f"Number of empirical rows: {len(empirical_rows)}")
    print(f"Number of regression rows: {len(regression_rows)}")


if __name__ == "__main__":
    main()
