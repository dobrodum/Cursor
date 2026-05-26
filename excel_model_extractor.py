from __future__ import annotations

import re
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Sequence, Tuple

import xlwings as xw
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

# -----------------------------
# User-configurable paths
# -----------------------------
input_dir = Path("input")
output_dir = Path("output")


EMPIRICAL_SHEET_NAME = "Empirical Model"
REGRESSION_SHEET_NAME = "Regression Model"
N_QUARTERS = 10


EMPIRICAL_COLUMNS: List[str] = [
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


REGRESSION_COLUMNS: List[str] = [
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


# Default offsets are relative to the located "max" anchor column.
# Header discovery (if present) overrides these values.
DEFAULT_EMPIRICAL_OFFSETS: Dict[str, int] = {
    "num_quarters_used": -7,
    "last_quarter_used": -6,
    "forecast_value": -1,
    "actual_value": -2,
    "forecast_max": 0,
    "forecast_min": 1,
    "quarterly_sales": -3,
    "reported_sales": -2,
    "growth_rate_pct": -4,
    "sales_captured_in_db_pct": -5,
    "avg_penetration_pct": -5,
}


DEFAULT_REGRESSION_OFFSETS: Dict[str, int] = {
    "num_quarters_used": -6,
    "forecast_value": -1,
    "actual_value": -2,
    "forecast_max": 0,
    "forecast_min": 1,
}


EMPIRICAL_LABEL_PATTERNS: Dict[str, Sequence[str]] = {
    "num_quarters_used": (r"num.*quarter", r"quarter.*used", r"\bn\s*quarter"),
    "last_quarter_used": (r"last.*quarter",),
    "forecast_value": (r"estimated.*total.*sold", r"tot.*fcst", r"forecast"),
    "actual_value": (r"actual", r"reported.*sales"),
    "forecast_max": (r"^max$",),
    "forecast_min": (r"^min$",),
    "avg_penetration_pct": (r"avg.*penetration", r"average.*penetration"),
    "quarterly_sales": (r"quarterly.*sales", r"qtr.*sales"),
    "reported_sales": (r"reported.*sales",),
    "growth_rate_pct": (r"growth.*rate",),
    "sales_captured_in_db_pct": (r"sales.*captured.*db", r"captured.*db", r"penetration"),
}


REGRESSION_LABEL_PATTERNS: Dict[str, Sequence[str]] = {
    "num_quarters_used": (r"num.*quarter", r"quarter.*used", r"\bn\s*quarter"),
    "forecast_value": (r"tot.*fcst.*w.*o.*sa", r"forecast.*w.*o.*sa", r"tot.*fcst"),
    "actual_value": (r"actual", r"reported.*sales"),
    "forecast_max": (r"^max$",),
    "forecast_min": (r"^min$",),
}


def normalize_text(value: Any) -> str:
    if value is None:
        return ""
    text = str(value).strip().lower()
    text = re.sub(r"[^a-z0-9]+", " ", text)
    return re.sub(r"\s+", " ", text).strip()


def normalize_matrix(values: Any, rows: int, cols: int) -> List[List[Any]]:
    if rows <= 0 or cols <= 0:
        return []

    if rows == 1 and cols == 1:
        return [[values]]

    if rows == 1:
        if isinstance(values, list):
            if values and isinstance(values[0], list):
                row = list(values[0])
            else:
                row = list(values)
        else:
            row = [values]
        if len(row) < cols:
            row.extend([None] * (cols - len(row)))
        return [row[:cols]]

    if cols == 1:
        if isinstance(values, list):
            if values and isinstance(values[0], list):
                col_vals = [(entry[0] if entry else None) for entry in values[:rows]]
            else:
                col_vals = list(values[:rows])
        else:
            col_vals = [values]
        if len(col_vals) < rows:
            col_vals.extend([None] * (rows - len(col_vals)))
        return [[entry] for entry in col_vals[:rows]]

    if isinstance(values, list) and values and isinstance(values[0], list):
        matrix: List[List[Any]] = []
        for row_idx in range(rows):
            source = values[row_idx] if row_idx < len(values) and isinstance(values[row_idx], list) else []
            row = list(source[:cols])
            if len(row) < cols:
                row.extend([None] * (cols - len(row)))
            matrix.append(row)
        return matrix

    flat = list(values) if isinstance(values, list) else [values]
    matrix = []
    cursor = 0
    for _ in range(rows):
        row = flat[cursor : cursor + cols]
        cursor += cols
        if len(row) < cols:
            row.extend([None] * (cols - len(row)))
        matrix.append(row)
    return matrix


def as_number(value: Any) -> Optional[float]:
    if value is None:
        return None
    if isinstance(value, (int, float)):
        return float(value)
    text = str(value).strip()
    if not text:
        return None
    if text.endswith("%"):
        try:
            return float(text[:-1].replace(",", "")) / 100.0
        except ValueError:
            return None
    try:
        return float(text.replace(",", ""))
    except ValueError:
        return None


def first_not_none(*values: Any) -> Any:
    for value in values:
        if value is None:
            continue
        if isinstance(value, str) and not value.strip():
            continue
        return value
    return None


def subtract_or_none(a: Optional[float], b: Optional[float]) -> Optional[float]:
    if a is None or b is None:
        return None
    return a - b


def month_to_number(month_text: str) -> int:
    for fmt in ("%b", "%B"):
        try:
            return datetime.strptime(month_text, fmt).month
        except ValueError:
            continue
    raise ValueError(f"Could not parse month token '{month_text}'.")


def parse_file_metadata(file_name: str) -> Dict[str, str]:
    stem = Path(file_name).stem
    parts = [part.strip() for part in stem.split(" - ")]

    ticker = ""
    if len(parts) >= 2:
        ticker = re.sub(r"[^A-Za-z0-9]", "", parts[1]).upper()
    if not ticker:
        ticker_match = re.search(r"\b[A-Z]{2,6}\b", stem)
        ticker = ticker_match.group(0) if ticker_match else "UNKNOWN"

    period_chunk = parts[2] if len(parts) >= 3 else stem
    period_token = re.split(r"[_\s-]+", period_chunk)[0]

    model_period = "UNKNOWN_PERIOD"
    model_date = ""
    period_match = re.match(r"(?i)^(early|mid|late)([A-Za-z]{3,9})(\d{4})$", period_token)
    if period_match:
        bucket = period_match.group(1).capitalize()
        month_text = period_match.group(2)
        year = int(period_match.group(3))
        month_num = month_to_number(month_text)
        month_abbrev = datetime(year, month_num, 1).strftime("%b")
        day_map = {"Early": 5, "Mid": 15, "Late": 25}
        model_period = f"{bucket}{month_abbrev}_{year}"
        model_date = f"{year:04d}-{month_num:02d}-{day_map[bucket]:02d}"
    else:
        model_period = re.sub(r"[^A-Za-z0-9]+", "_", period_token).strip("_") or "UNKNOWN_PERIOD"

    model = f"{ticker}_{model_period}"
    return {
        "model": model,
        "ticker": ticker,
        "model_period": model_period,
        "model_date": model_date,
    }


def next_output_path(source_input_dir: Path, target_output_dir: Path) -> Path:
    target_output_dir.mkdir(parents=True, exist_ok=True)
    base_name = f"{source_input_dir.name}_PARAM"
    candidate = target_output_dir / f"{base_name}.xlsx"
    suffix = 1
    while candidate.exists():
        candidate = target_output_dir / f"{base_name}.{suffix}.xlsx"
        suffix += 1
    return candidate


def set_formula2(cell: xw.main.Range, formula_r1c1: str) -> None:
    try:
        cell.formula2 = formula_r1c1
    except Exception:
        cell.formula = formula_r1c1


def safe_close_workbook(wb: xw.main.Book) -> None:
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
        wb.close()
    except Exception:
        pass


def find_max_anchor(sheet: xw.main.Sheet) -> Tuple[int, int]:
    used = sheet.used_range
    row_count = used.rows.count
    col_count = used.columns.count
    matrix = normalize_matrix(used.value, row_count, col_count)
    start_row = used.row
    start_col = used.column

    candidates: List[Tuple[int, int, int]] = []
    for r_idx, row_vals in enumerate(matrix):
        for c_idx, cell_value in enumerate(row_vals):
            if normalize_text(cell_value) != "max":
                continue
            score = 1
            for d_row, d_col in ((0, 1), (1, 0), (0, -1), (-1, 0)):
                rr = r_idx + d_row
                cc = c_idx + d_col
                if 0 <= rr < len(matrix) and 0 <= cc < len(matrix[rr]):
                    if normalize_text(matrix[rr][cc]) == "min":
                        score += 4
            candidates.append((score, start_row + r_idx, start_col + c_idx))

    if not candidates:
        raise ValueError(f'Could not locate "max" anchor in sheet "{sheet.name}".')

    candidates.sort(key=lambda item: (-item[0], item[1], item[2]))
    _, anchor_row, anchor_col = candidates[0]
    return anchor_row, anchor_col


def discover_header_row(
    sheet: xw.main.Sheet,
    anchor_row: int,
    anchor_col: int,
    patterns: Dict[str, Sequence[str]],
) -> Tuple[int, Dict[str, int]]:
    min_row = max(1, anchor_row - 3)
    max_row = anchor_row + 3
    min_col = max(1, anchor_col - 20)
    max_col = anchor_col + 20

    row_count = max_row - min_row + 1
    col_count = max_col - min_col + 1
    block = normalize_matrix(
        sheet.range((min_row, min_col), (max_row, max_col)).value,
        row_count,
        col_count,
    )
    best_row = anchor_row
    best_score = float("-inf")
    best_mapping: Dict[str, int] = {}

    for row_offset, row_vals in enumerate(block):
        row_number = min_row + row_offset
        mapping: Dict[str, int] = {}

        for col_offset, raw_cell in enumerate(row_vals):
            label = normalize_text(raw_cell)
            if not label:
                continue
            col_number = min_col + col_offset
            for logical_name, regexes in patterns.items():
                if logical_name in mapping:
                    continue
                if any(re.search(regex, label) for regex in regexes):
                    mapping[logical_name] = col_number
                    break

        score = float(len(mapping) * 2)
        if "forecast_max" in mapping:
            score += 1.0
        if "forecast_min" in mapping:
            score += 1.0
        score -= 0.2 * abs(row_number - anchor_row)

        if score > best_score:
            best_score = score
            best_row = row_number
            best_mapping = mapping

    return best_row, best_mapping


def resolve_columns(
    anchor_col: int,
    header_map: Dict[str, int],
    default_offsets: Dict[str, int],
) -> Dict[str, int]:
    resolved: Dict[str, int] = {}
    for name, offset in default_offsets.items():
        col = header_map.get(name, anchor_col + offset)
        resolved[name] = max(1, int(col))
    return resolved


def read_block_values(
    sheet: xw.main.Sheet,
    start_row: int,
    rows: int,
    columns: Iterable[int],
) -> Tuple[List[List[Any]], int, int]:
    cols = [max(1, int(col)) for col in columns]
    min_col = min(cols)
    max_col = max(cols)
    end_row = start_row + rows - 1
    col_count = max_col - min_col + 1
    matrix = normalize_matrix(
        sheet.range((start_row, min_col), (end_row, max_col)).value,
        rows,
        col_count,
    )
    return matrix, min_col, max_col


def block_get(matrix: List[List[Any]], row_idx: int, col: int, min_col: int) -> Any:
    if row_idx < 0 or row_idx >= len(matrix):
        return None
    col_idx = col - min_col
    row_vals = matrix[row_idx]
    if col_idx < 0 or col_idx >= len(row_vals):
        return None
    return row_vals[col_idx]


def extract_empirical_rows(
    wb: xw.main.Book,
    sheet: xw.main.Sheet,
    metadata: Dict[str, str],
    source_file: str,
) -> List[Dict[str, Any]]:
    anchor_row, anchor_col = find_max_anchor(sheet)
    header_row, header_map = discover_header_row(sheet, anchor_row, anchor_col, EMPIRICAL_LABEL_PATTERNS)
    column_map = resolve_columns(anchor_col, header_map, DEFAULT_EMPIRICAL_OFFSETS)
    data_start_row = header_row + 1

    block_columns = list(column_map.values())
    matrix, min_col, _ = read_block_values(sheet, data_start_row, N_QUARTERS, block_columns)

    helper_row = max(sheet.used_range.last_cell.row + 2, data_start_row + N_QUARTERS + 2)
    helper_col = max(anchor_col + 12, max(block_columns) + 2)
    avg_cell = sheet.range((helper_row, helper_col))
    forecast_cell = sheet.range((helper_row, helper_col + 1))

    rows: List[Dict[str, Any]] = []
    for idx in range(N_QUARTERS):
        row_num = data_start_row + idx
        num_quarters_value = as_number(block_get(matrix, idx, column_map["num_quarters_used"], min_col))
        num_quarters_used = int(num_quarters_value) if num_quarters_value is not None else (idx + 1)

        sales_captured_col = column_map["sales_captured_in_db_pct"]
        quarterly_sales_col = column_map["quarterly_sales"]
        start_hist_row = max(data_start_row, row_num - num_quarters_used + 1)

        set_formula2(
            avg_cell,
            f"=AVERAGE(R{start_hist_row}C{sales_captured_col}:R{row_num}C{sales_captured_col})",
        )
        set_formula2(
            forecast_cell,
            f'=IF(R{helper_row}C{helper_col}=0,"",R{row_num}C{quarterly_sales_col}/R{helper_row}C{helper_col})',
        )
        wb.app.calculate()

        avg_penetration_pct = as_number(avg_cell.value)
        forecast_value = first_not_none(
            as_number(block_get(matrix, idx, column_map["forecast_value"], min_col)),
            as_number(forecast_cell.value),
        )

        reported_sales = first_not_none(
            as_number(block_get(matrix, idx, column_map["reported_sales"], min_col)),
            as_number(block_get(matrix, idx, column_map["actual_value"], min_col)),
        )
        quarterly_sales = as_number(block_get(matrix, idx, column_map["quarterly_sales"], min_col))
        forecast_max = as_number(block_get(matrix, idx, column_map["forecast_max"], min_col))
        forecast_min = as_number(block_get(matrix, idx, column_map["forecast_min"], min_col))
        growth_rate_pct = as_number(block_get(matrix, idx, column_map["growth_rate_pct"], min_col))
        sales_captured_in_db_pct = as_number(
            block_get(matrix, idx, column_map["sales_captured_in_db_pct"], min_col)
        )
        last_quarter_used = block_get(matrix, idx, column_map["last_quarter_used"], min_col)

        # Skip empty candidate rows quickly.
        if all(
            value is None
            for value in (
                forecast_value,
                forecast_max,
                forecast_min,
                quarterly_sales,
                reported_sales,
                growth_rate_pct,
                sales_captured_in_db_pct,
            )
        ):
            continue

        row = {
            "model": metadata["model"],
            "ticker": metadata["ticker"],
            "model_period": metadata["model_period"],
            "model_date": metadata["model_date"],
            "method": "empirical",
            "parameter_name": "avg_penetration_pct",
            "parameter_value": avg_penetration_pct,
            "num_quarters_used": num_quarters_used,
            "last_quarter_used": last_quarter_used,
            "forecast_value": forecast_value,
            "actual_value": reported_sales,
            "forecast_max": forecast_max,
            "forecast_min": forecast_min,
            "range_width": subtract_or_none(forecast_max, forecast_min),
            "avg_penetration_pct": avg_penetration_pct,
            "quarterly_sales": quarterly_sales,
            "reported_sales": reported_sales,
            "growth_rate_pct": growth_rate_pct,
            "sales_captured_in_db_pct": sales_captured_in_db_pct,
            "source_file": source_file,
        }
        rows.append(row)

    sheet.range((helper_row, helper_col), (helper_row, helper_col + 1)).clear_contents()
    return rows


def extract_regression_rows(
    wb: xw.main.Book,
    sheet: xw.main.Sheet,
    metadata: Dict[str, str],
    source_file: str,
) -> List[Dict[str, Any]]:
    anchor_row, anchor_col = find_max_anchor(sheet)
    header_row, header_map = discover_header_row(sheet, anchor_row, anchor_col, REGRESSION_LABEL_PATTERNS)
    column_map = resolve_columns(anchor_col, header_map, DEFAULT_REGRESSION_OFFSETS)

    # Required by spec.
    y_col = anchor_col - 7
    x_col = anchor_col - 11
    if x_col < 1 or y_col < 1 or anchor_row <= 1:
        return []

    x_values = normalize_matrix(sheet.range((1, x_col), (anchor_row - 1, x_col)).value, anchor_row - 1, 1)
    y_values = normalize_matrix(sheet.range((1, y_col), (anchor_row - 1, y_col)).value, anchor_row - 1, 1)
    points: List[Tuple[int, float, float]] = []
    for row_idx in range(anchor_row - 1):
        x_val = as_number(x_values[row_idx][0] if row_idx < len(x_values) else None)
        y_val = as_number(y_values[row_idx][0] if row_idx < len(y_values) else None)
        if x_val is None or y_val is None:
            continue
        points.append((row_idx + 1, x_val, y_val))

    if len(points) < 2:
        return []

    data_start_row = header_row + 1
    table_block, table_min_col, _ = read_block_values(
        sheet,
        data_start_row,
        N_QUARTERS,
        column_map.values(),
    )

    helper_row = max(sheet.used_range.last_cell.row + 2, data_start_row + N_QUARTERS + 2)
    helper_col = max(anchor_col + 12, max(column_map.values()) + 2)
    intercept_cell = sheet.range((helper_row, helper_col))
    slope_cell = sheet.range((helper_row, helper_col + 1))
    forecast_cell = sheet.range((helper_row, helper_col + 2))
    max_cell = sheet.range((helper_row, helper_col + 3))
    min_cell = sheet.range((helper_row, helper_col + 4))

    rows: List[Dict[str, Any]] = []
    previous_signature: Optional[Tuple[Any, ...]] = None

    max_n = min(N_QUARTERS, len(points))
    for n_quarters_used in range(2, max_n + 1):
        start_row = points[-n_quarters_used][0]
        end_row = points[-1][0]

        set_formula2(
            intercept_cell,
            f"=INTERCEPT(R{start_row}C{y_col}:R{end_row}C{y_col},R{start_row}C{x_col}:R{end_row}C{x_col})",
        )
        set_formula2(
            slope_cell,
            f"=SLOPE(R{start_row}C{y_col}:R{end_row}C{y_col},R{start_row}C{x_col}:R{end_row}C{x_col})",
        )
        set_formula2(
            forecast_cell,
            f"=R{helper_row}C{helper_col}+R{helper_row}C{helper_col + 1}*R{end_row}C{x_col}",
        )
        set_formula2(max_cell, f"=MAX(R{start_row}C{y_col}:R{end_row}C{y_col})")
        set_formula2(min_cell, f"=MIN(R{start_row}C{y_col}:R{end_row}C{y_col})")
        wb.app.calculate()

        table_idx = min(max(n_quarters_used - 1, 0), N_QUARTERS - 1)
        table_forecast = as_number(block_get(table_block, table_idx, column_map["forecast_value"], table_min_col))
        table_actual = as_number(block_get(table_block, table_idx, column_map["actual_value"], table_min_col))
        table_max = as_number(block_get(table_block, table_idx, column_map["forecast_max"], table_min_col))
        table_min = as_number(block_get(table_block, table_idx, column_map["forecast_min"], table_min_col))

        intercept = as_number(intercept_cell.value)
        slope = as_number(slope_cell.value)
        forecast_value = first_not_none(table_forecast, as_number(forecast_cell.value))
        forecast_max = first_not_none(table_max, as_number(max_cell.value))
        forecast_min = first_not_none(table_min, as_number(min_cell.value))

        signature = (
            round(intercept, 12) if intercept is not None else None,
            round(slope, 12) if slope is not None else None,
            round(forecast_value, 12) if forecast_value is not None else None,
            round(forecast_max, 12) if forecast_max is not None else None,
            round(forecast_min, 12) if forecast_min is not None else None,
        )
        if signature == previous_signature:
            continue
        previous_signature = signature

        row = {
            "model": metadata["model"],
            "ticker": metadata["ticker"],
            "model_period": metadata["model_period"],
            "model_date": metadata["model_date"],
            "method": "regression",
            "parameter_name": "num_quarters_used",
            "parameter_value": n_quarters_used,
            "num_quarters_used": n_quarters_used,
            "forecast_value": forecast_value,
            "actual_value": table_actual,
            "forecast_max": forecast_max,
            "forecast_min": forecast_min,
            "range_width": subtract_or_none(forecast_max, forecast_min),
            "intercept": intercept,
            "slope": slope,
            "source_file": source_file,
        }
        rows.append(row)

    sheet.range((helper_row, helper_col), (helper_row, helper_col + 4)).clear_contents()
    return rows


def autosize_columns(ws) -> None:
    for col_idx in range(1, ws.max_column + 1):
        header = ws.cell(row=1, column=col_idx).value
        max_length = len(str(header)) if header is not None else 10
        for row_idx in range(2, ws.max_row + 1):
            cell_value = ws.cell(row=row_idx, column=col_idx).value
            if cell_value is None:
                continue
            max_length = max(max_length, len(str(cell_value)))
        ws.column_dimensions[get_column_letter(col_idx)].width = min(max(12, max_length + 2), 48)


def write_sheet(ws, columns: Sequence[str], rows: Sequence[Dict[str, Any]]) -> None:
    ws.append(list(columns))
    for row in rows:
        ws.append([row.get(column) for column in columns])

    for header_cell in ws[1]:
        header_cell.font = Font(bold=True)

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    autosize_columns(ws)


def should_skip_file(path: Path, input_folder_name: str) -> Optional[str]:
    if not path.is_file():
        return "not a file"
    if path.name.startswith("~"):
        return "temporary workbook"
    if path.suffix.lower() != ".xlsx":
        return "not an .xlsx file"
    if re.match(rf"^{re.escape(input_folder_name)}_PARAM(\.\d+)?\.xlsx$", path.name, flags=re.IGNORECASE):
        return "generated output workbook"
    return None


def main() -> None:
    output_path = next_output_path(input_dir, output_dir)
    if not input_dir.exists():
        raise FileNotFoundError(f"Input directory does not exist: {input_dir}")

    all_paths = sorted(input_dir.iterdir(), key=lambda p: p.name.lower())
    processed_files = 0
    empirical_rows: List[Dict[str, Any]] = []
    regression_rows: List[Dict[str, Any]] = []

    with xw.App(visible=False, add_book=False) as app:
        app.display_alerts = False
        app.screen_updating = False
        try:
            app.calculation = "manual"
        except Exception:
            pass

        for path in all_paths:
            skip_reason = should_skip_file(path, input_dir.name)
            if skip_reason:
                print(f"Skipping {path.name}: {skip_reason}")
                continue

            print(f"Processing {path.name}")
            wb: Optional[xw.main.Book] = None
            try:
                metadata = parse_file_metadata(path.name)
                wb = app.books.open(str(path), update_links=False)

                empirical_sheet = wb.sheets[EMPIRICAL_SHEET_NAME]
                regression_sheet = wb.sheets[REGRESSION_SHEET_NAME]

                empirical_rows.extend(
                    extract_empirical_rows(
                        wb=wb,
                        sheet=empirical_sheet,
                        metadata=metadata,
                        source_file=path.name,
                    )
                )
                regression_rows.extend(
                    extract_regression_rows(
                        wb=wb,
                        sheet=regression_sheet,
                        metadata=metadata,
                        source_file=path.name,
                    )
                )
                processed_files += 1
            except KeyError as exc:
                print(f"Skipping {path.name}: missing sheet {exc}")
            except Exception as exc:
                print(f"Skipping {path.name}: processing error: {exc}")
            finally:
                if wb is not None:
                    safe_close_workbook(wb)

    out_wb = Workbook()
    empirical_ws = out_wb.active
    empirical_ws.title = "empirical_candidates"
    write_sheet(empirical_ws, EMPIRICAL_COLUMNS, empirical_rows)

    regression_ws = out_wb.create_sheet("regression_candidates")
    write_sheet(regression_ws, REGRESSION_COLUMNS, regression_rows)

    out_wb.save(output_path)

    print(f"Output written: {output_path}")
    print(f"Files processed: {processed_files}")
    print(f"Empirical rows: {len(empirical_rows)}")
    print(f"Regression rows: {len(regression_rows)}")


if __name__ == "__main__":
    main()
