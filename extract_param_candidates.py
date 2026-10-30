from __future__ import annotations

import calendar
import math
import re
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Dict, List, Optional, Sequence, Tuple

import openpyxl
import xlwings as xw
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter


# ----------------------------
# User-configurable paths
# ----------------------------
input_dir = Path("/path/to/input_folder")
output_dir = Path("/path/to/output_folder")


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
TIMING_TO_DAY = {"early": 5, "mid": 15, "late": 25}
MONTH_TO_NUM = {
    "jan": 1,
    "january": 1,
    "feb": 2,
    "february": 2,
    "mar": 3,
    "march": 3,
    "apr": 4,
    "april": 4,
    "may": 5,
    "jun": 6,
    "june": 6,
    "jul": 7,
    "july": 7,
    "aug": 8,
    "august": 8,
    "sep": 9,
    "sept": 9,
    "september": 9,
    "oct": 10,
    "october": 10,
    "nov": 11,
    "november": 11,
    "dec": 12,
    "december": 12,
}


@dataclass
class SheetContext:
    sheet: Any
    used_row: int
    used_col: int
    last_row: int
    last_col: int
    label_positions: Dict[str, List[Tuple[int, int]]]
    anchor_row: int
    anchor_col: int


def normalize_label(value: Any) -> str:
    if not isinstance(value, str):
        return ""
    return re.sub(r"[^a-z0-9]+", " ", value.strip().lower()).strip()


def to_float(value: Any) -> Optional[float]:
    if value is None or value == "":
        return None

    if isinstance(value, (int, float)):
        num = float(value)
        if math.isnan(num):
            return None
        return num

    if isinstance(value, str):
        raw = value.strip()
        if not raw:
            return None
        has_pct = "%" in raw
        cleaned = raw.replace(",", "").replace("%", "").strip()
        try:
            parsed = float(cleaned)
        except ValueError:
            return None
        return parsed / 100.0 if has_pct else parsed

    return None


def flatten_column_values(values: Any, expected_len: int) -> List[Any]:
    if isinstance(values, list):
        if values and isinstance(values[0], list):
            flat = [row[0] if row else None for row in values]
        else:
            flat = values
    else:
        flat = [values]

    if len(flat) < expected_len:
        flat.extend([None] * (expected_len - len(flat)))
    return flat[:expected_len]


def get_next_output_path(input_path: Path, out_dir: Path) -> Path:
    folder_name = input_path.name
    base_name = f"{folder_name}_PARAM"
    candidate = out_dir / f"{base_name}.xlsx"
    counter = 1
    while candidate.exists():
        candidate = out_dir / f"{base_name}.{counter}.xlsx"
        counter += 1
    return candidate


def parse_file_metadata(file_path: Path) -> Dict[str, str]:
    stem = file_path.stem
    parts = [part.strip() for part in stem.split(" - ")]
    if len(parts) < 3:
        raise ValueError("filename does not match expected '... - TICKER - MidJan2026_...' format")

    ticker = parts[1].upper()
    period_token = parts[2].split("_")[0]
    compact = re.sub(r"[^A-Za-z0-9]", "", period_token)
    match = re.match(r"^(Early|Mid|Late)([A-Za-z]+)(\d{4})$", compact, flags=re.IGNORECASE)
    if not match:
        raise ValueError("cannot parse model period token (expected Early/Mid/Late + Month + Year)")

    timing_raw, month_raw, year_raw = match.groups()
    timing = timing_raw.lower()
    month_key = month_raw.lower()
    month_num = MONTH_TO_NUM.get(month_key) or MONTH_TO_NUM.get(month_key[:3])
    if month_num is None:
        raise ValueError(f"unsupported month in period token: {month_raw}")

    day = TIMING_TO_DAY[timing]
    timing_title = timing_raw[0].upper() + timing_raw[1:].lower()
    month_abbr = calendar.month_abbr[month_num]
    model_period = f"{timing_title}{month_abbr}_{year_raw}"
    model_date = f"{int(year_raw):04d}-{month_num:02d}-{day:02d}"
    model = f"{ticker}_{model_period}"

    return {
        "model": model,
        "ticker": ticker,
        "model_period": model_period,
        "model_date": model_date,
    }


def build_sheet_context(sheet: Any) -> SheetContext:
    used = sheet.used_range
    used_row = int(used.row)
    used_col = int(used.column)
    last_row = int(used.last_cell.row)
    last_col = int(used.last_cell.column)

    values = used.value
    if values is None:
        matrix: List[List[Any]] = []
    elif isinstance(values, list):
        if values and isinstance(values[0], list):
            matrix = values
        else:
            matrix = [values]
    else:
        matrix = [[values]]

    label_positions: Dict[str, List[Tuple[int, int]]] = {}
    max_positions: List[Tuple[int, int]] = []

    for r_offset, row_values in enumerate(matrix):
        if not isinstance(row_values, list):
            row_values = [row_values]
        for c_offset, cell_value in enumerate(row_values):
            label = normalize_label(cell_value)
            if not label:
                continue
            row = used_row + r_offset
            col = used_col + c_offset
            label_positions.setdefault(label, []).append((row, col))
            if label == "max":
                max_positions.append((row, col))

    if not max_positions:
        raise ValueError("could not find 'max' anchor cell")

    # In model templates the anchor tends to be near the lower output area.
    anchor_row, anchor_col = max(max_positions, key=lambda p: p[0])

    return SheetContext(
        sheet=sheet,
        used_row=used_row,
        used_col=used_col,
        last_row=last_row,
        last_col=last_col,
        label_positions=label_positions,
        anchor_row=anchor_row,
        anchor_col=anchor_col,
    )


def find_label_position(
    context: SheetContext,
    candidates: Sequence[str],
    anchor: Optional[Tuple[int, int]] = None,
) -> Optional[Tuple[int, int]]:
    normalized = [normalize_label(item) for item in candidates if item]
    normalized = [item for item in normalized if item]
    if not normalized:
        return None

    matches: List[Tuple[int, int]] = []
    for label, positions in context.label_positions.items():
        if any(candidate == label for candidate in normalized) or any(candidate in label for candidate in normalized):
            matches.extend(positions)

    if not matches:
        return None
    if anchor is None:
        return matches[0]

    a_row, a_col = anchor
    return min(matches, key=lambda p: abs(p[0] - a_row) + abs(p[1] - a_col))


def cell_from_label(
    context: SheetContext,
    candidates: Sequence[str],
    *,
    value_to_right: bool = True,
    fallback: Optional[Tuple[int, int]] = None,
) -> Optional[Tuple[int, int]]:
    pos = find_label_position(
        context,
        candidates,
        anchor=(context.anchor_row, context.anchor_col),
    )
    if pos is None:
        return fallback
    row, col = pos
    return (row, col + 1) if value_to_right else (row, col)


def read_cell(sheet: Any, cell_ref: Optional[Tuple[int, int]]) -> Any:
    if cell_ref is None:
        return None
    row, col = cell_ref
    return sheet.range((row, col)).value


def capture_cell_state(sheet: Any, cell_ref: Optional[Tuple[int, int]]) -> Optional[Tuple[bool, Any, Any]]:
    if cell_ref is None:
        return None
    rng = sheet.range(cell_ref)
    formula2 = None
    has_formula2 = False
    try:
        formula2 = rng.formula2
        has_formula2 = formula2 not in (None, "")
    except Exception:
        has_formula2 = False
    value = rng.value
    return (has_formula2, formula2, value)


def restore_cell_state(sheet: Any, cell_ref: Optional[Tuple[int, int]], state: Optional[Tuple[bool, Any, Any]]) -> None:
    if cell_ref is None or state is None:
        return
    has_formula2, formula2, value = state
    rng = sheet.range(cell_ref)
    if has_formula2:
        try:
            rng.formula2 = formula2
            return
        except Exception:
            try:
                rng.api.Formula2 = formula2
                return
            except Exception:
                pass
    rng.value = value


def set_formula2_r1c1(sheet: Any, cell_ref: Tuple[int, int], formula_r1c1: str) -> None:
    rng = sheet.range(cell_ref)
    try:
        rng.api.Formula2R1C1 = formula_r1c1
    except Exception:
        # Fallbacks preserve R1C1 behavior if Formula2R1C1 is unavailable.
        try:
            rng.formula2 = formula_r1c1
        except Exception:
            rng.api.FormulaR1C1 = formula_r1c1


def get_numeric_rows(sheet: Any, col: int, row_start: int, row_end: int) -> List[int]:
    if row_end < row_start:
        return []

    values = sheet.range((row_start, col), (row_end, col)).value
    flat = flatten_column_values(values, row_end - row_start + 1)

    rows: List[int] = []
    for offset, value in enumerate(flat):
        if to_float(value) is not None:
            rows.append(row_start + offset)
    return rows


def get_numeric_pairs(sheet: Any, x_col: int, y_col: int, row_start: int, row_end: int) -> List[Tuple[int, float, float]]:
    if row_end < row_start:
        return []

    length = row_end - row_start + 1
    x_values = flatten_column_values(sheet.range((row_start, x_col), (row_end, x_col)).value, length)
    y_values = flatten_column_values(sheet.range((row_start, y_col), (row_end, y_col)).value, length)

    pairs: List[Tuple[int, float, float]] = []
    for offset, (x_val, y_val) in enumerate(zip(x_values, y_values)):
        x_num = to_float(x_val)
        y_num = to_float(y_val)
        if x_num is None or y_num is None:
            continue
        pairs.append((row_start + offset, x_num, y_num))
    return pairs


def calc_range_width(forecast_max: Any, forecast_min: Any) -> Optional[float]:
    max_num = to_float(forecast_max)
    min_num = to_float(forecast_min)
    if max_num is None or min_num is None:
        return None
    return max_num - min_num


def safe_close_workbook(wb: Any) -> None:
    close_errors: List[Exception] = []
    close_attempts = [
        lambda: wb.close(save=False),
        lambda: wb.close(SaveChanges=False),
        lambda: wb.close(),
        lambda: wb.api.Close(SaveChanges=False),
    ]

    for closer in close_attempts:
        try:
            closer()
            return
        except TypeError:
            continue
        except Exception as exc:  # noqa: PERF203 - keep error detail for final failure
            close_errors.append(exc)

    details = "; ".join(str(err) for err in close_errors) if close_errors else "unknown close error"
    raise RuntimeError(f"failed to close workbook safely: {details}")


def get_last_quarter_value(sheet: Any, quarter_col: int, row: int) -> Any:
    if quarter_col <= 0 or row <= 0:
        return None
    return sheet.range((row, quarter_col)).value


def process_empirical_sheet(
    wb: Any,
    metadata: Dict[str, str],
    source_file: str,
) -> List[Dict[str, Any]]:
    sheet = wb.sheets["Empirical Model"]
    context = build_sheet_context(sheet)

    anchor_row = context.anchor_row
    anchor_col = context.anchor_col
    penetration_col = anchor_col - 11
    quarter_col = penetration_col - 1
    data_rows = get_numeric_rows(sheet, penetration_col, context.used_row, anchor_row - 1)

    if not data_rows:
        return []

    avg_pen_cell = cell_from_label(
        context,
        ["avg penetration pct", "avg penetration", "avg_penetration_pct", "penetration avg"],
        fallback=(anchor_row - 2, anchor_col + 1),
    )
    num_quarters_cell = cell_from_label(
        context,
        ["num quarters used", "num_quarters_used", "quarters used", "n quarters"],
        fallback=None,
    )
    forecast_value_cell = cell_from_label(
        context,
        ["estimated total sold", "est total sold", "estimated sold", "forecast value"],
        fallback=(anchor_row - 1, anchor_col + 1),
    )
    actual_value_cell = cell_from_label(
        context,
        ["reported sales", "actual sales", "reported"],
        fallback=None,
    )
    quarterly_sales_cell = cell_from_label(
        context,
        ["quarterly sales", "quarter sales"],
        fallback=None,
    )
    growth_rate_cell = cell_from_label(
        context,
        ["growth rate pct", "growth rate", "growth %"],
        fallback=None,
    )
    sales_captured_cell = cell_from_label(
        context,
        ["sales captured in db pct", "sales captured in db", "captured in db", "db capture"],
        fallback=None,
    )
    forecast_max_cell = (anchor_row, anchor_col + 1)
    forecast_min_cell = cell_from_label(
        context,
        ["min"],
        fallback=(anchor_row + 1, anchor_col + 1),
    )

    avg_pen_state = capture_cell_state(sheet, avg_pen_cell)
    num_quarters_state = capture_cell_state(sheet, num_quarters_cell)

    rows: List[Dict[str, Any]] = []
    max_n = min(N_QUARTERS, len(data_rows))

    try:
        for n in range(1, max_n + 1):
            selected_rows = data_rows[-n:]
            start_row = selected_rows[0]
            end_row = selected_rows[-1]
            avg_formula = f"=AVERAGE(R{start_row}C{penetration_col}:R{end_row}C{penetration_col})"

            set_formula2_r1c1(sheet, avg_pen_cell, avg_formula)
            if num_quarters_cell is not None:
                sheet.range(num_quarters_cell).value = n
            wb.app.calculate()

            avg_penetration = read_cell(sheet, avg_pen_cell)
            forecast_value = read_cell(sheet, forecast_value_cell)
            actual_value = read_cell(sheet, actual_value_cell)
            forecast_max = read_cell(sheet, forecast_max_cell)
            forecast_min = read_cell(sheet, forecast_min_cell)
            quarterly_sales = read_cell(sheet, quarterly_sales_cell)
            reported_sales = read_cell(sheet, actual_value_cell)
            growth_rate_pct = read_cell(sheet, growth_rate_cell)
            sales_captured_pct = read_cell(sheet, sales_captured_cell)
            last_quarter_used = get_last_quarter_value(sheet, quarter_col, end_row)

            rows.append(
                {
                    **metadata,
                    "method": "empirical",
                    "parameter_name": "avg_penetration_pct",
                    "parameter_value": avg_penetration,
                    "num_quarters_used": n,
                    "last_quarter_used": last_quarter_used,
                    "forecast_value": forecast_value,
                    "actual_value": actual_value,
                    "forecast_max": forecast_max,
                    "forecast_min": forecast_min,
                    "range_width": calc_range_width(forecast_max, forecast_min),
                    "avg_penetration_pct": avg_penetration,
                    "quarterly_sales": quarterly_sales,
                    "reported_sales": reported_sales,
                    "growth_rate_pct": growth_rate_pct,
                    "sales_captured_in_db_pct": sales_captured_pct,
                    "source_file": source_file,
                }
            )
    finally:
        restore_cell_state(sheet, avg_pen_cell, avg_pen_state)
        restore_cell_state(sheet, num_quarters_cell, num_quarters_state)
        wb.app.calculate()

    return rows


def is_duplicate_regression_row(prev: Dict[str, Any], curr: Dict[str, Any]) -> bool:
    keys = ["num_quarters_used", "intercept", "slope", "forecast_value", "forecast_max", "forecast_min"]
    tolerance = 1e-10

    for key in keys:
        prev_num = to_float(prev.get(key))
        curr_num = to_float(curr.get(key))
        if prev_num is not None and curr_num is not None:
            if abs(prev_num - curr_num) > tolerance:
                return False
            continue
        if prev.get(key) != curr.get(key):
            return False
    return True


def process_regression_sheet(
    wb: Any,
    metadata: Dict[str, str],
    source_file: str,
) -> List[Dict[str, Any]]:
    sheet = wb.sheets["Regression Model"]
    context = build_sheet_context(sheet)

    anchor_row = context.anchor_row
    anchor_col = context.anchor_col
    y_col = anchor_col - 7
    x_col = anchor_col - 11
    data_pairs = get_numeric_pairs(sheet, x_col, y_col, context.used_row, anchor_row - 1)

    if not data_pairs:
        return []

    scratch_col = context.last_col + 2

    num_quarters_cell = cell_from_label(
        context,
        ["num quarters used", "num_quarters_used", "quarters used", "n quarters"],
        fallback=None,
    )
    intercept_cell = cell_from_label(
        context,
        ["intercept"],
        fallback=(anchor_row, scratch_col),
    )
    slope_cell = cell_from_label(
        context,
        ["slope"],
        fallback=(anchor_row + 1, scratch_col),
    )
    forecast_cell = cell_from_label(
        context,
        ["tot fcst w o sa", "tot fcst without sa", "forecast total without sa", "tot fcst wo sa", "tot fcst w sa"],
        fallback=None,
    )
    actual_value_cell = cell_from_label(
        context,
        ["reported sales", "actual sales", "actual"],
        fallback=None,
    )
    forecast_max_cell = (anchor_row, anchor_col + 1)
    forecast_min_cell = cell_from_label(
        context,
        ["min"],
        fallback=(anchor_row + 1, anchor_col + 1),
    )

    intercept_state = capture_cell_state(sheet, intercept_cell)
    slope_state = capture_cell_state(sheet, slope_cell)
    num_quarters_state = capture_cell_state(sheet, num_quarters_cell)

    rows: List[Dict[str, Any]] = []
    max_n = min(N_QUARTERS, len(data_pairs))

    try:
        for n in range(1, max_n + 1):
            selected = data_pairs[-n:]
            start_row = selected[0][0]
            end_row = selected[-1][0]
            intercept_formula = (
                f"=INTERCEPT(R{start_row}C{y_col}:R{end_row}C{y_col},R{start_row}C{x_col}:R{end_row}C{x_col})"
            )
            slope_formula = f"=SLOPE(R{start_row}C{y_col}:R{end_row}C{y_col},R{start_row}C{x_col}:R{end_row}C{x_col})"

            set_formula2_r1c1(sheet, intercept_cell, intercept_formula)
            set_formula2_r1c1(sheet, slope_cell, slope_formula)
            if num_quarters_cell is not None:
                sheet.range(num_quarters_cell).value = n
            wb.app.calculate()

            intercept = read_cell(sheet, intercept_cell)
            slope = read_cell(sheet, slope_cell)
            forecast_value = read_cell(sheet, forecast_cell)

            if forecast_value is None:
                last_x = selected[-1][1]
                intercept_num = to_float(intercept)
                slope_num = to_float(slope)
                if intercept_num is not None and slope_num is not None:
                    forecast_value = intercept_num + slope_num * (last_x + 1.0)

            forecast_max = read_cell(sheet, forecast_max_cell)
            forecast_min = read_cell(sheet, forecast_min_cell)
            actual_value = read_cell(sheet, actual_value_cell)

            row = {
                **metadata,
                "method": "regression",
                "parameter_name": "num_quarters_used",
                "parameter_value": n,
                "num_quarters_used": n,
                "forecast_value": forecast_value,
                "actual_value": actual_value if actual_value is not None else "",
                "forecast_max": forecast_max,
                "forecast_min": forecast_min,
                "range_width": calc_range_width(forecast_max, forecast_min),
                "intercept": intercept,
                "slope": slope,
                "source_file": source_file,
            }

            if rows and is_duplicate_regression_row(rows[-1], row):
                continue
            rows.append(row)
    finally:
        restore_cell_state(sheet, intercept_cell, intercept_state)
        restore_cell_state(sheet, slope_cell, slope_state)
        restore_cell_state(sheet, num_quarters_cell, num_quarters_state)
        wb.app.calculate()

    return rows


def write_rows_to_sheet(
    ws: openpyxl.worksheet.worksheet.Worksheet,
    columns: Sequence[str],
    rows: Sequence[Dict[str, Any]],
) -> None:
    ws.append(list(columns))
    for row in rows:
        ws.append([row.get(col, "") for col in columns])

    for cell in ws[1]:
        cell.font = Font(bold=True)
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(columns))}{ws.max_row}"

    for col_idx, header in enumerate(columns, start=1):
        max_len = len(header)
        for row_idx in range(2, ws.max_row + 1):
            value = ws.cell(row=row_idx, column=col_idx).value
            if value is None:
                continue
            text = str(value)
            if len(text) > max_len:
                max_len = len(text)
        ws.column_dimensions[get_column_letter(col_idx)].width = max(12, min(max_len + 2, 48))


def process_workbooks(in_dir: Path, out_dir: Path) -> Tuple[Path, int, List[Dict[str, Any]], List[Dict[str, Any]]]:
    out_dir.mkdir(parents=True, exist_ok=True)
    output_path = get_next_output_path(in_dir, out_dir)

    empirical_rows: List[Dict[str, Any]] = []
    regression_rows: List[Dict[str, Any]] = []
    processed_files = 0

    app = xw.App(visible=False, add_book=False)
    app.display_alerts = False
    app.screen_updating = False

    try:
        for file_path in sorted(in_dir.iterdir(), key=lambda p: p.name.lower()):
            if not file_path.is_file():
                continue
            if file_path.name.startswith("~"):
                print(f"skipped file: {file_path.name} (temp file)")
                continue
            if file_path.suffix.lower() != ".xlsx":
                print(f"skipped file: {file_path.name} (not .xlsx)")
                continue

            wb = None
            try:
                metadata = parse_file_metadata(file_path)
                wb = app.books.open(str(file_path), update_links=False)
                empirical_rows.extend(process_empirical_sheet(wb, metadata, file_path.name))
                regression_rows.extend(process_regression_sheet(wb, metadata, file_path.name))
                processed_files += 1
                print(f"processed file: {file_path.name}")
            except Exception as exc:
                print(f"skipped file: {file_path.name} ({exc})")
            finally:
                if wb is not None:
                    safe_close_workbook(wb)
    finally:
        app.quit()

    out_wb = openpyxl.Workbook()
    default_sheet = out_wb.active
    out_wb.remove(default_sheet)
    empirical_ws = out_wb.create_sheet("empirical_candidates")
    regression_ws = out_wb.create_sheet("regression_candidates")

    write_rows_to_sheet(empirical_ws, EMPIRICAL_COLUMNS, empirical_rows)
    write_rows_to_sheet(regression_ws, REGRESSION_COLUMNS, regression_rows)
    out_wb.save(output_path)

    return output_path, processed_files, empirical_rows, regression_rows


def main() -> None:
    in_dir = Path(input_dir)
    out_dir = Path(output_dir)

    if not in_dir.exists() or not in_dir.is_dir():
        raise FileNotFoundError(f"input_dir does not exist or is not a directory: {in_dir}")

    output_path, processed_files, empirical_rows, regression_rows = process_workbooks(in_dir, out_dir)

    print(f"output path: {output_path}")
    print(f"number of files processed: {processed_files}")
    print(f"number of empirical rows: {len(empirical_rows)}")
    print(f"number of regression rows: {len(regression_rows)}")


if __name__ == "__main__":
    main()
