#!/usr/bin/env python3
from __future__ import annotations

from datetime import date
from pathlib import Path
import re
from typing import Any, Dict, List, Optional, Sequence, Tuple

import xlwings as xw
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter


# Configure folders here.
input_dir = Path("input")
output_dir = Path("output")

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

PERIOD_DAY = {"early": 5, "mid": 15, "late": 25}

EMPIRICAL_HEADER_ALIASES: Dict[str, Sequence[str]] = {
    "num_quarters_used": (
        "num quarters used",
        "number of quarters used",
        "quarters used",
        "num qtrs used",
        "n_quarters",
        "n quarters",
        "qtrs",
    ),
    "last_quarter_used": (
        "last quarter used",
        "last qtr used",
        "last quarter",
    ),
    "forecast_value": (
        "estimated total sold",
        "forecast",
        "forecast value",
        "tot fcst",
    ),
    "actual_value": (
        "reported sales",
        "actual",
        "actual value",
    ),
    "forecast_max": ("max", "forecast max"),
    "forecast_min": ("min", "forecast min"),
    "avg_penetration_pct": (
        "avg penetration",
        "avg penetration pct",
        "average penetration",
        "penetration pct",
    ),
    "quarterly_sales": ("quarterly sales",),
    "reported_sales": ("reported sales",),
    "growth_rate_pct": ("growth rate", "growth rate pct"),
    "sales_captured_in_db_pct": ("sales captured in db pct", "captured in db"),
}

# Fallback offsets relative to the "max" anchor column.
EMPIRICAL_FALLBACK_OFFSETS = {
    "num_quarters_used": -13,
    "last_quarter_used": -12,
    "forecast_value": -11,
    "actual_value": -10,
    "quarterly_sales": -9,
    "reported_sales": -8,
    "growth_rate_pct": -7,
    "sales_captured_in_db_pct": -6,
    "avg_penetration_pct": -5,
    "forecast_max": 0,
    "forecast_min": 1,
}

REGRESSION_HEADER_ALIASES: Dict[str, Sequence[str]] = {
    "num_quarters_used": (
        "num quarters used",
        "number of quarters used",
        "quarters used",
        "num qtrs used",
        "n_quarters",
    ),
    "forecast_value": (
        "tot fcst w/o sa",
        "tot fcst wo sa",
        "total forecast without sa",
        "forecast w/o sa",
    ),
    "actual_value": ("actual", "actual value", "reported sales"),
    "forecast_max": ("max", "forecast max"),
    "forecast_min": ("min", "forecast min"),
}

REGRESSION_FALLBACK_OFFSETS = {
    "num_quarters_used": -13,
    "forecast_value": -10,
    "actual_value": -9,
    "forecast_max": 0,
    "forecast_min": 1,
}


def normalize_text(value: Any) -> str:
    if value is None:
        return ""
    text = str(value).strip().lower()
    text = re.sub(r"[^a-z0-9]+", " ", text)
    return text.strip()


def as_2d(values: Any) -> List[List[Any]]:
    if values is None:
        return []
    if isinstance(values, list):
        if not values:
            return []
        if isinstance(values[0], list):
            return values
        return [values]
    return [[values]]


def is_blank(value: Any) -> bool:
    return value is None or (isinstance(value, str) and value.strip() == "")


def to_float(value: Any) -> Optional[float]:
    if value is None:
        return None
    if isinstance(value, (int, float)):
        return float(value)
    if isinstance(value, str):
        txt = value.strip().replace(",", "")
        if not txt:
            return None
        try:
            return float(txt)
        except ValueError:
            return None
    return None


def parse_file_metadata(file_path: Path) -> Dict[str, str]:
    stem = file_path.stem
    parts = [part.strip() for part in stem.split(" - ")]

    ticker = parts[1].upper() if len(parts) > 1 else "UNKNOWN"
    period_token = parts[2] if len(parts) > 2 else ""
    period_token = re.sub(r"_.*$", "", period_token).strip()

    model_period = "UNKNOWN_0000"
    model_date = ""

    period_match = re.match(
        r"^(Early|Mid|Late)(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)(\d{4})$",
        period_token,
        re.IGNORECASE,
    )
    if period_match:
        timing = period_match.group(1).title()
        month_abbr = period_match.group(2).title()
        year = int(period_match.group(3))
        model_period = f"{timing}{month_abbr}_{year}"

        month_num = MONTH_TO_NUM[month_abbr.lower()]
        day = PERIOD_DAY[timing.lower()]
        model_date = date(year, month_num, day).isoformat()

    model = f"{ticker}_{model_period}"
    return {
        "ticker": ticker,
        "model_period": model_period,
        "model_date": model_date,
        "model": model,
    }


def build_output_path(input_path: Path, output_path: Path) -> Path:
    output_path.mkdir(parents=True, exist_ok=True)
    base_name = f"{input_path.name}_PARAM"

    candidate = output_path / f"{base_name}.xlsx"
    counter = 1
    while candidate.exists():
        candidate = output_path / f"{base_name}.{counter}.xlsx"
        counter += 1
    return candidate


def close_workbook_safely(wb: xw.Book) -> None:
    try:
        wb.close(save=False)
        return
    except TypeError:
        pass
    except Exception:
        pass

    # Older xlwings / Excel API fallback paths.
    for closer in (
        lambda: wb.close(False),
        lambda: wb.api.Close(SaveChanges=False),
        lambda: wb.api.Close(False),
    ):
        try:
            closer()
            return
        except Exception:
            continue


def find_anchor(sheet: xw.Sheet, target: str = "max") -> Optional[Tuple[int, int]]:
    used_range = sheet.used_range
    values_2d = as_2d(used_range.value)
    if not values_2d:
        return None

    target_norm = normalize_text(target)
    base_row = used_range.row
    base_col = used_range.column

    for r_idx, row_vals in enumerate(values_2d):
        for c_idx, cell_val in enumerate(row_vals):
            if normalize_text(cell_val) == target_norm:
                return base_row + r_idx, base_col + c_idx
    return None


def header_map_for_row(
    sheet: xw.Sheet, row: int, col_start: int, col_end: int
) -> Dict[str, int]:
    col_start = max(1, col_start)
    col_end = max(col_start, col_end)
    values = sheet.range((row, col_start), (row, col_end)).value
    values_1d = values if isinstance(values, list) else [values]

    mapping: Dict[str, int] = {}
    for idx, value in enumerate(values_1d):
        key = normalize_text(value)
        if key:
            mapping[key] = col_start + idx
    return mapping


def resolve_column(
    header_map: Dict[str, int],
    aliases: Sequence[str],
    anchor_col: int,
    fallback_offset: Optional[int],
) -> Optional[int]:
    normalized_aliases = [normalize_text(a) for a in aliases]

    for alias in normalized_aliases:
        if alias in header_map:
            return header_map[alias]

    for header_text, col in header_map.items():
        for alias in normalized_aliases:
            if alias and (alias in header_text or header_text in alias):
                return col

    if fallback_offset is None:
        return None
    return anchor_col + fallback_offset


def set_formula2(cell: xw.Range, formula: str) -> None:
    try:
        cell.formula2 = formula
    except Exception:
        cell.api.Formula2 = formula


def rows_are_equivalent(
    left: Dict[str, Any], right: Dict[str, Any], keys: Sequence[str], tol: float = 1e-9
) -> bool:
    for key in keys:
        l_val = left.get(key)
        r_val = right.get(key)
        l_num = to_float(l_val)
        r_num = to_float(r_val)

        if l_num is not None and r_num is not None:
            if abs(l_num - r_num) > tol:
                return False
        else:
            if l_val != r_val:
                return False
    return True


def process_empirical_sheet(
    wb: xw.Book, metadata: Dict[str, str], source_file: str
) -> List[Dict[str, Any]]:
    if "Empirical Model" not in [sheet.name for sheet in wb.sheets]:
        return []

    sheet = wb.sheets["Empirical Model"]
    anchor = find_anchor(sheet, target="max")
    if anchor is None:
        return []

    anchor_row, anchor_col = anchor
    start_row = anchor_row + 1
    end_row = start_row + N_QUARTERS - 1

    header_map = header_map_for_row(sheet, anchor_row, anchor_col - 25, anchor_col + 25)
    resolved_cols: Dict[str, Optional[int]] = {}
    for key, aliases in EMPIRICAL_HEADER_ALIASES.items():
        resolved_cols[key] = resolve_column(
            header_map,
            aliases,
            anchor_col,
            EMPIRICAL_FALLBACK_OFFSETS.get(key),
        )

    candidate_cols = [col for col in resolved_cols.values() if col is not None and col > 0]
    if not candidate_cols:
        return []

    min_col = min(candidate_cols)
    max_col = max(candidate_cols)
    data_block = as_2d(sheet.range((start_row, min_col), (end_row, max_col)).value)

    def read_field(row_idx: int, field: str) -> Any:
        col = resolved_cols.get(field)
        if col is None or col < min_col or col > max_col:
            return None
        if row_idx >= len(data_block):
            return None
        row_data = data_block[row_idx]
        offset = col - min_col
        if offset >= len(row_data):
            return None
        return row_data[offset]

    # Temporary formula column for avg penetration calc (R1C1/formula2 as required).
    avg_formula_col = anchor_col + 30
    wrote_formulas = False
    for idx in range(N_QUARTERS):
        row_num = start_row + idx
        avg_col = resolved_cols.get("avg_penetration_pct")
        q_sales_col = resolved_cols.get("quarterly_sales")
        rep_sales_col = resolved_cols.get("reported_sales")

        formula: Optional[str] = None
        if avg_col is not None:
            formula = f"=RC[{avg_col - avg_formula_col}]"
        elif q_sales_col is not None and rep_sales_col is not None:
            formula = (
                f"=IFERROR(RC[{q_sales_col - avg_formula_col}]"
                f"/RC[{rep_sales_col - avg_formula_col}],0)"
            )

        if formula:
            set_formula2(sheet.cells(row_num, avg_formula_col), formula)
            wrote_formulas = True

    avg_values = [[None] for _ in range(N_QUARTERS)]
    if wrote_formulas:
        wb.app.calculate()
        avg_values = as_2d(
            sheet.range((start_row, avg_formula_col), (end_row, avg_formula_col)).value
        )
        sheet.range((start_row, avg_formula_col), (end_row, avg_formula_col)).clear_contents()

    rows: List[Dict[str, Any]] = []
    for idx in range(N_QUARTERS):
        num_quarters_used = read_field(idx, "num_quarters_used")
        last_quarter_used = read_field(idx, "last_quarter_used")
        forecast_value = read_field(idx, "forecast_value")
        actual_value = read_field(idx, "actual_value")
        forecast_max = read_field(idx, "forecast_max")
        forecast_min = read_field(idx, "forecast_min")
        quarterly_sales = read_field(idx, "quarterly_sales")
        reported_sales = read_field(idx, "reported_sales")
        growth_rate_pct = read_field(idx, "growth_rate_pct")
        sales_captured_in_db_pct = read_field(idx, "sales_captured_in_db_pct")

        avg_penetration_pct = (
            avg_values[idx][0]
            if idx < len(avg_values) and avg_values[idx]
            else read_field(idx, "avg_penetration_pct")
        )

        if is_blank(num_quarters_used):
            num_quarters_used = idx + 1

        if all(
            is_blank(v)
            for v in (
                forecast_value,
                actual_value,
                forecast_max,
                forecast_min,
                avg_penetration_pct,
            )
        ):
            continue

        max_num = to_float(forecast_max)
        min_num = to_float(forecast_min)
        range_width = (max_num - min_num) if max_num is not None and min_num is not None else None

        rows.append(
            {
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

    return rows


def process_regression_sheet(
    wb: xw.Book, metadata: Dict[str, str], source_file: str
) -> List[Dict[str, Any]]:
    if "Regression Model" not in [sheet.name for sheet in wb.sheets]:
        return []

    sheet = wb.sheets["Regression Model"]
    anchor = find_anchor(sheet, target="max")
    if anchor is None:
        return []

    anchor_row, anchor_col = anchor
    start_row = anchor_row + 1
    end_row = start_row + N_QUARTERS - 1

    y_col = anchor_col - 7
    x_col = anchor_col - 11

    header_map = header_map_for_row(sheet, anchor_row, anchor_col - 25, anchor_col + 25)
    resolved_cols: Dict[str, Optional[int]] = {}
    for key, aliases in REGRESSION_HEADER_ALIASES.items():
        resolved_cols[key] = resolve_column(
            header_map,
            aliases,
            anchor_col,
            REGRESSION_FALLBACK_OFFSETS.get(key),
        )

    candidate_cols = [col for col in resolved_cols.values() if col is not None and col > 0]
    if candidate_cols:
        min_col = min(candidate_cols)
        max_col = max(candidate_cols)
        data_block = as_2d(sheet.range((start_row, min_col), (end_row, max_col)).value)
    else:
        min_col = 1
        max_col = 1
        data_block = [[None] for _ in range(N_QUARTERS)]

    def read_field(row_idx: int, field: str) -> Any:
        col = resolved_cols.get(field)
        if col is None or col < min_col or col > max_col:
            return None
        if row_idx >= len(data_block):
            return None
        row_data = data_block[row_idx]
        offset = col - min_col
        if offset >= len(row_data):
            return None
        return row_data[offset]

    # Pull numeric history for x/y once, then calculate rolling INTERCEPT/SLOPE.
    history_start_row = max(1, anchor_row - 250)
    history_block = as_2d(sheet.range((history_start_row, x_col), (anchor_row - 1, y_col)).value)
    y_offset = y_col - x_col

    history_rows: List[Tuple[int, float, float]] = []
    for idx, row_vals in enumerate(history_block):
        if y_offset >= len(row_vals):
            continue
        x_value = to_float(row_vals[0])
        y_value = to_float(row_vals[y_offset])
        if x_value is None or y_value is None:
            continue
        history_rows.append((history_start_row + idx, x_value, y_value))

    temp_col_intercept = anchor_col + 30
    temp_col_slope = anchor_col + 31
    temp_col_forecast = anchor_col + 32

    calc_rows = min(N_QUARTERS, len(history_rows))
    wrote_formulas = False
    for idx in range(calc_rows):
        n = idx + 1
        hist_start = history_rows[-n][0]
        hist_end = history_rows[-1][0]
        target_row = start_row + idx

        intercept_formula = (
            f"=INTERCEPT(R{hist_start}C{y_col}:R{hist_end}C{y_col},"
            f"R{hist_start}C{x_col}:R{hist_end}C{x_col})"
        )
        slope_formula = (
            f"=SLOPE(R{hist_start}C{y_col}:R{hist_end}C{y_col},"
            f"R{hist_start}C{x_col}:R{hist_end}C{x_col})"
        )

        next_x = history_rows[-1][1] + 1
        forecast_formula = (
            f"=R{target_row}C{temp_col_slope}*{next_x}+R{target_row}C{temp_col_intercept}"
        )

        set_formula2(sheet.cells(target_row, temp_col_intercept), intercept_formula)
        set_formula2(sheet.cells(target_row, temp_col_slope), slope_formula)
        set_formula2(sheet.cells(target_row, temp_col_forecast), forecast_formula)
        wrote_formulas = True

    calc_block = [[None, None, None] for _ in range(N_QUARTERS)]
    if wrote_formulas:
        wb.app.calculate()
        calc_block = as_2d(
            sheet.range((start_row, temp_col_intercept), (end_row, temp_col_forecast)).value
        )
        sheet.range((start_row, temp_col_intercept), (end_row, temp_col_forecast)).clear_contents()

    rows: List[Dict[str, Any]] = []
    for idx in range(N_QUARTERS):
        num_quarters_used = read_field(idx, "num_quarters_used")
        if is_blank(num_quarters_used):
            num_quarters_used = idx + 1

        calc_row = calc_block[idx] if idx < len(calc_block) else [None, None, None]
        intercept = calc_row[0] if len(calc_row) > 0 else None
        slope = calc_row[1] if len(calc_row) > 1 else None
        calc_forecast = calc_row[2] if len(calc_row) > 2 else None

        forecast_value = read_field(idx, "forecast_value")
        if is_blank(forecast_value):
            forecast_value = calc_forecast

        actual_value = read_field(idx, "actual_value")
        forecast_max = read_field(idx, "forecast_max")
        forecast_min = read_field(idx, "forecast_min")

        if all(
            is_blank(v)
            for v in (
                forecast_value,
                forecast_max,
                forecast_min,
                intercept,
                slope,
            )
        ):
            continue

        max_num = to_float(forecast_max)
        min_num = to_float(forecast_min)
        range_width = (max_num - min_num) if max_num is not None and min_num is not None else None

        rows.append(
            {
                "model": metadata["model"],
                "ticker": metadata["ticker"],
                "model_period": metadata["model_period"],
                "model_date": metadata["model_date"],
                "method": "regression",
                "parameter_name": "num_quarters_used",
                "parameter_value": num_quarters_used,
                "num_quarters_used": num_quarters_used,
                "forecast_value": forecast_value,
                "actual_value": actual_value,
                "forecast_max": forecast_max,
                "forecast_min": forecast_min,
                "range_width": range_width,
                "intercept": intercept,
                "slope": slope,
                "source_file": source_file,
            }
        )

    if len(rows) >= 2 and rows_are_equivalent(
        rows[-1],
        rows[-2],
        keys=(
            "num_quarters_used",
            "forecast_value",
            "forecast_max",
            "forecast_min",
            "intercept",
            "slope",
        ),
    ):
        rows.pop()

    return rows


def auto_size_columns(ws) -> None:
    for col_idx in range(1, ws.max_column + 1):
        max_len = 0
        for row_idx in range(1, ws.max_row + 1):
            value = ws.cell(row=row_idx, column=col_idx).value
            if value is None:
                continue
            max_len = max(max_len, len(str(value)))
        ws.column_dimensions[get_column_letter(col_idx)].width = min(max_len + 2, 42)


def write_output_workbook(
    output_file: Path,
    empirical_rows: Sequence[Dict[str, Any]],
    regression_rows: Sequence[Dict[str, Any]],
) -> None:
    workbook = Workbook()
    empirical_ws = workbook.active
    empirical_ws.title = "empirical_candidates"
    regression_ws = workbook.create_sheet("regression_candidates")

    empirical_ws.append(EMPIRICAL_COLUMNS)
    for row in empirical_rows:
        empirical_ws.append([row.get(col) for col in EMPIRICAL_COLUMNS])

    regression_ws.append(REGRESSION_COLUMNS)
    for row in regression_rows:
        regression_ws.append([row.get(col) for col in REGRESSION_COLUMNS])

    for ws in (empirical_ws, regression_ws):
        for cell in ws[1]:
            cell.font = Font(bold=True)
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions
        auto_size_columns(ws)

    workbook.save(output_file)


def main() -> None:
    input_path = Path(input_dir)
    output_path = Path(output_dir)

    if not input_path.exists():
        print(f"Input directory not found: {input_path}")
        return

    empirical_rows: List[Dict[str, Any]] = []
    regression_rows: List[Dict[str, Any]] = []
    processed_files = 0

    app = xw.App(visible=False, add_book=False)
    app.display_alerts = False
    app.screen_updating = False
    try:
        app.calculation = "manual"
    except Exception:
        pass

    try:
        for file_path in sorted(input_path.iterdir()):
            if not file_path.is_file():
                print(f"Skipped {file_path.name}: not a file")
                continue
            if file_path.name.startswith("~"):
                print(f"Skipped {file_path.name}: temporary file")
                continue
            if file_path.suffix.lower() != ".xlsx":
                print(f"Skipped {file_path.name}: not .xlsx")
                continue

            print(f"Processing {file_path.name}")
            metadata = parse_file_metadata(file_path)
            source_wb: Optional[xw.Book] = None
            try:
                source_wb = app.books.open(str(file_path), update_links=False)

                empirical_rows.extend(
                    process_empirical_sheet(source_wb, metadata, file_path.name)
                )
                regression_rows.extend(
                    process_regression_sheet(source_wb, metadata, file_path.name)
                )
                processed_files += 1
            except Exception as exc:
                print(f"Skipped {file_path.name}: {exc}")
            finally:
                if source_wb is not None:
                    close_workbook_safely(source_wb)
    finally:
        try:
            app.quit()
        except Exception:
            app.kill()

    output_file = build_output_path(input_path, output_path)
    write_output_workbook(output_file, empirical_rows, regression_rows)

    print(f"Output path: {output_file}")
    print(f"Files processed: {processed_files}")
    print(f"Empirical rows: {len(empirical_rows)}")
    print(f"Regression rows: {len(regression_rows)}")


if __name__ == "__main__":
    main()
