from __future__ import annotations

import re
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional

import xlwings as xw
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter


# User-editable paths
input_dir = "./input"
output_dir = "./output"


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


EMPIRICAL_DEFAULT_OFFSETS = {
    "num_quarters_used": -12,
    "last_quarter_used": -11,
    "forecast_value": -7,  # estimated total sold
    "actual_value": -6,  # reported sales
    "forecast_max": 0,
    "forecast_min": 1,
    "avg_penetration_pct": -3,
    "quarterly_sales": -5,
    "reported_sales": -6,
    "growth_rate_pct": -2,
    "sales_captured_in_db_pct": -1,
    "penetration_source": -1,
}

REGRESSION_DEFAULT_OFFSETS = {
    "num_quarters_used": -12,
    "forecast_value": -1,  # TOT FCST w/o SA is usually left of max
    "actual_value": -2,
    "forecast_max": 0,
    "forecast_min": 1,
}


HEADER_ALIASES_EMPIRICAL = {
    "num_quarters_used": ("num quarters used", "quarters used", "n quarters"),
    "last_quarter_used": ("last quarter used", "last quarter"),
    "forecast_value": ("estimated total sold", "tot fcst", "forecast"),
    "actual_value": ("actual value", "reported sales", "actual"),
    "forecast_max": ("max",),
    "forecast_min": ("min",),
    "avg_penetration_pct": ("avg penetration", "average penetration"),
    "quarterly_sales": ("quarterly sales",),
    "reported_sales": ("reported sales",),
    "growth_rate_pct": ("growth rate",),
    "sales_captured_in_db_pct": ("sales captured in db", "captured in db"),
    "penetration_source": ("penetration",),
}

HEADER_ALIASES_REGRESSION = {
    "num_quarters_used": ("num quarters used", "quarters used", "n quarters"),
    "forecast_value": ("tot fcst w/o sa", "tot fcst", "forecast"),
    "actual_value": ("actual value", "actual"),
    "forecast_max": ("max",),
    "forecast_min": ("min",),
}


DAY_BY_PHASE = {"early": 5, "mid": 15, "late": 25}
MONTH_LOOKUP = {
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


def normalize_header(value: Any) -> str:
    if value is None:
        return ""
    return re.sub(r"[^a-z0-9]+", " ", str(value).strip().lower()).strip()


def parse_float(value: Any) -> Optional[float]:
    if value is None or value == "":
        return None
    if isinstance(value, (int, float)):
        return float(value)
    text = str(value).strip().replace(",", "")
    if text.endswith("%"):
        text = text[:-1]
        try:
            return float(text) / 100.0
        except ValueError:
            return None
    try:
        return float(text)
    except ValueError:
        return None


def parse_int(value: Any) -> Optional[int]:
    as_float = parse_float(value)
    if as_float is None:
        return None
    return int(round(as_float))


def values_equal(left: Any, right: Any, tol: float = 1e-9) -> bool:
    lf = parse_float(left)
    rf = parse_float(right)
    if lf is not None and rf is not None:
        return abs(lf - rf) <= tol
    if left in (None, "") and right in (None, ""):
        return True
    return str(left).strip() == str(right).strip()


def coerce_2d(values: Any) -> List[List[Any]]:
    if values is None:
        return []
    if isinstance(values, tuple):
        values = list(values)
    if not isinstance(values, list):
        return [[values]]
    if not values:
        return []
    first = values[0]
    if isinstance(first, tuple):
        values = [list(row) for row in values]
    if values and not isinstance(values[0], list):
        return [values]
    return values


def safe_close_workbook(workbook: xw.Book) -> None:
    try:
        workbook.close(save=False)
        return
    except TypeError:
        pass
    except Exception:
        pass

    try:
        workbook.close(False)
        return
    except Exception:
        pass

    try:
        workbook.api.Close(SaveChanges=False)
        return
    except Exception:
        pass

    try:
        workbook.close()
    except Exception:
        # Last resort; workbook is intentionally left unsaved.
        pass


def next_output_path(input_path: Path, out_dir: Path) -> Path:
    base_name = f"{input_path.name}_PARAM"
    candidate = out_dir / f"{base_name}.xlsx"
    suffix = 1
    while candidate.exists():
        candidate = out_dir / f"{base_name}.{suffix}.xlsx"
        suffix += 1
    return candidate


def parse_file_metadata(file_path: Path) -> Dict[str, str]:
    stem = file_path.stem
    parts = [part.strip() for part in stem.split(" - ")]
    if len(parts) < 3:
        raise ValueError("filename does not match expected pattern '<prefix> - <ticker> - <period>...'")

    ticker = re.sub(r"[^A-Za-z0-9]", "", parts[1]).upper()
    if not ticker:
        raise ValueError("ticker token is empty")

    period_match = re.search(
        r"(Early|Mid|Late)\s*([A-Za-z]{3,9})\s*[_-]?\s*(\d{4})",
        stem,
        flags=re.IGNORECASE,
    )
    if not period_match:
        raise ValueError("period token is missing or malformed")

    phase_raw = period_match.group(1).lower()
    month_raw = period_match.group(2).lower()
    year = int(period_match.group(3))

    if month_raw not in MONTH_LOOKUP:
        raise ValueError(f"unrecognized month token '{month_raw}'")
    month_num = MONTH_LOOKUP[month_raw]
    month_abbrev = datetime(year, month_num, 1).strftime("%b")

    day = DAY_BY_PHASE[phase_raw]
    model_period = f"{phase_raw.title()}{month_abbrev}_{year}"
    model_date = datetime(year, month_num, day).strftime("%Y-%m-%d")
    model = f"{ticker}_{model_period}"
    return {
        "ticker": ticker,
        "model_period": model_period,
        "model_date": model_date,
        "model": model,
    }


def find_anchor(ws: xw.Sheet, token: str = "max") -> Optional[xw.Range]:
    try:
        found = ws.api.UsedRange.Find(What=token, MatchCase=False)
        if found is not None:
            return ws.range((int(found.Row), int(found.Column)))
    except Exception:
        pass

    used = ws.used_range
    grid = coerce_2d(used.value)
    if not grid:
        return None

    for r_idx, row_values in enumerate(grid, start=0):
        for c_idx, cell_value in enumerate(row_values, start=0):
            if normalize_header(cell_value) == token:
                return ws.range((used.row + r_idx, used.column + c_idx))
    return None


def read_header_row(
    ws: xw.Sheet,
    row: int,
    anchor_col: int,
    search_span: int = 30,
) -> Dict[str, int]:
    start_col = max(1, anchor_col - search_span)
    end_col = anchor_col + search_span
    values = ws.range((row, start_col), (row, end_col)).value
    if not isinstance(values, list):
        values = [values]

    headers: Dict[str, int] = {}
    for i, value in enumerate(values):
        key = normalize_header(value)
        if key:
            headers[key] = start_col + i
    return headers


def resolve_columns(
    header_columns: Dict[str, int],
    aliases: Dict[str, Iterable[str]],
    default_offsets: Dict[str, int],
    anchor_col: int,
) -> Dict[str, int]:
    resolved: Dict[str, int] = {}
    for field, tokens in aliases.items():
        selected_col = None
        for header_text, col in header_columns.items():
            if any(token in header_text for token in tokens):
                selected_col = col
                break
        if selected_col is not None:
            resolved[field] = selected_col
            continue

        offset = default_offsets.get(field)
        if offset is not None:
            resolved[field] = max(1, anchor_col + offset)
    return resolved


def read_cell(ws: xw.Sheet, row: int, col: Optional[int]) -> Any:
    if col is None or col < 1 or row < 1:
        return None
    try:
        return ws.range((row, col)).value
    except Exception:
        return None


def set_formula2_r1c1(cell: xw.Range, formula_r1c1: str) -> None:
    try:
        cell.formula2 = formula_r1c1
        return
    except Exception:
        pass
    try:
        cell.api.Formula2R1C1 = formula_r1c1
    except Exception:
        cell.api.FormulaR1C1 = formula_r1c1


def build_empirical_rows(
    workbook: xw.Book,
    metadata: Dict[str, str],
    source_file: str,
    n_quarters: int = 10,
) -> List[Dict[str, Any]]:
    sheet_name = "Empirical Model"
    if sheet_name not in [sheet.name for sheet in workbook.sheets]:
        print(f"skipped empirical extraction for {source_file}: missing sheet '{sheet_name}'")
        return []

    ws = workbook.sheets[sheet_name]
    anchor = find_anchor(ws, token="max")
    if anchor is None:
        print(f"skipped empirical extraction for {source_file}: 'max' anchor not found")
        return []

    anchor_row = anchor.row
    anchor_col = anchor.column
    start_row = anchor_row + 1

    headers = read_header_row(ws, anchor_row, anchor_col)
    cols = resolve_columns(headers, HEADER_ALIASES_EMPIRICAL, EMPIRICAL_DEFAULT_OFFSETS, anchor_col)

    penetration_source_col = cols.get("penetration_source")
    avg_pen_col = cols.get("avg_penetration_pct")

    temp_col = ws.used_range.last_cell.column + 3
    candidates: List[Dict[str, Any]] = []
    formulas_written = False

    for i in range(n_quarters):
        row = start_row + i
        num_quarters_used = parse_int(read_cell(ws, row, cols.get("num_quarters_used")))
        if num_quarters_used is None or num_quarters_used <= 0:
            num_quarters_used = i + 1

        candidate = {
            "row": row,
            "num_quarters_used": num_quarters_used,
            "last_quarter_used": read_cell(ws, row, cols.get("last_quarter_used")),
            "forecast_value": read_cell(ws, row, cols.get("forecast_value")),
            "actual_value": read_cell(ws, row, cols.get("actual_value")),
            "forecast_max": read_cell(ws, row, cols.get("forecast_max")),
            "forecast_min": read_cell(ws, row, cols.get("forecast_min")),
            "quarterly_sales": read_cell(ws, row, cols.get("quarterly_sales")),
            "reported_sales": read_cell(ws, row, cols.get("reported_sales")),
            "growth_rate_pct": read_cell(ws, row, cols.get("growth_rate_pct")),
            "sales_captured_in_db_pct": read_cell(ws, row, cols.get("sales_captured_in_db_pct")),
            "avg_penetration_pct": read_cell(ws, row, avg_pen_col),
            "temp_avg_col": temp_col,
        }
        candidates.append(candidate)

        if penetration_source_col is not None:
            formula_target = ws.range((row, temp_col))
            source_start_row = max(start_row, row - num_quarters_used + 1)
            source_end_row = row
            formula = (
                f"=AVERAGE("
                f"R{source_start_row}C{penetration_source_col}:R{source_end_row}C{penetration_source_col}"
                f")"
            )
            set_formula2_r1c1(formula_target, formula)
            formulas_written = True

    if formulas_written:
        workbook.app.calculate()

    output_rows: List[Dict[str, Any]] = []
    for candidate in candidates:
        row = candidate["row"]
        avg_pen = candidate["avg_penetration_pct"]
        if avg_pen in (None, "") and formulas_written:
            avg_pen = read_cell(ws, row, candidate["temp_avg_col"])

        forecast_max = candidate["forecast_max"]
        forecast_min = candidate["forecast_min"]
        max_num = parse_float(forecast_max)
        min_num = parse_float(forecast_min)
        range_width = (max_num - min_num) if (max_num is not None and min_num is not None) else None

        core_values = (
            candidate["forecast_value"],
            candidate["actual_value"],
            candidate["forecast_max"],
            candidate["forecast_min"],
            avg_pen,
        )
        if all(value in (None, "") for value in core_values):
            continue

        output_rows.append(
            {
                "model": metadata["model"],
                "ticker": metadata["ticker"],
                "model_period": metadata["model_period"],
                "model_date": metadata["model_date"],
                "method": "empirical",
                "parameter_name": "avg_penetration_pct",
                "parameter_value": avg_pen,
                "num_quarters_used": candidate["num_quarters_used"],
                "last_quarter_used": candidate["last_quarter_used"],
                "forecast_value": candidate["forecast_value"],
                "actual_value": candidate["actual_value"],
                "forecast_max": forecast_max,
                "forecast_min": forecast_min,
                "range_width": range_width,
                "avg_penetration_pct": avg_pen,
                "quarterly_sales": candidate["quarterly_sales"],
                "reported_sales": candidate["reported_sales"],
                "growth_rate_pct": candidate["growth_rate_pct"],
                "sales_captured_in_db_pct": candidate["sales_captured_in_db_pct"],
                "source_file": source_file,
            }
        )

    return output_rows


def build_regression_rows(
    workbook: xw.Book,
    metadata: Dict[str, str],
    source_file: str,
    n_quarters: int = 10,
) -> List[Dict[str, Any]]:
    sheet_name = "Regression Model"
    if sheet_name not in [sheet.name for sheet in workbook.sheets]:
        print(f"skipped regression extraction for {source_file}: missing sheet '{sheet_name}'")
        return []

    ws = workbook.sheets[sheet_name]
    anchor = find_anchor(ws, token="max")
    if anchor is None:
        print(f"skipped regression extraction for {source_file}: 'max' anchor not found")
        return []

    anchor_row = anchor.row
    anchor_col = anchor.column
    start_row = anchor_row + 1

    headers = read_header_row(ws, anchor_row, anchor_col)
    cols = resolve_columns(headers, HEADER_ALIASES_REGRESSION, REGRESSION_DEFAULT_OFFSETS, anchor_col)

    y_col = anchor_col - 7
    x_col = anchor_col - 11

    temp_intercept_col = ws.used_range.last_cell.column + 4
    temp_slope_col = temp_intercept_col + 1
    formulas_written = False

    candidates: List[Dict[str, Any]] = []
    for i in range(n_quarters):
        row = start_row + i
        num_quarters_used = parse_int(read_cell(ws, row, cols.get("num_quarters_used")))
        if num_quarters_used is None or num_quarters_used <= 0:
            num_quarters_used = i + 1

        history_end = anchor_row - 1
        history_start = max(1, history_end - num_quarters_used + 1)
        if history_end - history_start + 1 >= 2 and x_col >= 1 and y_col >= 1:
            intercept_formula = (
                f"=INTERCEPT(R{history_start}C{y_col}:R{history_end}C{y_col},"
                f"R{history_start}C{x_col}:R{history_end}C{x_col})"
            )
            slope_formula = (
                f"=SLOPE(R{history_start}C{y_col}:R{history_end}C{y_col},"
                f"R{history_start}C{x_col}:R{history_end}C{x_col})"
            )
            set_formula2_r1c1(ws.range((row, temp_intercept_col)), intercept_formula)
            set_formula2_r1c1(ws.range((row, temp_slope_col)), slope_formula)
            formulas_written = True

        candidates.append(
            {
                "row": row,
                "num_quarters_used": num_quarters_used,
                "forecast_value": read_cell(ws, row, cols.get("forecast_value")),
                "actual_value": read_cell(ws, row, cols.get("actual_value")),
                "forecast_max": read_cell(ws, row, cols.get("forecast_max")),
                "forecast_min": read_cell(ws, row, cols.get("forecast_min")),
            }
        )

    if formulas_written:
        workbook.app.calculate()

    output_rows: List[Dict[str, Any]] = []
    for candidate in candidates:
        row = candidate["row"]
        intercept_value = read_cell(ws, row, temp_intercept_col) if formulas_written else None
        slope_value = read_cell(ws, row, temp_slope_col) if formulas_written else None

        forecast_max = candidate["forecast_max"]
        forecast_min = candidate["forecast_min"]
        max_num = parse_float(forecast_max)
        min_num = parse_float(forecast_min)
        range_width = (max_num - min_num) if (max_num is not None and min_num is not None) else None

        core_values = (
            candidate["forecast_value"],
            candidate["forecast_max"],
            candidate["forecast_min"],
            intercept_value,
            slope_value,
        )
        if all(value in (None, "") for value in core_values):
            continue

        current_row = {
            "model": metadata["model"],
            "ticker": metadata["ticker"],
            "model_period": metadata["model_period"],
            "model_date": metadata["model_date"],
            "method": "regression",
            "parameter_name": "num_quarters_used",
            "parameter_value": candidate["num_quarters_used"],
            "num_quarters_used": candidate["num_quarters_used"],
            "forecast_value": candidate["forecast_value"],
            "actual_value": candidate["actual_value"],
            "forecast_max": forecast_max,
            "forecast_min": forecast_min,
            "range_width": range_width,
            "intercept": intercept_value,
            "slope": slope_value,
            "source_file": source_file,
        }

        if output_rows:
            previous = output_rows[-1]
            duplicate = (
                values_equal(previous.get("intercept"), current_row.get("intercept"))
                and values_equal(previous.get("slope"), current_row.get("slope"))
                and values_equal(previous.get("forecast_value"), current_row.get("forecast_value"))
                and values_equal(previous.get("forecast_max"), current_row.get("forecast_max"))
                and values_equal(previous.get("forecast_min"), current_row.get("forecast_min"))
            )
            if duplicate:
                continue

        output_rows.append(current_row)

    return output_rows


def write_sheet(ws, headers: List[str], rows: List[Dict[str, Any]]) -> None:
    ws.append(headers)
    for row in rows:
        ws.append([row.get(header) for header in headers])

    for cell in ws[1]:
        cell.font = Font(bold=True)

    ws.freeze_panes = "A2"
    if ws.max_row >= 1 and ws.max_column >= 1:
        ws.auto_filter.ref = ws.dimensions

    for col_idx, header in enumerate(headers, start=1):
        max_len = len(header)
        for r in range(2, ws.max_row + 1):
            value = ws.cell(r, col_idx).value
            if value is None:
                continue
            max_len = max(max_len, len(str(value)))
        ws.column_dimensions[get_column_letter(col_idx)].width = min(max_len + 2, 45)


def write_output_workbook(
    empirical_rows: List[Dict[str, Any]],
    regression_rows: List[Dict[str, Any]],
    output_path: Path,
) -> None:
    out_wb = Workbook()
    empirical_ws = out_wb.active
    empirical_ws.title = "empirical_candidates"
    regression_ws = out_wb.create_sheet("regression_candidates")

    write_sheet(empirical_ws, EMPIRICAL_COLUMNS, empirical_rows)
    write_sheet(regression_ws, REGRESSION_COLUMNS, regression_rows)

    out_wb.save(output_path)


def main() -> None:
    in_dir = Path(input_dir).expanduser().resolve()
    out_dir = Path(output_dir).expanduser().resolve()

    if not in_dir.exists() or not in_dir.is_dir():
        raise FileNotFoundError(f"input_dir does not exist or is not a directory: {in_dir}")
    out_dir.mkdir(parents=True, exist_ok=True)

    output_path = next_output_path(in_dir, out_dir)

    empirical_rows: List[Dict[str, Any]] = []
    regression_rows: List[Dict[str, Any]] = []
    processed_files = 0

    files = sorted(in_dir.iterdir(), key=lambda path: path.name.lower())

    with xw.App(visible=False, add_book=False) as app:
        app.display_alerts = False
        app.screen_updating = False
        app.calculation = "manual"

        for file_path in files:
            if not file_path.is_file():
                print(f"skipped file: {file_path.name} (not a file)")
                continue
            if file_path.name.startswith("~"):
                print(f"skipped file: {file_path.name} (temp file)")
                continue
            if file_path.suffix.lower() != ".xlsx":
                print(f"skipped file: {file_path.name} (not .xlsx)")
                continue

            try:
                metadata = parse_file_metadata(file_path)
            except Exception as exc:
                print(f"skipped file: {file_path.name} (metadata parse error: {exc})")
                continue

            workbook = None
            try:
                workbook = app.books.open(str(file_path), update_links=False)
                file_empirical = build_empirical_rows(workbook, metadata, file_path.name)
                file_regression = build_regression_rows(workbook, metadata, file_path.name)
                empirical_rows.extend(file_empirical)
                regression_rows.extend(file_regression)
                processed_files += 1
                print(f"processed file: {file_path.name}")
            except Exception as exc:
                print(f"skipped file: {file_path.name} (processing error: {exc})")
            finally:
                if workbook is not None:
                    safe_close_workbook(workbook)

    write_output_workbook(empirical_rows, regression_rows, output_path)

    print(f"output path: {output_path}")
    print(f"number of files processed: {processed_files}")
    print(f"number of empirical rows: {len(empirical_rows)}")
    print(f"number of regression rows: {len(regression_rows)}")


if __name__ == "__main__":
    main()
