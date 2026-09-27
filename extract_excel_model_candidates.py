from __future__ import annotations

import math
import re
from datetime import date, datetime
from pathlib import Path
from typing import Any, Dict, List, Optional, Sequence, Tuple

import xlwings as xw
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

# ========= user-configurable paths =========
input_dir = Path("/path/to/input")
output_dir = Path("/path/to/output")
# ==========================================

N_QUARTERS = 10
EMPIRICAL_MODEL_SHEET = "Empirical Model"
REGRESSION_MODEL_SHEET = "Regression Model"

# Anchor-based column offsets for the empirical model (relative to the "max" anchor column)
EMPIRICAL_OFFSETS = {
    "quarter_label": -12,
    "quarterly_sales": -11,
    "reported_sales": -10,
    "growth_rate_pct": -9,
    "sales_captured_in_db_pct": -8,
}

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


def safe_float(value: Any) -> Optional[float]:
    if value is None or value == "":
        return None
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        if isinstance(value, float) and (math.isnan(value) or math.isinf(value)):
            return None
        return float(value)
    if isinstance(value, str):
        cleaned = value.strip().replace(",", "")
        if not cleaned:
            return None
        if cleaned.endswith("%"):
            try:
                return float(cleaned[:-1]) / 100.0
            except ValueError:
                return None
        try:
            return float(cleaned)
        except ValueError:
            return None
    return None


def normalize_pct(value: Any) -> Optional[float]:
    num = safe_float(value)
    if num is None:
        return None
    if abs(num) > 1 and abs(num) <= 100:
        return num / 100.0
    return num


def format_scalar(value: Any) -> Any:
    if value is None:
        return ""
    if isinstance(value, datetime):
        return value.date().isoformat()
    if isinstance(value, date):
        return value.isoformat()
    if isinstance(value, float) and (math.isnan(value) or math.isinf(value)):
        return ""
    return value


def subtract_or_none(left: Optional[float], right: Optional[float]) -> Optional[float]:
    if left is None or right is None:
        return None
    return left - right


def r1c1_ref(cell: xw.main.Range) -> str:
    return f"R{cell.row}C{cell.column}"


def set_formula2_r1c1(cell: xw.main.Range, formula: str) -> None:
    api = cell.api
    if hasattr(api, "Formula2R1C1"):
        api.Formula2R1C1 = formula
    elif hasattr(api, "FormulaR1C1"):
        api.FormulaR1C1 = formula
    else:
        # Last fallback if Formula2R1C1 is not supported.
        cell.formula2 = formula


def find_anchor_cell(sheet: xw.main.Sheet, anchor_text: str = "max") -> xw.main.Range:
    used_api = sheet.api.UsedRange
    for probe in (anchor_text, anchor_text.upper(), anchor_text.title()):
        found = used_api.Find(What=probe, LookAt=1, MatchCase=False)
        if found is not None:
            return sheet.range((found.Row, found.Column))

    # Fallback for engines where Find behaves differently.
    used = sheet.used_range
    values = used.value
    if values is None:
        raise ValueError(f"Anchor '{anchor_text}' not found (empty used range).")
    if not isinstance(values, list):
        values = [[values]]
    elif values and not isinstance(values[0], list):
        values = [values]

    wanted = anchor_text.strip().lower()
    for r_idx, row_vals in enumerate(values):
        for c_idx, raw in enumerate(row_vals):
            if isinstance(raw, str) and raw.strip().lower() == wanted:
                return sheet.range((used.row + r_idx, used.column + c_idx))
    raise ValueError(f"Anchor '{anchor_text}' not found on sheet '{sheet.name}'.")


def collect_recent_numeric_rows(
    sheet: xw.main.Sheet,
    anchor_row: int,
    required_cols: Sequence[int],
    max_count: int,
) -> List[int]:
    if anchor_row <= 1:
        return []
    if not required_cols:
        return []

    # Read a single block once to avoid repeated COM calls.
    lookback = 500
    start_row = max(1, anchor_row - lookback)
    min_col = min(required_cols)
    max_col = max(required_cols)
    block = sheet.range((start_row, min_col), (anchor_row - 1, max_col)).value

    if block is None:
        return []
    if not isinstance(block, list):
        block = [[block]]
    elif block and not isinstance(block[0], list):
        block = [block]

    rows: List[int] = []
    for idx in range(len(block) - 1, -1, -1):
        row_vals = block[idx]
        is_valid = True
        for col in required_cols:
            raw = row_vals[col - min_col]
            if safe_float(raw) is None:
                is_valid = False
                break
        if is_valid:
            rows.append(start_row + idx)
            if len(rows) >= max_count:
                break
    rows.sort()
    return rows


def parse_file_label(file_name: str) -> Optional[Dict[str, str]]:
    stem = Path(file_name).stem
    pieces = [part.strip() for part in stem.split("-")]
    if len(pieces) < 3:
        return None

    ticker_raw = re.sub(r"[^A-Za-z0-9]", "", pieces[1]).upper()
    if not ticker_raw:
        return None

    period_match = re.search(r"(Early|Mid|Late)\s*([A-Za-z]{3,9})\s*(\d{4})", stem, flags=re.IGNORECASE)
    if not period_match:
        return None

    period_slot = period_match.group(1).title()
    month_token = period_match.group(2).title()
    year_token = period_match.group(3)

    month_num = parse_month(month_token)
    if month_num is None:
        return None

    day_lookup = {"Early": 5, "Mid": 15, "Late": 25}
    model_day = day_lookup[period_slot]
    model_date = date(int(year_token), month_num, model_day).isoformat()

    short_month = datetime(2000, month_num, 1).strftime("%b")
    model_period = f"{period_slot}{short_month}_{year_token}"
    model = f"{ticker_raw}_{model_period}"

    return {
        "model": model,
        "ticker": ticker_raw,
        "model_period": model_period,
        "model_date": model_date,
    }


def parse_month(token: str) -> Optional[int]:
    for fmt in ("%b", "%B"):
        try:
            return datetime.strptime(token[:3] if fmt == "%b" else token, fmt).month
        except ValueError:
            continue
    return None


def close_source_workbook(wb: xw.main.Book) -> None:
    # Preferred safe close path.
    try:
        wb.close(save=False)
        return
    except TypeError:
        pass
    except Exception:
        pass

    # COM fallback for environments where xlwings close signature differs.
    try:
        wb.api.Close(SaveChanges=False)
        return
    except Exception:
        pass

    # Last fallback (alerts are disabled at app level).
    try:
        wb.close()
    except Exception:
        pass


def extract_empirical_rows(
    wb: xw.main.Book,
    meta: Dict[str, str],
    source_file: str,
) -> List[Dict[str, Any]]:
    try:
        sheet = wb.sheets[EMPIRICAL_MODEL_SHEET]
    except Exception:
        return []

    anchor = find_anchor_cell(sheet, "max")
    anchor_row, anchor_col = anchor.row, anchor.column

    quarter_col = anchor_col + EMPIRICAL_OFFSETS["quarter_label"]
    quarterly_sales_col = anchor_col + EMPIRICAL_OFFSETS["quarterly_sales"]
    reported_sales_col = anchor_col + EMPIRICAL_OFFSETS["reported_sales"]
    growth_col = anchor_col + EMPIRICAL_OFFSETS["growth_rate_pct"]
    captured_col = anchor_col + EMPIRICAL_OFFSETS["sales_captured_in_db_pct"]

    if min(quarter_col, quarterly_sales_col, reported_sales_col, growth_col, captured_col) < 1:
        return []

    recent_rows = collect_recent_numeric_rows(
        sheet=sheet,
        anchor_row=anchor_row,
        required_cols=[quarterly_sales_col, captured_col],
        max_count=N_QUARTERS,
    )
    if not recent_rows:
        return []

    temp_col = anchor_col + 2
    if temp_col < 1:
        temp_col = anchor_col + 1
    temp_cells = {
        "avg_pen": sheet.range((anchor_row, temp_col)),
        "max_pen": sheet.range((anchor_row, temp_col + 1)),
        "min_pen": sheet.range((anchor_row, temp_col + 2)),
        "forecast": sheet.range((anchor_row, temp_col + 3)),
        "forecast_max": sheet.range((anchor_row, temp_col + 4)),
        "forecast_min": sheet.range((anchor_row, temp_col + 5)),
    }

    rows: List[Dict[str, Any]] = []
    try:
        max_anchor_value = safe_float(anchor.value)
        min_anchor_value = safe_float(sheet.range((anchor_row + 1, anchor_col)).value)

        loop_count = min(N_QUARTERS, len(recent_rows))
        for n_used in range(1, loop_count + 1):
            sample_rows = recent_rows[-n_used:]
            start_row = sample_rows[0]
            end_row = sample_rows[-1]

            set_formula2_r1c1(temp_cells["avg_pen"], f"=AVERAGE(R{start_row}C{captured_col}:R{end_row}C{captured_col})")
            set_formula2_r1c1(temp_cells["max_pen"], f"=MAX(R{start_row}C{captured_col}:R{end_row}C{captured_col})")
            set_formula2_r1c1(temp_cells["min_pen"], f"=MIN(R{start_row}C{captured_col}:R{end_row}C{captured_col})")
            set_formula2_r1c1(
                temp_cells["forecast"],
                f"=IFERROR(R{end_row}C{quarterly_sales_col}/{r1c1_ref(temp_cells['avg_pen'])},\"\")",
            )
            set_formula2_r1c1(
                temp_cells["forecast_max"],
                f"=IFERROR(R{end_row}C{quarterly_sales_col}/{r1c1_ref(temp_cells['min_pen'])},\"\")",
            )
            set_formula2_r1c1(
                temp_cells["forecast_min"],
                f"=IFERROR(R{end_row}C{quarterly_sales_col}/{r1c1_ref(temp_cells['max_pen'])},\"\")",
            )

            wb.app.calculate()

            avg_pen = normalize_pct(temp_cells["avg_pen"].value)
            forecast_value = safe_float(temp_cells["forecast"].value)
            forecast_max = safe_float(temp_cells["forecast_max"].value)
            forecast_min = safe_float(temp_cells["forecast_min"].value)

            if forecast_max is None:
                forecast_max = max_anchor_value
            if forecast_min is None:
                forecast_min = min_anchor_value
            if forecast_max is not None and forecast_min is not None and forecast_max < forecast_min:
                forecast_max, forecast_min = forecast_min, forecast_max

            quarterly_sales = safe_float(sheet.range((end_row, quarterly_sales_col)).value)
            reported_sales = safe_float(sheet.range((end_row, reported_sales_col)).value)
            growth_rate = normalize_pct(sheet.range((end_row, growth_col)).value)
            sales_captured = normalize_pct(sheet.range((end_row, captured_col)).value)
            quarter_label = format_scalar(sheet.range((end_row, quarter_col)).value)

            rows.append(
                {
                    "model": meta["model"],
                    "ticker": meta["ticker"],
                    "model_period": meta["model_period"],
                    "model_date": meta["model_date"],
                    "method": "empirical",
                    "parameter_name": "avg_penetration_pct",
                    "parameter_value": avg_pen,
                    "num_quarters_used": n_used,
                    "last_quarter_used": quarter_label,
                    "forecast_value": forecast_value,
                    "actual_value": reported_sales,
                    "forecast_max": forecast_max,
                    "forecast_min": forecast_min,
                    "range_width": subtract_or_none(forecast_max, forecast_min),
                    "avg_penetration_pct": avg_pen,
                    "quarterly_sales": quarterly_sales,
                    "reported_sales": reported_sales,
                    "growth_rate_pct": growth_rate,
                    "sales_captured_in_db_pct": sales_captured,
                    "source_file": source_file,
                }
            )
    finally:
        for cell in temp_cells.values():
            cell.value = None

    return rows


def extract_regression_rows(
    wb: xw.main.Book,
    meta: Dict[str, str],
    source_file: str,
) -> List[Dict[str, Any]]:
    try:
        sheet = wb.sheets[REGRESSION_MODEL_SHEET]
    except Exception:
        return []

    anchor = find_anchor_cell(sheet, "max")
    anchor_row, anchor_col = anchor.row, anchor.column

    y_col = anchor_col - 7
    x_col = anchor_col - 11
    if min(y_col, x_col) < 1:
        return []

    recent_rows = collect_recent_numeric_rows(
        sheet=sheet,
        anchor_row=anchor_row,
        required_cols=[x_col, y_col],
        max_count=N_QUARTERS,
    )
    if len(recent_rows) < 2:
        return []

    temp_col = anchor_col + 2
    if temp_col < 1:
        temp_col = anchor_col + 1
    temp_cells = {
        "intercept": sheet.range((anchor_row, temp_col)),
        "slope": sheet.range((anchor_row, temp_col + 1)),
        "forecast": sheet.range((anchor_row, temp_col + 2)),
        "std_dev": sheet.range((anchor_row, temp_col + 3)),
        "forecast_max": sheet.range((anchor_row, temp_col + 4)),
        "forecast_min": sheet.range((anchor_row, temp_col + 5)),
    }

    rows: List[Dict[str, Any]] = []
    previous_signature: Optional[Tuple[Optional[float], ...]] = None
    try:
        max_anchor_value = safe_float(anchor.value)
        min_anchor_value = safe_float(sheet.range((anchor_row + 1, anchor_col)).value)

        loop_count = min(N_QUARTERS, len(recent_rows))
        for n_used in range(2, loop_count + 1):
            sample_rows = recent_rows[-n_used:]
            start_row = sample_rows[0]
            end_row = sample_rows[-1]

            set_formula2_r1c1(
                temp_cells["intercept"],
                f"=INTERCEPT(R{start_row}C{y_col}:R{end_row}C{y_col},R{start_row}C{x_col}:R{end_row}C{x_col})",
            )
            set_formula2_r1c1(
                temp_cells["slope"],
                f"=SLOPE(R{start_row}C{y_col}:R{end_row}C{y_col},R{start_row}C{x_col}:R{end_row}C{x_col})",
            )
            set_formula2_r1c1(
                temp_cells["forecast"],
                f"={r1c1_ref(temp_cells['intercept'])}+{r1c1_ref(temp_cells['slope'])}*(R{end_row}C{x_col}+1)",
            )
            set_formula2_r1c1(temp_cells["std_dev"], f"=STDEV.S(R{start_row}C{y_col}:R{end_row}C{y_col})")
            set_formula2_r1c1(
                temp_cells["forecast_max"],
                f"=IFERROR({r1c1_ref(temp_cells['forecast'])}+{r1c1_ref(temp_cells['std_dev'])},\"\")",
            )
            set_formula2_r1c1(
                temp_cells["forecast_min"],
                f"=IFERROR({r1c1_ref(temp_cells['forecast'])}-{r1c1_ref(temp_cells['std_dev'])},\"\")",
            )

            wb.app.calculate()

            intercept = safe_float(temp_cells["intercept"].value)
            slope = safe_float(temp_cells["slope"].value)
            forecast_value = safe_float(temp_cells["forecast"].value)
            forecast_max = safe_float(temp_cells["forecast_max"].value)
            forecast_min = safe_float(temp_cells["forecast_min"].value)

            if forecast_max is None:
                forecast_max = max_anchor_value
            if forecast_min is None:
                forecast_min = min_anchor_value
            if forecast_max is not None and forecast_min is not None and forecast_max < forecast_min:
                forecast_max, forecast_min = forecast_min, forecast_max

            signature = tuple(
                None if value is None else round(value, 10)
                for value in (forecast_value, forecast_max, forecast_min, intercept, slope)
            )
            if signature == previous_signature:
                continue
            previous_signature = signature

            rows.append(
                {
                    "model": meta["model"],
                    "ticker": meta["ticker"],
                    "model_period": meta["model_period"],
                    "model_date": meta["model_date"],
                    "method": "regression",
                    "parameter_name": "num_quarters_used",
                    "parameter_value": n_used,
                    "num_quarters_used": n_used,
                    "forecast_value": forecast_value,
                    "actual_value": "",
                    "forecast_max": forecast_max,
                    "forecast_min": forecast_min,
                    "range_width": subtract_or_none(forecast_max, forecast_min),
                    "intercept": intercept,
                    "slope": slope,
                    "source_file": source_file,
                }
            )
    finally:
        for cell in temp_cells.values():
            cell.value = None

    return rows


def unique_output_path(in_dir: Path, out_dir: Path) -> Path:
    base_name = f"{in_dir.name}_PARAM"
    candidate = out_dir / f"{base_name}.xlsx"
    if not candidate.exists():
        return candidate
    idx = 1
    while True:
        candidate = out_dir / f"{base_name}.{idx}.xlsx"
        if not candidate.exists():
            return candidate
        idx += 1


def write_output_workbook(
    destination: Path,
    empirical_rows: List[Dict[str, Any]],
    regression_rows: List[Dict[str, Any]],
) -> None:
    wb = Workbook()
    default_sheet = wb.active
    wb.remove(default_sheet)

    empirical_ws = wb.create_sheet("empirical_candidates")
    regression_ws = wb.create_sheet("regression_candidates")

    write_sheet(empirical_ws, EMPIRICAL_COLUMNS, empirical_rows)
    write_sheet(regression_ws, REGRESSION_COLUMNS, regression_rows)

    wb.save(destination)


def write_sheet(ws, columns: Sequence[str], rows: Sequence[Dict[str, Any]]) -> None:
    for col_idx, col_name in enumerate(columns, start=1):
        ws.cell(row=1, column=col_idx, value=col_name)

    for row_idx, row_data in enumerate(rows, start=2):
        for col_idx, col_name in enumerate(columns, start=1):
            ws.cell(row=row_idx, column=col_idx, value=format_scalar(row_data.get(col_name)))

    for header_cell in ws[1]:
        header_cell.font = Font(bold=True)

    ws.freeze_panes = "A2"
    last_col = get_column_letter(len(columns))
    ws.auto_filter.ref = f"A1:{last_col}{max(1, ws.max_row)}"

    for col_idx, col_name in enumerate(columns, start=1):
        letter = get_column_letter(col_idx)
        max_len = len(col_name)
        for row_idx in range(2, ws.max_row + 1):
            value = ws.cell(row=row_idx, column=col_idx).value
            if value is None:
                continue
            max_len = max(max_len, len(str(value)))
        ws.column_dimensions[letter].width = min(max(max_len + 2, 12), 48)


def process_workbooks() -> None:
    if not input_dir.exists():
        raise FileNotFoundError(f"input_dir does not exist: {input_dir}")
    output_dir.mkdir(parents=True, exist_ok=True)

    all_entries = sorted(input_dir.iterdir(), key=lambda p: p.name.lower())
    output_path = unique_output_path(input_dir, output_dir)

    empirical_rows: List[Dict[str, Any]] = []
    regression_rows: List[Dict[str, Any]] = []
    processed_file_count = 0

    app = xw.App(visible=False, add_book=False)
    app.display_alerts = False
    app.screen_updating = False
    try:
        try:
            app.calculation = "manual"
        except Exception:
            pass

        for path in all_entries:
            if not path.is_file():
                print(f"Skipped {path.name}: not a file")
                continue
            if path.name.startswith("~"):
                print(f"Skipped {path.name}: temporary file")
                continue
            if path.suffix.lower() != ".xlsx":
                print(f"Skipped {path.name}: not an .xlsx file")
                continue

            meta = parse_file_label(path.name)
            if meta is None:
                print(f"Skipped {path.name}: filename does not match expected model pattern")
                continue

            wb: Optional[xw.main.Book] = None
            try:
                wb = app.books.open(str(path), update_links=False)
                file_empirical = extract_empirical_rows(wb, meta, path.name)
                file_regression = extract_regression_rows(wb, meta, path.name)
                empirical_rows.extend(file_empirical)
                regression_rows.extend(file_regression)
                processed_file_count += 1
                print(f"Processed {path.name}")
            except Exception as exc:
                print(f"Skipped {path.name}: processing error ({exc})")
            finally:
                if wb is not None:
                    close_source_workbook(wb)
    finally:
        try:
            app.quit()
        except Exception:
            pass

    write_output_workbook(output_path, empirical_rows, regression_rows)

    print(f"Output path: {output_path}")
    print(f"Files processed: {processed_file_count}")
    print(f"Empirical rows: {len(empirical_rows)}")
    print(f"Regression rows: {len(regression_rows)}")


if __name__ == "__main__":
    process_workbooks()
