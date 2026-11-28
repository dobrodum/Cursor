#!/usr/bin/env python3
from __future__ import annotations

import calendar
import math
import re
from dataclasses import dataclass
from datetime import date
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

import xlwings as xw
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

# --- user-configurable paths ---
input_dir = Path("./input")
output_dir = Path("./output")

MAX_QUARTERS = 10

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

_NON_ALNUM_RE = re.compile(r"[^a-z0-9]+")

HEADER_ALIASES = {
    "quarter_label": {
        "quarter",
        "quarterlabel",
        "lastquarter",
        "period",
        "fiscalquarter",
        "quarterused",
    },
    "quarterly_sales": {
        "quarterlysales",
        "qtrsales",
        "quartersales",
        "salesquarterly",
    },
    "reported_sales": {
        "reportedsales",
        "actualsales",
        "salesreported",
    },
    "growth_rate_pct": {
        "growthrate",
        "growthratepct",
        "growthpct",
    },
    "sales_captured_in_db_pct": {
        "salescapturedindb",
        "salescapturedindbpct",
        "capturedindb",
        "dbcapturepct",
        "penetrationpct",
    },
}


@dataclass
class SheetSnapshot:
    base_row: int
    base_col: int
    values: List[List[Any]]

    @property
    def row_count(self) -> int:
        return len(self.values)

    @property
    def col_count(self) -> int:
        if not self.values:
            return 0
        return len(self.values[0])

    def get(self, abs_row: int, abs_col: int) -> Any:
        row_idx = abs_row - self.base_row
        col_idx = abs_col - self.base_col
        if row_idx < 0 or col_idx < 0:
            return None
        if row_idx >= self.row_count or col_idx >= self.col_count:
            return None
        return self.values[row_idx][col_idx]


def normalize_label(value: Any) -> str:
    if value is None:
        return ""
    return _NON_ALNUM_RE.sub("", str(value).strip().lower())


def ensure_2d(values: Any) -> List[List[Any]]:
    if values is None:
        return []
    if isinstance(values, (list, tuple)):
        if len(values) == 0:
            return []
        first = values[0]
        if isinstance(first, (list, tuple)):
            return [list(row) for row in values]
        return [list(values)]
    return [[values]]


def to_float(value: Any) -> Optional[float]:
    if value is None:
        return None
    if isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        val = float(value)
        if math.isnan(val):
            return None
        return val
    if isinstance(value, str):
        text = value.strip()
        if text == "":
            return None
        is_pct = text.endswith("%")
        text = text.replace("%", "").replace(",", "")
        try:
            number = float(text)
        except ValueError:
            return None
        if is_pct:
            number /= 100.0
        return number
    return None


def parse_month(month_token: str) -> Optional[int]:
    clean = re.sub(r"[^A-Za-z]", "", month_token).lower()
    if not clean:
        return None
    month_lookup: Dict[str, int] = {}
    for i in range(1, 13):
        month_lookup[calendar.month_abbr[i].lower()] = i
        month_lookup[calendar.month_name[i].lower()] = i
    if clean in month_lookup:
        return month_lookup[clean]
    if clean[:3] in month_lookup:
        return month_lookup[clean[:3]]
    return None


def parse_file_metadata(file_name: str) -> Dict[str, str]:
    stem = Path(file_name).stem
    parts = [part.strip() for part in stem.split(" - ") if part.strip()]
    ticker = parts[1] if len(parts) >= 2 else "UNKNOWN"

    period_match = re.search(r"(Early|Mid|Late)([A-Za-z]+)(\d{4})", stem, re.IGNORECASE)
    model_period = "UNKNOWN_PERIOD"
    model_date = ""

    if period_match:
        timing = period_match.group(1).title()
        month_token = period_match.group(2)
        year = int(period_match.group(3))
        month_num = parse_month(month_token)
        if month_num is not None:
            month_abbr = calendar.month_abbr[month_num]
            model_period = f"{timing}{month_abbr}_{year}"
            day_lookup = {"Early": 5, "Mid": 15, "Late": 25}
            model_date = date(year, month_num, day_lookup[timing]).isoformat()
    model = f"{ticker}_{model_period}"
    return {
        "model": model,
        "ticker": ticker,
        "model_period": model_period,
        "model_date": model_date,
    }


def make_sheet_snapshot(sheet: xw.Sheet) -> SheetSnapshot:
    used = sheet.used_range
    values = ensure_2d(used.value)
    if not values:
        return SheetSnapshot(base_row=1, base_col=1, values=[])
    max_cols = max(len(row) for row in values)
    for row in values:
        if len(row) < max_cols:
            row.extend([None] * (max_cols - len(row)))
    return SheetSnapshot(base_row=used.row, base_col=used.column, values=values)


def find_max_anchor(snapshot: SheetSnapshot) -> Tuple[int, int]:
    candidates: List[Tuple[int, int, int]] = []
    for r_idx, row in enumerate(snapshot.values):
        for c_idx, value in enumerate(row):
            if normalize_label(value) != "max":
                continue
            abs_row = snapshot.base_row + r_idx
            abs_col = snapshot.base_col + c_idx
            score = abs_row * 10_000 + abs_col
            if normalize_label(snapshot.get(abs_row + 1, abs_col)) == "min":
                score += 1_000_000_000
            if normalize_label(snapshot.get(abs_row, abs_col + 1)) == "min":
                score += 500_000_000
            candidates.append((score, abs_row, abs_col))
    if not candidates:
        raise ValueError("could not find a 'max' anchor cell")
    candidates.sort(reverse=True)
    _, anchor_row, anchor_col = candidates[0]
    return anchor_row, anchor_col


def read_anchor_max_min(snapshot: SheetSnapshot, anchor_row: int, anchor_col: int) -> Tuple[Optional[float], Optional[float]]:
    max_value = None
    min_value = None

    if normalize_label(snapshot.get(anchor_row + 1, anchor_col)) == "min":
        max_value = to_float(snapshot.get(anchor_row, anchor_col + 1))
        min_value = to_float(snapshot.get(anchor_row + 1, anchor_col + 1))
        return max_value, min_value

    if normalize_label(snapshot.get(anchor_row, anchor_col + 2)) == "min":
        max_value = to_float(snapshot.get(anchor_row, anchor_col + 1))
        min_value = to_float(snapshot.get(anchor_row, anchor_col + 3))
        return max_value, min_value

    max_value = to_float(snapshot.get(anchor_row, anchor_col + 1))
    min_value = to_float(snapshot.get(anchor_row + 1, anchor_col + 1))
    return max_value, min_value


def resolve_header_key(normalized_text: str) -> Optional[str]:
    for key, aliases in HEADER_ALIASES.items():
        if normalized_text in aliases:
            return key
    return None


def find_empirical_header_map(snapshot: SheetSnapshot, anchor_row: int, anchor_col: int) -> Tuple[int, Dict[str, int]]:
    if snapshot.row_count == 0 or snapshot.col_count == 0:
        raise ValueError("sheet appears empty")

    first_row = max(snapshot.base_row, anchor_row - 120)
    last_row = min(snapshot.base_row + snapshot.row_count - 1, anchor_row + 10)
    first_col = max(snapshot.base_col, anchor_col - 35)
    last_col = min(snapshot.base_col + snapshot.col_count - 1, anchor_col + 8)

    best_row = None
    best_score = (-1, -1)
    best_map: Dict[str, int] = {}

    core_headers = ("quarterly_sales", "reported_sales", "sales_captured_in_db_pct")
    for row in range(first_row, last_row + 1):
        row_map: Dict[str, int] = {}
        for col in range(first_col, last_col + 1):
            key = resolve_header_key(normalize_label(snapshot.get(row, col)))
            if key and key not in row_map:
                row_map[key] = col
        score = sum(1 for key in core_headers if key in row_map)
        if score == 0:
            continue
        rank = (score, row)
        if rank > best_score:
            best_score = rank
            best_row = row
            best_map = row_map

    if best_row is None:
        best_row = max(snapshot.base_row, anchor_row - MAX_QUARTERS - 2)
        best_map = {}

    fallback_offsets = {
        "quarter_label": anchor_col - 12,
        "quarterly_sales": anchor_col - 11,
        "reported_sales": anchor_col - 10,
        "growth_rate_pct": anchor_col - 9,
        "sales_captured_in_db_pct": anchor_col - 8,
    }
    for key, fallback_col in fallback_offsets.items():
        if key not in best_map:
            best_map[key] = fallback_col

    return best_row, best_map


def collect_empirical_history(
    snapshot: SheetSnapshot,
    header_row: int,
    anchor_row: int,
    col_map: Dict[str, int],
) -> List[Dict[str, Any]]:
    rows: List[Dict[str, Any]] = []
    start_row = max(snapshot.base_row, header_row + 1)
    end_row = min(anchor_row - 1, snapshot.base_row + snapshot.row_count - 1)
    empty_streak = 0

    for row in range(start_row, end_row + 1):
        quarter_label = snapshot.get(row, col_map["quarter_label"])
        quarterly_sales = to_float(snapshot.get(row, col_map["quarterly_sales"]))
        reported_sales = to_float(snapshot.get(row, col_map["reported_sales"]))
        growth_rate_pct = to_float(snapshot.get(row, col_map["growth_rate_pct"]))
        sales_captured = to_float(snapshot.get(row, col_map["sales_captured_in_db_pct"]))

        if all(value is None for value in (quarterly_sales, reported_sales, growth_rate_pct, sales_captured)):
            if rows:
                empty_streak += 1
                if empty_streak >= 3:
                    break
            continue

        empty_streak = 0
        rows.append(
            {
                "row": row,
                "quarter_label": quarter_label,
                "quarterly_sales": quarterly_sales,
                "reported_sales": reported_sales,
                "growth_rate_pct": growth_rate_pct,
                "sales_captured_in_db_pct": sales_captured,
            }
        )
    return rows


def close_source_workbook(wb: xw.Book) -> None:
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


def get_sheet_case_insensitive(wb: xw.Book, name: str) -> Optional[xw.Sheet]:
    target = name.strip().lower()
    for sheet in wb.sheets:
        if sheet.name.strip().lower() == target:
            return sheet
    return None


def calc_range_width(max_value: Optional[float], min_value: Optional[float]) -> Optional[float]:
    if max_value is None or min_value is None:
        return None
    return max_value - min_value


def write_formula2(cell: xw.Range, formula: str) -> None:
    # formula2 keeps dynamic-array support and allows R1C1 references directly.
    cell.formula2 = formula


def extract_empirical_candidates(sheet: xw.Sheet, metadata: Dict[str, str], source_file: str) -> List[Dict[str, Any]]:
    snapshot = make_sheet_snapshot(sheet)
    anchor_row, anchor_col = find_max_anchor(snapshot)
    anchor_max, anchor_min = read_anchor_max_min(snapshot, anchor_row, anchor_col)
    header_row, col_map = find_empirical_header_map(snapshot, anchor_row, anchor_col)
    history = collect_empirical_history(snapshot, header_row, anchor_row, col_map)

    usable_history = [
        row
        for row in history
        if row["quarterly_sales"] is not None and row["sales_captured_in_db_pct"] is not None
    ]
    if not usable_history:
        return []

    max_n = min(MAX_QUARTERS, len(usable_history))
    temp_row = max(anchor_row + 2, snapshot.base_row + snapshot.row_count + 2)
    temp_col = max(anchor_col + 2, snapshot.base_col + snapshot.col_count + 2)

    avg_cell = sheet.range((temp_row, temp_col))
    forecast_cell = sheet.range((temp_row, temp_col + 1))
    max_cell = sheet.range((temp_row, temp_col + 2))
    min_cell = sheet.range((temp_row, temp_col + 3))

    output_rows: List[Dict[str, Any]] = []
    for n_quarters in range(1, max_n + 1):
        segment = usable_history[-n_quarters:]
        start_row = segment[0]["row"]
        end_row = segment[-1]["row"]
        latest = segment[-1]
        latest_row = latest["row"]

        sales_capture_col = col_map["sales_captured_in_db_pct"]
        quarterly_sales_col = col_map["quarterly_sales"]
        reported_sales_col = col_map["reported_sales"]

        write_formula2(
            avg_cell,
            f"=AVERAGE(R{start_row}C{sales_capture_col}:R{end_row}C{sales_capture_col})",
        )
        write_formula2(
            forecast_cell,
            f"=IFERROR(R{latest_row}C{quarterly_sales_col}/R{temp_row}C{temp_col},\"\")",
        )
        if latest["reported_sales"] is not None:
            write_formula2(max_cell, f"=MAX(R{temp_row}C{temp_col + 1},R{latest_row}C{reported_sales_col})")
            write_formula2(min_cell, f"=MIN(R{temp_row}C{temp_col + 1},R{latest_row}C{reported_sales_col})")
        else:
            write_formula2(max_cell, f"=R{temp_row}C{temp_col + 1}")
            write_formula2(min_cell, f"=R{temp_row}C{temp_col + 1}")

        sheet.book.app.calculate()

        avg_penetration = to_float(avg_cell.value)
        forecast_value = to_float(forecast_cell.value)
        forecast_max = to_float(max_cell.value)
        forecast_min = to_float(min_cell.value)
        if forecast_max is None:
            forecast_max = anchor_max
        if forecast_min is None:
            forecast_min = anchor_min

        if forecast_value is None and avg_penetration and latest["quarterly_sales"] is not None and avg_penetration != 0:
            forecast_value = latest["quarterly_sales"] / avg_penetration

        range_width = calc_range_width(forecast_max, forecast_min)
        last_quarter_used = latest["quarter_label"] if latest["quarter_label"] not in (None, "") else latest_row

        output_rows.append(
            {
                "model": metadata["model"],
                "ticker": metadata["ticker"],
                "model_period": metadata["model_period"],
                "model_date": metadata["model_date"],
                "method": "empirical",
                "parameter_name": "avg_penetration_pct",
                "parameter_value": avg_penetration,
                "num_quarters_used": n_quarters,
                "last_quarter_used": last_quarter_used,
                "forecast_value": forecast_value,
                "actual_value": latest["reported_sales"],
                "forecast_max": forecast_max,
                "forecast_min": forecast_min,
                "range_width": range_width,
                "avg_penetration_pct": avg_penetration,
                "quarterly_sales": latest["quarterly_sales"],
                "reported_sales": latest["reported_sales"],
                "growth_rate_pct": latest["growth_rate_pct"],
                "sales_captured_in_db_pct": latest["sales_captured_in_db_pct"],
                "source_file": source_file,
            }
        )
    return output_rows


def collect_regression_history(snapshot: SheetSnapshot, anchor_row: int, x_col: int, y_col: int) -> List[Dict[str, Any]]:
    rows: List[Dict[str, Any]] = []
    r = min(anchor_row - 1, snapshot.base_row + snapshot.row_count - 1)
    found_block = False

    while r >= snapshot.base_row:
        x_value = to_float(snapshot.get(r, x_col))
        y_value = to_float(snapshot.get(r, y_col))
        if x_value is not None and y_value is not None:
            rows.append({"row": r, "x": x_value, "y": y_value})
            found_block = True
        elif found_block:
            break
        r -= 1

    rows.reverse()
    return rows


def rounded_signature(values: Tuple[Optional[float], ...]) -> Tuple[Optional[float], ...]:
    signature = []
    for value in values:
        if value is None:
            signature.append(None)
        else:
            signature.append(round(value, 10))
    return tuple(signature)


def extract_regression_candidates(sheet: xw.Sheet, metadata: Dict[str, str], source_file: str) -> List[Dict[str, Any]]:
    snapshot = make_sheet_snapshot(sheet)
    anchor_row, anchor_col = find_max_anchor(snapshot)
    x_col = anchor_col - 11
    y_col = anchor_col - 7

    history = collect_regression_history(snapshot, anchor_row, x_col, y_col)
    if len(history) < 2:
        return []

    max_n = min(MAX_QUARTERS, len(history))
    temp_row = max(anchor_row + 2, snapshot.base_row + snapshot.row_count + 3)
    temp_col = max(anchor_col + 3, snapshot.base_col + snapshot.col_count + 3)

    intercept_cell = sheet.range((temp_row, temp_col))
    slope_cell = sheet.range((temp_row, temp_col + 1))
    forecast_cell = sheet.range((temp_row, temp_col + 2))
    max_cell = sheet.range((temp_row, temp_col + 3))
    min_cell = sheet.range((temp_row, temp_col + 4))

    next_x = to_float(snapshot.get(anchor_row, x_col))
    if next_x is None:
        next_x = history[-1]["x"] + 1

    output_rows: List[Dict[str, Any]] = []
    previous_signature: Optional[Tuple[Optional[float], ...]] = None

    for n_quarters in range(2, max_n + 1):
        segment = history[-n_quarters:]
        start_row = segment[0]["row"]
        end_row = segment[-1]["row"]

        write_formula2(
            intercept_cell,
            f"=INTERCEPT(R{start_row}C{y_col}:R{end_row}C{y_col},R{start_row}C{x_col}:R{end_row}C{x_col})",
        )
        write_formula2(
            slope_cell,
            f"=SLOPE(R{start_row}C{y_col}:R{end_row}C{y_col},R{start_row}C{x_col}:R{end_row}C{x_col})",
        )
        write_formula2(
            forecast_cell,
            f"=R{temp_row}C{temp_col}+R{temp_row}C{temp_col + 1}*{next_x}",
        )
        write_formula2(
            max_cell,
            f"=MAX(R{start_row}C{y_col}:R{end_row}C{y_col},R{temp_row}C{temp_col + 2})",
        )
        write_formula2(
            min_cell,
            f"=MIN(R{start_row}C{y_col}:R{end_row}C{y_col},R{temp_row}C{temp_col + 2})",
        )
        sheet.book.app.calculate()

        intercept_value = to_float(intercept_cell.value)
        slope_value = to_float(slope_cell.value)
        forecast_total_without_sa = to_float(forecast_cell.value)
        forecast_max = to_float(max_cell.value)
        forecast_min = to_float(min_cell.value)

        signature = rounded_signature(
            (intercept_value, slope_value, forecast_total_without_sa, forecast_max, forecast_min)
        )
        if signature == previous_signature:
            continue
        previous_signature = signature

        output_rows.append(
            {
                "model": metadata["model"],
                "ticker": metadata["ticker"],
                "model_period": metadata["model_period"],
                "model_date": metadata["model_date"],
                "method": "regression",
                "parameter_name": "num_quarters_used",
                "parameter_value": n_quarters,
                "num_quarters_used": n_quarters,
                "forecast_value": forecast_total_without_sa,
                "actual_value": "",
                "forecast_max": forecast_max,
                "forecast_min": forecast_min,
                "range_width": calc_range_width(forecast_max, forecast_min),
                "intercept": intercept_value,
                "slope": slope_value,
                "source_file": source_file,
            }
        )
    return output_rows


def next_output_path(output_root: Path, input_folder_name: str) -> Path:
    output_root.mkdir(parents=True, exist_ok=True)
    base = output_root / f"{input_folder_name}_PARAM.xlsx"
    if not base.exists():
        return base

    index = 1
    while True:
        candidate = output_root / f"{input_folder_name}_PARAM.{index}.xlsx"
        if not candidate.exists():
            return candidate
        index += 1


def write_sheet(ws, columns: List[str], rows: List[Dict[str, Any]]) -> None:
    ws.append(columns)
    for row in rows:
        ws.append([row.get(col, "") for col in columns])

    for col_idx in range(1, len(columns) + 1):
        ws.cell(row=1, column=col_idx).font = Font(bold=True)

    ws.freeze_panes = "A2"
    end_col = get_column_letter(len(columns))
    ws.auto_filter.ref = f"A1:{end_col}{max(ws.max_row, 1)}"

    for col_idx, col_name in enumerate(columns, start=1):
        max_len = len(col_name)
        for row in rows:
            value = row.get(col_name, "")
            value_len = len(str(value)) if value is not None else 0
            if value_len > max_len:
                max_len = value_len
        ws.column_dimensions[get_column_letter(col_idx)].width = min(60, max(12, max_len + 2))


def write_output_workbook(
    output_path: Path,
    empirical_rows: List[Dict[str, Any]],
    regression_rows: List[Dict[str, Any]],
) -> None:
    out_wb = Workbook()
    default_sheet = out_wb.active
    out_wb.remove(default_sheet)

    empirical_ws = out_wb.create_sheet("empirical_candidates")
    regression_ws = out_wb.create_sheet("regression_candidates")

    write_sheet(empirical_ws, EMPIRICAL_COLUMNS, empirical_rows)
    write_sheet(regression_ws, REGRESSION_COLUMNS, regression_rows)
    out_wb.save(output_path)


def process_workbook(wb: xw.Book, metadata: Dict[str, str], source_file: str) -> Tuple[List[Dict[str, Any]], List[Dict[str, Any]]]:
    empirical_rows: List[Dict[str, Any]] = []
    regression_rows: List[Dict[str, Any]] = []

    empirical_sheet = get_sheet_case_insensitive(wb, "Empirical Model")
    regression_sheet = get_sheet_case_insensitive(wb, "Regression Model")

    if empirical_sheet is None:
        print(f"Skipped empirical extraction in {source_file}: sheet 'Empirical Model' not found")
    else:
        empirical_rows = extract_empirical_candidates(empirical_sheet, metadata, source_file)

    if regression_sheet is None:
        print(f"Skipped regression extraction in {source_file}: sheet 'Regression Model' not found")
    else:
        regression_rows = extract_regression_candidates(regression_sheet, metadata, source_file)

    return empirical_rows, regression_rows


def main() -> None:
    if not input_dir.exists():
        print(f"Input folder not found: {input_dir}")
        return

    app = xw.App(visible=False, add_book=False)
    app.display_alerts = False
    app.screen_updating = False

    files_processed = 0
    empirical_rows: List[Dict[str, Any]] = []
    regression_rows: List[Dict[str, Any]] = []

    try:
        for file_path in sorted(input_dir.iterdir()):
            if not file_path.is_file():
                print(f"Skipped {file_path.name}: not a file")
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
                wb = app.books.open(str(file_path), update_links=False)
                metadata = parse_file_metadata(file_path.name)
                empirical_part, regression_part = process_workbook(wb, metadata, file_path.name)
                empirical_rows.extend(empirical_part)
                regression_rows.extend(regression_part)
                files_processed += 1
            except Exception as exc:
                print(f"Skipped {file_path.name}: {exc}")
            finally:
                if wb is not None:
                    close_source_workbook(wb)
    finally:
        app.quit()

    output_path = next_output_path(output_dir, input_dir.name)
    write_output_workbook(output_path, empirical_rows, regression_rows)

    print(f"Output path: {output_path}")
    print(f"Number of files processed: {files_processed}")
    print(f"Number of empirical rows: {len(empirical_rows)}")
    print(f"Number of regression rows: {len(regression_rows)}")


if __name__ == "__main__":
    main()
