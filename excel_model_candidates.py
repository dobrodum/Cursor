#!/usr/bin/env python3
from __future__ import annotations

import math
import re
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Sequence, Tuple

import xlwings as xw
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

# User-configurable paths
input_dir = "input"
output_dir = "output"

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

DAY_BY_PERIOD = {"early": 5, "mid": 15, "late": 25}


def normalize_label(value: Any) -> str:
    if value is None:
        return ""
    text = str(value).strip().lower()
    text = text.replace("%", " pct ")
    text = text.replace("&", " and ")
    text = re.sub(r"[^a-z0-9]+", " ", text)
    return re.sub(r"\s+", " ", text).strip()


def as_row(values: Any) -> List[Any]:
    if values is None:
        return []
    if isinstance(values, list):
        if values and isinstance(values[0], list):
            return values[0]
        return values
    return [values]


def as_col(values: Any) -> List[Any]:
    if values is None:
        return []
    if isinstance(values, list):
        if values and isinstance(values[0], list):
            return [row[0] for row in values]
        return values
    return [values]


def to_float(value: Any) -> Optional[float]:
    if value is None or value == "":
        return None
    if isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        return float(value)
    text = str(value).strip().replace(",", "")
    if text == "":
        return None
    try:
        return float(text)
    except ValueError:
        return None


def to_int(value: Any) -> Optional[int]:
    number = to_float(value)
    if number is None:
        return None
    if math.isnan(number):
        return None
    return int(round(number))


def find_unique_output_path(input_path: Path, output_path: Path) -> Path:
    output_path.mkdir(parents=True, exist_ok=True)
    base = f"{input_path.name}_PARAM"
    candidate = output_path / f"{base}.xlsx"
    if not candidate.exists():
        return candidate

    index = 1
    while True:
        candidate = output_path / f"{base}.{index}.xlsx"
        if not candidate.exists():
            return candidate
        index += 1


def parse_filename_metadata(file_name: str) -> Optional[Dict[str, str]]:
    stem = Path(file_name).stem
    parts = [part.strip() for part in stem.split(" - ")]
    if len(parts) < 3:
        return None

    ticker = parts[1].strip()
    period_token = parts[2].replace("_Send", "")
    match = re.fullmatch(
        r"(?P<timing>Early|Mid|Late)(?P<month>[A-Za-z]+)(?P<year>\d{4}).*",
        period_token,
        flags=re.IGNORECASE,
    )
    if not match:
        return None

    timing_raw = match.group("timing")
    month_raw = match.group("month")
    year_raw = match.group("year")

    timing = timing_raw[0].upper() + timing_raw[1:].lower()
    month_token = month_raw[0].upper() + month_raw[1:].lower()
    try:
        month_number = datetime.strptime(month_token[:3], "%b").month
    except ValueError:
        return None
    year_number = int(year_raw)
    day = DAY_BY_PERIOD[timing.lower()]
    model_period = f"{timing}{month_token}_{year_number}"
    model_date = datetime(year_number, month_number, day).date().isoformat()
    model = f"{ticker}_{model_period}"
    return {
        "ticker": ticker,
        "model_period": model_period,
        "model_date": model_date,
        "model": model,
    }


def safe_close_source_workbook(wb: xw.Book) -> None:
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
        # Last-resort fallback when COM state is already torn down.
        pass


def find_anchor(sheet: xw.Sheet, anchor_text: str = "max") -> Optional[Tuple[int, int]]:
    used = sheet.used_range
    used_values = used.value
    if used_values is None:
        return None
    if not isinstance(used_values, list):
        used_values = [[used_values]]
    elif used_values and not isinstance(used_values[0], list):
        used_values = [used_values]

    start_row = used.row
    start_col = used.column
    for row_offset, row_values in enumerate(used_values):
        for col_offset, cell_value in enumerate(row_values):
            if isinstance(cell_value, str) and normalize_label(cell_value) == anchor_text:
                return start_row + row_offset, start_col + col_offset
    return None


def build_header_map(
    sheet: xw.Sheet, header_row: int, first_col: int, last_col: int
) -> Dict[str, int]:
    row_values = as_row(sheet.range((header_row, first_col), (header_row, last_col)).value)
    header_map: Dict[str, int] = {}
    for index, value in enumerate(row_values):
        normalized = normalize_label(value)
        if normalized:
            header_map.setdefault(normalized, first_col + index)
    return header_map


def lookup_column(
    header_map: Dict[str, int], aliases: Sequence[str], fallback: Optional[int] = None
) -> Optional[int]:
    alias_keys = [normalize_label(alias) for alias in aliases]
    for alias in alias_keys:
        if alias in header_map:
            return header_map[alias]
    for key, col in header_map.items():
        for alias in alias_keys:
            if alias and (alias in key or key in alias):
                return col
    return fallback


def has_any_data(values: Iterable[Any]) -> bool:
    return any(value not in (None, "") for value in values)


def collect_empirical_rows(
    wb: xw.Book, sheet: xw.Sheet, meta: Dict[str, str], source_file: str
) -> List[Dict[str, Any]]:
    anchor = find_anchor(sheet, "max")
    if not anchor:
        print(f"  - skipped empirical sheet (missing 'max' anchor)")
        return []

    anchor_row, anchor_col = anchor
    used = sheet.used_range
    last_row = used.last_cell.row
    last_col = used.last_cell.column

    header_map = build_header_map(sheet, anchor_row, 1, last_col)
    min_col = lookup_column(header_map, ["min"], fallback=anchor_col + 1)
    forecast_col = lookup_column(
        header_map,
        ["estimated total sold", "forecast value", "forecast", "total sold estimate"],
        fallback=anchor_col - 2,
    )
    num_quarters_col = lookup_column(
        header_map, ["num quarters used", "num_quarters_used", "n quarters"]
    )
    last_quarter_col = lookup_column(
        header_map, ["last quarter used", "last_quarter_used", "last quarter"]
    )
    avg_penetration_col = lookup_column(
        header_map, ["avg penetration pct", "average penetration", "avg penetration"]
    )
    quarterly_sales_col = lookup_column(
        header_map, ["quarterly sales", "quarterly_sales", "qtr sales"]
    )
    reported_sales_col = lookup_column(
        header_map, ["reported sales", "reported_sales", "actual sales"]
    )
    growth_rate_col = lookup_column(
        header_map, ["growth rate pct", "growth_rate_pct", "growth rate"]
    )
    captured_pct_col = lookup_column(
        header_map,
        ["sales captured in db pct", "sales_captured_in_db_pct", "captured in db pct"],
    )

    data_start = anchor_row + 1
    data_end = min(last_row, data_start + N_QUARTERS - 1)
    if data_start > data_end:
        print("  - skipped empirical sheet (no data rows)")
        return []

    temp_col = last_col + 2
    formula_rows: List[int] = []
    rows: List[Dict[str, Any]] = []

    for row in range(data_start, data_end + 1):
        row_probe = [
            sheet.range((row, col)).value
            for col in [
                num_quarters_col or anchor_col - 12,
                forecast_col,
                anchor_col,
                min_col,
                quarterly_sales_col or anchor_col - 9,
                reported_sales_col or anchor_col - 8,
            ]
            if col and col > 0
        ]
        if not has_any_data(row_probe):
            continue

        avg_pen = sheet.range((row, avg_penetration_col)).value if avg_penetration_col else None
        if avg_pen in (None, ""):
            quarter_end_col = anchor_col - 1
            quarter_start_col = quarter_end_col - (N_QUARTERS - 1)
            if quarter_start_col > 0:
                start_offset = quarter_start_col - temp_col
                end_offset = quarter_end_col - temp_col
                sheet.range((row, temp_col)).formula2 = (
                    f"=AVERAGE(RC[{start_offset}]:RC[{end_offset}])"
                )
                formula_rows.append(row)

        num_quarters = (
            to_int(sheet.range((row, num_quarters_col)).value)
            if num_quarters_col
            else (row - data_start + 1)
        )
        last_quarter = (
            sheet.range((row, last_quarter_col)).value if last_quarter_col else None
        )
        forecast_value = (
            to_float(sheet.range((row, forecast_col)).value) if forecast_col else None
        )
        forecast_max = to_float(sheet.range((row, anchor_col)).value)
        forecast_min = (
            to_float(sheet.range((row, min_col)).value) if min_col and min_col > 0 else None
        )
        quarterly_sales = (
            to_float(sheet.range((row, quarterly_sales_col)).value)
            if quarterly_sales_col
            else None
        )
        reported_sales = (
            to_float(sheet.range((row, reported_sales_col)).value)
            if reported_sales_col
            else None
        )
        growth_rate = (
            to_float(sheet.range((row, growth_rate_col)).value) if growth_rate_col else None
        )
        captured_pct = (
            to_float(sheet.range((row, captured_pct_col)).value) if captured_pct_col else None
        )

        rows.append(
            {
                "model": meta["model"],
                "ticker": meta["ticker"],
                "model_period": meta["model_period"],
                "model_date": meta["model_date"],
                "method": "empirical",
                "parameter_name": "avg_penetration_pct",
                "parameter_value": None,  # populated below
                "num_quarters_used": num_quarters,
                "last_quarter_used": last_quarter,
                "forecast_value": forecast_value,
                "actual_value": reported_sales,
                "forecast_max": forecast_max,
                "forecast_min": forecast_min,
                "range_width": (
                    forecast_max - forecast_min
                    if forecast_max is not None and forecast_min is not None
                    else None
                ),
                "avg_penetration_pct": avg_pen,
                "quarterly_sales": quarterly_sales,
                "reported_sales": reported_sales,
                "growth_rate_pct": growth_rate,
                "sales_captured_in_db_pct": captured_pct,
                "source_file": source_file,
                "_sheet_row": row,
                "_temp_col": temp_col,
            }
        )

    if formula_rows:
        wb.app.calculate()
        for row_data in rows:
            if row_data["avg_penetration_pct"] in (None, ""):
                row_data["avg_penetration_pct"] = to_float(
                    sheet.range((row_data["_sheet_row"], temp_col)).value
                )
        sheet.range((data_start, temp_col), (data_end, temp_col)).clear_contents()

    for row_data in rows:
        row_data["parameter_value"] = row_data["avg_penetration_pct"]
        row_data.pop("_sheet_row", None)
        row_data.pop("_temp_col", None)

    return rows


def rows_equal_for_dedup(previous: Dict[str, Any], current: Dict[str, Any]) -> bool:
    compare_keys = [
        "num_quarters_used",
        "forecast_value",
        "forecast_max",
        "forecast_min",
        "intercept",
        "slope",
    ]
    for key in compare_keys:
        left = previous.get(key)
        right = current.get(key)
        if isinstance(left, (int, float)) and isinstance(right, (int, float)):
            if abs(float(left) - float(right)) > 1e-9:
                return False
        else:
            if (left or "") != (right or ""):
                return False
    return True


def collect_regression_rows(
    wb: xw.Book, sheet: xw.Sheet, meta: Dict[str, str], source_file: str
) -> List[Dict[str, Any]]:
    anchor = find_anchor(sheet, "max")
    if not anchor:
        print(f"  - skipped regression sheet (missing 'max' anchor)")
        return []

    anchor_row, anchor_col = anchor
    used = sheet.used_range
    last_row = used.last_cell.row
    last_col = used.last_cell.column
    header_map = build_header_map(sheet, anchor_row, 1, last_col)

    y_col = anchor_col - 7
    x_col = anchor_col - 11
    min_col = lookup_column(header_map, ["min"], fallback=anchor_col + 1)
    num_quarters_col = lookup_column(
        header_map, ["num quarters used", "num_quarters_used", "n quarters"]
    )
    forecast_col = lookup_column(
        header_map,
        ["tot fcst w o sa", "tot fcst w/o sa", "forecast total without sa"],
        fallback=anchor_col - 3,
    )
    actual_col = lookup_column(header_map, ["actual value", "actual", "reported sales"])

    data_start = anchor_row + 1
    if data_start > last_row:
        print("  - skipped regression sheet (no data rows)")
        return []

    temp_int_col = last_col + 2
    temp_slope_col = last_col + 3
    rows: List[Dict[str, Any]] = []
    formula_rows: List[Tuple[int, int, int]] = []
    blank_streak = 0

    for row in range(data_start, last_row + 1):
        probe_values = [
            sheet.range((row, col)).value
            for col in [num_quarters_col, forecast_col, anchor_col, min_col]
            if col and col > 0
        ]
        if not has_any_data(probe_values):
            blank_streak += 1
            if blank_streak >= 3:
                break
            continue
        blank_streak = 0

        num_quarters = (
            to_int(sheet.range((row, num_quarters_col)).value)
            if num_quarters_col
            else (row - data_start + 1)
        )
        if num_quarters is None or num_quarters < 2:
            continue

        window_start = row
        window_end = min(last_row, row + num_quarters - 1)
        if window_end - window_start + 1 < 2:
            continue

        int_formula = (
            f"=INTERCEPT(R{window_start}C{y_col}:R{window_end}C{y_col},"
            f"R{window_start}C{x_col}:R{window_end}C{x_col})"
        )
        slope_formula = (
            f"=SLOPE(R{window_start}C{y_col}:R{window_end}C{y_col},"
            f"R{window_start}C{x_col}:R{window_end}C{x_col})"
        )
        sheet.range((row, temp_int_col)).formula2 = int_formula
        sheet.range((row, temp_slope_col)).formula2 = slope_formula
        formula_rows.append((row, temp_int_col, temp_slope_col))

        forecast_value = (
            to_float(sheet.range((row, forecast_col)).value) if forecast_col else None
        )
        forecast_max = to_float(sheet.range((row, anchor_col)).value)
        forecast_min = (
            to_float(sheet.range((row, min_col)).value) if min_col and min_col > 0 else None
        )
        actual_value = (
            to_float(sheet.range((row, actual_col)).value) if actual_col else None
        )

        rows.append(
            {
                "model": meta["model"],
                "ticker": meta["ticker"],
                "model_period": meta["model_period"],
                "model_date": meta["model_date"],
                "method": "regression",
                "parameter_name": "num_quarters_used",
                "parameter_value": num_quarters,
                "num_quarters_used": num_quarters,
                "forecast_value": forecast_value,
                "actual_value": actual_value,
                "forecast_max": forecast_max,
                "forecast_min": forecast_min,
                "range_width": (
                    forecast_max - forecast_min
                    if forecast_max is not None and forecast_min is not None
                    else None
                ),
                "intercept": None,
                "slope": None,
                "source_file": source_file,
                "_sheet_row": row,
            }
        )

    if formula_rows:
        wb.app.calculate()
        for row_data in rows:
            row = row_data["_sheet_row"]
            row_data["intercept"] = to_float(sheet.range((row, temp_int_col)).value)
            row_data["slope"] = to_float(sheet.range((row, temp_slope_col)).value)
        sheet.range((data_start, temp_int_col), (last_row, temp_slope_col)).clear_contents()

    deduped_rows: List[Dict[str, Any]] = []
    for row_data in rows:
        row_data.pop("_sheet_row", None)
        if deduped_rows and rows_equal_for_dedup(deduped_rows[-1], row_data):
            continue
        deduped_rows.append(row_data)
    return deduped_rows


def write_sheet(
    workbook: Workbook, sheet_name: str, columns: Sequence[str], rows: Sequence[Dict[str, Any]]
) -> None:
    ws = workbook.create_sheet(sheet_name)
    ws.append(list(columns))
    for value_cell in ws[1]:
        value_cell.font = Font(bold=True)

    for row_data in rows:
        ws.append([row_data.get(column) for column in columns])

    ws.freeze_panes = "A2"
    last_col_letter = get_column_letter(len(columns))
    last_row = max(1, ws.max_row)
    ws.auto_filter.ref = f"A1:{last_col_letter}{last_row}"

    for index, column_name in enumerate(columns, start=1):
        max_len = len(column_name)
        for row_idx in range(2, ws.max_row + 1):
            value = ws.cell(row=row_idx, column=index).value
            if value is None:
                continue
            value_len = len(str(value))
            if value_len > max_len:
                max_len = value_len
        ws.column_dimensions[get_column_letter(index)].width = max(12, min(max_len + 2, 40))


def save_output(
    output_path: Path,
    empirical_rows: Sequence[Dict[str, Any]],
    regression_rows: Sequence[Dict[str, Any]],
) -> None:
    wb = Workbook()
    wb.remove(wb.active)
    write_sheet(wb, "empirical_candidates", EMPIRICAL_COLUMNS, empirical_rows)
    write_sheet(wb, "regression_candidates", REGRESSION_COLUMNS, regression_rows)
    wb.save(output_path)


def main() -> None:
    input_path = Path(input_dir).expanduser().resolve()
    output_path = Path(output_dir).expanduser().resolve()

    if not input_path.exists() or not input_path.is_dir():
        raise FileNotFoundError(f"input_dir does not exist or is not a directory: {input_path}")

    output_file = find_unique_output_path(input_path, output_path)

    empirical_rows: List[Dict[str, Any]] = []
    regression_rows: List[Dict[str, Any]] = []
    processed_files = 0

    files = sorted(input_path.iterdir(), key=lambda path: path.name.lower())
    source_files: List[Path] = []
    for file_path in files:
        if not file_path.is_file():
            continue
        if file_path.name.startswith("~"):
            print(f"skipped {file_path.name}: temporary workbook")
            continue
        if file_path.suffix.lower() != ".xlsx":
            print(f"skipped {file_path.name}: not an .xlsx file")
            continue
        source_files.append(file_path)

    if not source_files:
        print(f"output path: {output_file}")
        print("number of files processed: 0")
        print("number of empirical rows: 0")
        print("number of regression rows: 0")
        save_output(output_file, empirical_rows, regression_rows)
        return

    app = xw.App(visible=False, add_book=False)
    app.display_alerts = False
    app.screen_updating = False
    app.enable_events = False

    try:
        for file_path in source_files:
            metadata = parse_filename_metadata(file_path.name)
            if metadata is None:
                print(f"skipped {file_path.name}: cannot parse ticker/model period from filename")
                continue

            print(f"processed {file_path.name}")
            wb = None
            try:
                wb = app.books.open(str(file_path), update_links=False)
                empirical_sheet = None
                regression_sheet = None
                try:
                    empirical_sheet = wb.sheets["Empirical Model"]
                except Exception:
                    print("  - skipped empirical sheet: not found")
                try:
                    regression_sheet = wb.sheets["Regression Model"]
                except Exception:
                    print("  - skipped regression sheet: not found")

                if empirical_sheet is not None:
                    empirical_rows.extend(
                        collect_empirical_rows(wb, empirical_sheet, metadata, file_path.name)
                    )
                if regression_sheet is not None:
                    regression_rows.extend(
                        collect_regression_rows(wb, regression_sheet, metadata, file_path.name)
                    )
                processed_files += 1
            except Exception as exc:
                print(f"skipped {file_path.name}: processing error: {exc}")
            finally:
                if wb is not None:
                    safe_close_source_workbook(wb)
    finally:
        app.quit()

    save_output(output_file, empirical_rows, regression_rows)
    print(f"output path: {output_file}")
    print(f"number of files processed: {processed_files}")
    print(f"number of empirical rows: {len(empirical_rows)}")
    print(f"number of regression rows: {len(regression_rows)}")


if __name__ == "__main__":
    main()
