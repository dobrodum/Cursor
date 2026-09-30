#!/usr/bin/env python3
from __future__ import annotations

import re
from dataclasses import dataclass
from datetime import date
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional

import xlwings as xw
from openpyxl import Workbook
from openpyxl.styles import Font

# Configure these paths before running.
input_dir = Path("./input")
output_dir = Path("./output")

EMPIRICAL_HEADERS = [
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

REGRESSION_HEADERS = [
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


@dataclass(frozen=True)
class FileLabel:
    model: str
    ticker: str
    model_period: str
    model_date: str


def normalize_text(value: Any) -> str:
    if value is None:
        return ""
    text = str(value).strip().lower()
    text = re.sub(r"[%/()\-]+", " ", text)
    return re.sub(r"\s+", " ", text).strip()


def is_blank(value: Any) -> bool:
    if value is None:
        return True
    if isinstance(value, str) and value.strip() == "":
        return True
    return False


def to_float(value: Any) -> Optional[float]:
    if value is None:
        return None
    if isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        return float(value)
    if isinstance(value, str):
        text = value.strip().replace(",", "").replace("%", "")
        if text == "":
            return None
        try:
            return float(text)
        except ValueError:
            return None
    return None


def safe_diff(lhs: Any, rhs: Any) -> Optional[float]:
    left_num = to_float(lhs)
    right_num = to_float(rhs)
    if left_num is None or right_num is None:
        return None
    return left_num - right_num


def parse_file_label(file_name: str) -> Optional[FileLabel]:
    base_name = Path(file_name).stem
    parts = [part.strip() for part in base_name.split(" - ")]
    if len(parts) < 3:
        return None

    ticker = parts[1]
    period_segment = parts[2]
    period_match = re.search(r"(Early|Mid|Late)([A-Za-z]{3})(\d{4})", period_segment, re.IGNORECASE)
    if not period_match:
        return None

    band = period_match.group(1).title()
    month_abbrev = period_match.group(2).title()
    year = int(period_match.group(3))

    month_map = {
        "Jan": 1,
        "Feb": 2,
        "Mar": 3,
        "Apr": 4,
        "May": 5,
        "Jun": 6,
        "Jul": 7,
        "Aug": 8,
        "Sep": 9,
        "Oct": 10,
        "Nov": 11,
        "Dec": 12,
    }
    day_map = {"Early": 5, "Mid": 15, "Late": 25}

    if month_abbrev not in month_map or band not in day_map:
        return None

    model_period = f"{band}{month_abbrev}_{year}"
    model_date = date(year, month_map[month_abbrev], day_map[band]).isoformat()
    model = f"{ticker}_{model_period}"
    return FileLabel(model=model, ticker=ticker, model_period=model_period, model_date=model_date)


def get_unique_output_path(output_root: Path, input_folder_name: str) -> Path:
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
        try:
            wb.close()
        except Exception:
            pass


def find_anchor_cell(ws: xw.Sheet) -> Optional[xw.Range]:
    used = ws.used_range
    start_row = used.row
    start_col = used.column
    values = used.value
    if values is None:
        return None

    if not isinstance(values, list):
        values = [[values]]
    elif values and not isinstance(values[0], list):
        values = [values]

    for r_idx, row_values in enumerate(values):
        for c_idx, cell_value in enumerate(row_values):
            if isinstance(cell_value, str) and cell_value.strip().lower() == "max":
                return ws.cells(start_row + r_idx, start_col + c_idx)
    return None


def build_header_offsets(
    ws: xw.Sheet, header_row: int, anchor_col: int, scan_window: int = 35
) -> Dict[str, int]:
    offsets: Dict[str, int] = {}
    first_col = max(1, anchor_col - scan_window)
    last_col = anchor_col + scan_window
    header_values = ws.range((header_row, first_col), (header_row, last_col)).value
    if not isinstance(header_values, list):
        header_values = [header_values]

    for index, cell_value in enumerate(header_values):
        cell_text = normalize_text(cell_value)
        if cell_text:
            col = first_col + index
            offsets[cell_text] = col - anchor_col
    return offsets


def offset_for_aliases(
    header_offsets: Dict[str, int], aliases: Iterable[str], default: Optional[int] = None
) -> Optional[int]:
    for alias in aliases:
        alias_text = normalize_text(alias)
        if alias_text in header_offsets:
            return header_offsets[alias_text]
    for header_text, offset in header_offsets.items():
        for alias in aliases:
            if normalize_text(alias) in header_text:
                return offset
    return default


def value_at_offset(ws: xw.Sheet, row: int, anchor_col: int, offset: Optional[int]) -> Any:
    if offset is None:
        return None
    col = anchor_col + offset
    if col < 1:
        return None
    return ws.cells(row, col).value


def is_empirical_row_empty(row_data: Dict[str, Any]) -> bool:
    checks = (
        row_data.get("forecast_value"),
        row_data.get("actual_value"),
        row_data.get("forecast_max"),
        row_data.get("forecast_min"),
        row_data.get("avg_penetration_pct"),
        row_data.get("quarterly_sales"),
        row_data.get("reported_sales"),
    )
    return all(is_blank(value) for value in checks)


def is_regression_row_empty(row_data: Dict[str, Any]) -> bool:
    checks = (
        row_data.get("forecast_value"),
        row_data.get("forecast_max"),
        row_data.get("forecast_min"),
        row_data.get("intercept"),
        row_data.get("slope"),
    )
    return all(is_blank(value) for value in checks)


def process_empirical_sheet(wb: xw.Book, label: FileLabel, source_file: str) -> List[Dict[str, Any]]:
    sheet_names = {sheet.name for sheet in wb.sheets}
    if "Empirical Model" not in sheet_names:
        return []

    ws = wb.sheets["Empirical Model"]
    anchor = find_anchor_cell(ws)
    if anchor is None:
        return []

    anchor_row = anchor.row
    anchor_col = anchor.column
    header_offsets = build_header_offsets(ws, anchor_row, anchor_col)
    data_start_row = anchor_row + 1
    n_quarters = 10

    used = ws.used_range
    helper_col = used.column + used.columns.count + 6
    penetration_source_offset = offset_for_aliases(
        header_offsets,
        aliases=[
            "penetration",
            "penetration pct",
            "sales captured in db pct",
            "sales captured",
        ],
    )
    avg_penetration_direct_offset = offset_for_aliases(
        header_offsets,
        aliases=["avg penetration pct", "average penetration", "avg penetration"],
    )

    formulas_written = False
    if penetration_source_offset is not None:
        source_col = anchor_col + penetration_source_offset
        for i in range(n_quarters):
            row = data_start_row + i
            helper = ws.cells(row, helper_col)
            helper.formula2 = (
                f'=IFERROR(AVERAGE(R{data_start_row}C{source_col}:R{row}C{source_col}),"")'
            )
            formulas_written = True
    if formulas_written:
        wb.app.calculate()

    num_q_offset = offset_for_aliases(
        header_offsets,
        aliases=["num quarters used", "num qtrs used", "quarters used", "n quarters"],
    )
    last_quarter_offset = offset_for_aliases(
        header_offsets,
        aliases=["last quarter used", "last quarter", "last qtr used"],
    )
    forecast_value_offset = offset_for_aliases(
        header_offsets,
        aliases=[
            "estimated total sold",
            "est total sold",
            "total sold",
            "forecast value",
            "forecast",
            "tot fcst",
        ],
    )
    actual_value_offset = offset_for_aliases(
        header_offsets,
        aliases=["reported sales", "actual value", "actual sales"],
    )
    forecast_min_offset = offset_for_aliases(
        header_offsets,
        aliases=["forecast min", "min"],
        default=1,
    )
    quarterly_sales_offset = offset_for_aliases(
        header_offsets,
        aliases=["quarterly sales", "qtr sales"],
    )
    reported_sales_offset = offset_for_aliases(
        header_offsets,
        aliases=["reported sales"],
        default=actual_value_offset,
    )
    growth_rate_offset = offset_for_aliases(
        header_offsets,
        aliases=["growth rate pct", "growth rate", "growth"],
    )
    captured_pct_offset = offset_for_aliases(
        header_offsets,
        aliases=["sales captured in db pct", "captured in db", "captured pct"],
    )

    rows: List[Dict[str, Any]] = []
    consecutive_empty = 0
    for i in range(n_quarters):
        row = data_start_row + i
        num_quarters_used = value_at_offset(ws, row, anchor_col, num_q_offset)
        if is_blank(num_quarters_used):
            num_quarters_used = i + 1

        avg_penetration_pct = None
        if formulas_written:
            avg_penetration_pct = ws.cells(row, helper_col).value
        if is_blank(avg_penetration_pct):
            avg_penetration_pct = value_at_offset(ws, row, anchor_col, avg_penetration_direct_offset)

        forecast_max = value_at_offset(ws, row, anchor_col, 0)
        forecast_min = value_at_offset(ws, row, anchor_col, forecast_min_offset)
        forecast_value = value_at_offset(ws, row, anchor_col, forecast_value_offset)
        actual_value = value_at_offset(ws, row, anchor_col, actual_value_offset)

        output_row = {
            "model": label.model,
            "ticker": label.ticker,
            "model_period": label.model_period,
            "model_date": label.model_date,
            "method": "empirical",
            "parameter_name": "avg_penetration_pct",
            "parameter_value": avg_penetration_pct,
            "num_quarters_used": num_quarters_used,
            "last_quarter_used": value_at_offset(ws, row, anchor_col, last_quarter_offset),
            "forecast_value": forecast_value,
            "actual_value": actual_value,
            "forecast_max": forecast_max,
            "forecast_min": forecast_min,
            "range_width": safe_diff(forecast_max, forecast_min),
            "avg_penetration_pct": avg_penetration_pct,
            "quarterly_sales": value_at_offset(ws, row, anchor_col, quarterly_sales_offset),
            "reported_sales": value_at_offset(ws, row, anchor_col, reported_sales_offset),
            "growth_rate_pct": value_at_offset(ws, row, anchor_col, growth_rate_offset),
            "sales_captured_in_db_pct": value_at_offset(ws, row, anchor_col, captured_pct_offset),
            "source_file": source_file,
        }

        if is_empirical_row_empty(output_row):
            consecutive_empty += 1
            if consecutive_empty >= 2:
                break
            continue

        consecutive_empty = 0
        rows.append(output_row)

    return rows


def process_regression_sheet(wb: xw.Book, label: FileLabel, source_file: str) -> List[Dict[str, Any]]:
    sheet_names = {sheet.name for sheet in wb.sheets}
    if "Regression Model" not in sheet_names:
        return []

    ws = wb.sheets["Regression Model"]
    anchor = find_anchor_cell(ws)
    if anchor is None:
        return []

    anchor_row = anchor.row
    anchor_col = anchor.column
    header_offsets = build_header_offsets(ws, anchor_row, anchor_col)
    data_start_row = anchor_row + 1
    n_quarters = 10

    y_col = anchor_col - 7
    x_col = anchor_col - 11

    used = ws.used_range
    intercept_col = used.column + used.columns.count + 6
    slope_col = intercept_col + 1

    for i in range(n_quarters):
        row = data_start_row + i
        ws.cells(row, intercept_col).formula2 = (
            f'=IFERROR(INTERCEPT(R{data_start_row}C{y_col}:R{row}C{y_col},'
            f'R{data_start_row}C{x_col}:R{row}C{x_col}),"")'
        )
        ws.cells(row, slope_col).formula2 = (
            f'=IFERROR(SLOPE(R{data_start_row}C{y_col}:R{row}C{y_col},'
            f'R{data_start_row}C{x_col}:R{row}C{x_col}),"")'
        )
    wb.app.calculate()

    num_q_offset = offset_for_aliases(
        header_offsets,
        aliases=["num quarters used", "num qtrs used", "quarters used", "n quarters"],
    )
    forecast_value_offset = offset_for_aliases(
        header_offsets,
        aliases=[
            "tot fcst w/o sa",
            "tot fcst wo sa",
            "total forecast without sa",
            "forecast total without sa",
        ],
    )
    actual_value_offset = offset_for_aliases(
        header_offsets,
        aliases=["actual value", "actual sales", "reported sales"],
    )
    forecast_min_offset = offset_for_aliases(
        header_offsets,
        aliases=["forecast min", "min"],
        default=1,
    )

    rows: List[Dict[str, Any]] = []
    consecutive_empty = 0
    for i in range(n_quarters):
        row = data_start_row + i
        num_quarters_used = value_at_offset(ws, row, anchor_col, num_q_offset)
        if is_blank(num_quarters_used):
            num_quarters_used = i + 1

        intercept_value = ws.cells(row, intercept_col).value
        slope_value = ws.cells(row, slope_col).value
        forecast_max = value_at_offset(ws, row, anchor_col, 0)
        forecast_min = value_at_offset(ws, row, anchor_col, forecast_min_offset)
        forecast_value = value_at_offset(ws, row, anchor_col, forecast_value_offset)
        actual_value = value_at_offset(ws, row, anchor_col, actual_value_offset)

        output_row = {
            "model": label.model,
            "ticker": label.ticker,
            "model_period": label.model_period,
            "model_date": label.model_date,
            "method": "regression",
            "parameter_name": "num_quarters_used",
            "parameter_value": num_quarters_used,
            "num_quarters_used": num_quarters_used,
            "forecast_value": forecast_value,
            "actual_value": actual_value if not is_blank(actual_value) else None,
            "forecast_max": forecast_max,
            "forecast_min": forecast_min,
            "range_width": safe_diff(forecast_max, forecast_min),
            "intercept": intercept_value,
            "slope": slope_value,
            "source_file": source_file,
        }

        if is_regression_row_empty(output_row):
            consecutive_empty += 1
            if consecutive_empty >= 2:
                break
            continue

        consecutive_empty = 0
        rows.append(output_row)

    if len(rows) >= 2:
        prev_row = rows[-2]
        last_row = rows[-1]
        prev_signature = (
            prev_row.get("forecast_value"),
            prev_row.get("forecast_max"),
            prev_row.get("forecast_min"),
            prev_row.get("intercept"),
            prev_row.get("slope"),
        )
        last_signature = (
            last_row.get("forecast_value"),
            last_row.get("forecast_max"),
            last_row.get("forecast_min"),
            last_row.get("intercept"),
            last_row.get("slope"),
        )
        if prev_signature == last_signature:
            rows.pop()

    return rows


def write_sheet(ws: Any, headers: List[str], rows: List[Dict[str, Any]]) -> None:
    ws.append(headers)
    for row in rows:
        ws.append([row.get(header) for header in headers])

    for cell in ws[1]:
        cell.font = Font(bold=True)

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions

    for col_idx, header in enumerate(headers, start=1):
        max_len = len(header)
        for row_idx in range(2, ws.max_row + 1):
            value = ws.cell(row=row_idx, column=col_idx).value
            if value is None:
                continue
            max_len = max(max_len, len(str(value)))
        ws.column_dimensions[ws.cell(row=1, column=col_idx).column_letter].width = min(max_len + 2, 40)


def write_output_workbook(
    output_path: Path, empirical_rows: List[Dict[str, Any]], regression_rows: List[Dict[str, Any]]
) -> None:
    wb = Workbook()
    default_sheet = wb.active
    wb.remove(default_sheet)

    empirical_ws = wb.create_sheet("empirical_candidates")
    regression_ws = wb.create_sheet("regression_candidates")

    write_sheet(empirical_ws, EMPIRICAL_HEADERS, empirical_rows)
    write_sheet(regression_ws, REGRESSION_HEADERS, regression_rows)
    wb.save(output_path)


def main() -> None:
    input_root = input_dir.expanduser().resolve()
    output_root = output_dir.expanduser().resolve()

    if not input_root.exists() or not input_root.is_dir():
        raise FileNotFoundError(f"Input directory does not exist: {input_root}")

    input_folder_name = input_root.name
    output_path = get_unique_output_path(output_root, input_folder_name)

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
        for file_path in sorted(input_root.iterdir(), key=lambda p: p.name.lower()):
            if not file_path.is_file():
                print(f"Skipped file: {file_path.name} (not a file)")
                continue
            if file_path.name.startswith("~"):
                print(f"Skipped file: {file_path.name} (temporary file)")
                continue
            if file_path.suffix.lower() != ".xlsx":
                print(f"Skipped file: {file_path.name} (not .xlsx)")
                continue

            label = parse_file_label(file_path.name)
            if label is None:
                print(f"Skipped file: {file_path.name} (could not parse ticker/period)")
                continue

            print(f"Processed file: {file_path.name}")
            wb: Optional[xw.Book] = None
            try:
                wb = app.books.open(str(file_path), update_links=False)
                empirical_rows.extend(process_empirical_sheet(wb, label, file_path.name))
                regression_rows.extend(process_regression_sheet(wb, label, file_path.name))
                processed_files += 1
            except Exception as exc:
                print(f"Skipped file: {file_path.name} (error: {exc})")
            finally:
                if wb is not None:
                    close_workbook_safely(wb)
    finally:
        app.quit()

    write_output_workbook(output_path, empirical_rows, regression_rows)

    print(f"Output path: {output_path}")
    print(f"Number of files processed: {processed_files}")
    print(f"Number of empirical rows: {len(empirical_rows)}")
    print(f"Number of regression rows: {len(regression_rows)}")


if __name__ == "__main__":
    main()
