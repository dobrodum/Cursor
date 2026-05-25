#!/usr/bin/env python3
"""Extract empirical and regression model candidates from Excel workbooks."""

from __future__ import annotations

import re
from datetime import datetime
from pathlib import Path
from typing import Any

import xwings as xw
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

# Configure run folders here.
input_dir = Path("./input")
output_dir = Path("./output")

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

HEADER_ALIASES = {
    "num_quarters_used": (
        "num quarters used",
        "quarters used",
        "quarter count",
        "num qtrs",
        "n quarters",
    ),
    "last_quarter_used": ("last quarter used", "last quarter", "quarter"),
    "forecast_value": (
        "estimated total sold",
        "tot fcst w/o sa",
        "tot fcst wo sa",
        "total forecast without sa",
        "forecast",
        "forecast value",
    ),
    "actual_value": (
        "actual",
        "actual value",
        "actual sales",
        "reported sales",
    ),
    "reported_sales": ("reported sales", "actual sales"),
    "forecast_max": ("max",),
    "forecast_min": ("min",),
    "avg_penetration_pct": (
        "avg penetration pct",
        "avg penetration",
        "average penetration",
        "penetration pct",
    ),
    "quarterly_sales": (
        "quarterly sales",
        "sales in db",
        "captured sales",
    ),
    "growth_rate_pct": ("growth rate pct", "growth rate", "growth %"),
    "sales_captured_in_db_pct": (
        "sales captured in db pct",
        "sales captured in db %",
        "captured in db %",
        "captured in db pct",
    ),
    "intercept": ("intercept",),
    "slope": ("slope",),
}

EMPIRICAL_FALLBACK_OFFSETS = {
    "num_quarters_used": -12,
    "last_quarter_used": -11,
    "quarterly_sales": -9,
    "reported_sales": -8,
    "avg_penetration_pct": -7,
    "growth_rate_pct": -6,
    "sales_captured_in_db_pct": -5,
    "forecast_value": -3,
    "actual_value": -2,
    "forecast_max": 0,
    "forecast_min": 1,
}

REGRESSION_FALLBACK_OFFSETS = {
    "num_quarters_used": -13,
    "forecast_value": -1,
    "actual_value": -2,
    "forecast_max": 0,
    "forecast_min": 1,
    "intercept": -5,
    "slope": -4,
}


def normalize_text(value: Any) -> str:
    if value is None:
        return ""
    return re.sub(r"\s+", " ", str(value).strip().lower())


def normalize_header(value: Any) -> str:
    text = normalize_text(value)
    text = text.replace("_", " ").replace("%", " pct")
    text = text.replace("/", " ")
    return re.sub(r"\s+", " ", text).strip()


def safe_float(value: Any) -> float | None:
    if value in (None, ""):
        return None
    if isinstance(value, bool):
        return float(value)
    if isinstance(value, (int, float)):
        return float(value)
    text = str(value).replace(",", "").strip()
    if text in ("", "-", "NA", "N/A"):
        return None
    try:
        return float(text)
    except ValueError:
        return None


def safe_int(value: Any) -> int | None:
    number = safe_float(value)
    if number is None:
        return None
    return int(round(number))


def value_or_blank(value: Any) -> Any:
    return "" if value is None else value


def select_output_path(input_folder: Path, destination_folder: Path) -> Path:
    base_name = f"{input_folder.name}_PARAM"
    first = destination_folder / f"{base_name}.xlsx"
    if not first.exists():
        return first

    index = 1
    while True:
        candidate = destination_folder / f"{base_name}.{index}.xlsx"
        if not candidate.exists():
            return candidate
        index += 1


def parse_file_label(file_name: str) -> dict[str, str]:
    stem = Path(file_name).stem
    parts = [part.strip() for part in stem.split(" - ")]
    ticker = parts[1].strip().upper() if len(parts) >= 2 and parts[1].strip() else "UNKNOWN"
    period_source = parts[2] if len(parts) >= 3 else stem

    period_match = re.search(
        r"(Early|Mid|Late)\s*([A-Za-z]+)\s*(\d{4})",
        period_source,
        flags=re.IGNORECASE,
    )
    if not period_match:
        model_period = "UnknownPeriod"
        model_date = ""
    else:
        period_prefix = period_match.group(1).title()
        month_text = period_match.group(2)
        year_text = period_match.group(3)

        month_key = month_text[:3].title()
        if month_key == "Sep" and month_text.lower().startswith("sept"):
            month_key = "Sep"

        try:
            month_num = datetime.strptime(month_key, "%b").month
        except ValueError:
            month_num = 1

        day_map = {"Early": 5, "Mid": 15, "Late": 25}
        day_num = day_map.get(period_prefix, 15)
        year_num = int(year_text)

        model_period = f"{period_prefix}{month_key}_{year_num}"
        model_date = datetime(year_num, month_num, day_num).date().isoformat()

    model = f"{ticker}_{model_period}"
    return {
        "model": model,
        "ticker": ticker,
        "model_period": model_period,
        "model_date": model_date,
    }


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
        pass


def find_anchor(sheet: xw.Sheet, target: str = "max") -> tuple[int, int] | None:
    used = sheet.used_range
    values = used.value
    if values is None:
        return None

    if not isinstance(values, list):
        values = [[values]]
    elif values and not isinstance(values[0], list):
        values = [values]

    first_row = used.row
    first_col = used.column

    best_anchor = None
    best_score = -1
    for r_idx, row in enumerate(values):
        for c_idx, cell_value in enumerate(row):
            if normalize_text(cell_value) != target:
                continue
            score = 0
            for probe in range(1, 4):
                check_idx = c_idx + probe
                if check_idx < len(row) and normalize_text(row[check_idx]) == "min":
                    score += 2
            # Prefer "max" cells that likely sit in a header row.
            if any(normalize_text(x) in ("num quarters used", "quarters used") for x in row):
                score += 1
            if score > best_score:
                best_score = score
                best_anchor = (first_row + r_idx, first_col + c_idx)

    return best_anchor


def find_header_columns(
    sheet: xw.Sheet, anchor_row: int, anchor_col: int, window: int = 35
) -> dict[str, int]:
    start_col = max(1, anchor_col - window)
    end_col = max(start_col, anchor_col + window)
    header_values = sheet.range((anchor_row, start_col), (anchor_row, end_col)).value
    if header_values is None:
        return {}
    if not isinstance(header_values, list):
        header_values = [header_values]

    normalized_to_col: dict[str, int] = {}
    for idx, raw_value in enumerate(header_values, start=start_col):
        normalized = normalize_header(raw_value)
        if normalized:
            normalized_to_col[normalized] = idx

    result: dict[str, int] = {}
    for canonical_name, aliases in HEADER_ALIASES.items():
        for alias in aliases:
            alias_key = normalize_header(alias)
            if alias_key in normalized_to_col:
                result[canonical_name] = normalized_to_col[alias_key]
                break

    # Always anchor max/min to max cell neighborhood if present.
    if "forecast_max" not in result:
        result["forecast_max"] = anchor_col
    if "forecast_min" not in result:
        result["forecast_min"] = anchor_col + 1
    return result


def resolve_columns(
    header_columns: dict[str, int], anchor_col: int, fallback_offsets: dict[str, int]
) -> dict[str, int]:
    resolved = dict(header_columns)
    for key, offset in fallback_offsets.items():
        if key not in resolved:
            resolved[key] = max(1, anchor_col + offset)
    return resolved


def build_row_map_from_n_quarters(
    sheet: xw.Sheet, num_quarters_col: int, first_data_row: int
) -> dict[int, int]:
    max_probe_rows = first_data_row + 80
    values = sheet.range((first_data_row, num_quarters_col), (max_probe_rows, num_quarters_col)).value
    if values is None:
        return {}
    if not isinstance(values, list):
        values = [values]

    row_map: dict[int, int] = {}
    for offset, raw_value in enumerate(values):
        n_q = safe_int(raw_value)
        if n_q is None or not (1 <= n_q <= N_QUARTERS):
            continue
        row_map.setdefault(n_q, first_data_row + offset)
    return row_map


def row_is_empty(values: list[Any]) -> bool:
    return all(value in (None, "") for value in values)


def read_cell(sheet: xw.Sheet, row: int, col: int) -> Any:
    return sheet.range((row, col)).value


def calculate_range_width(max_value: Any, min_value: Any) -> Any:
    max_num = safe_float(max_value)
    min_num = safe_float(min_value)
    if max_num is None or min_num is None:
        return ""
    return max_num - min_num


def write_empirical_formulas(
    sheet: xw.Sheet,
    row_pairs: list[tuple[int, int]],
    first_data_row: int,
    anchor_col: int,
    quarterly_col: int | None,
    reported_col: int | None,
) -> dict[tuple[int, int], int]:
    if quarterly_col is None or reported_col is None:
        return {}

    helper_col = anchor_col + 18
    formula_cells: dict[tuple[int, int], int] = {}

    for n_q, row in row_pairs:
        points = min(n_q, row - first_data_row + 1)
        if points <= 0:
            continue
        back = points - 1

        q_off = quarterly_col - helper_col
        r_off = reported_col - helper_col
        if back == 0:
            q_range = f"RC[{q_off}]"
            r_range = f"RC[{r_off}]"
        else:
            q_range = f"R[-{back}]C[{q_off}]:RC[{q_off}]"
            r_range = f"R[-{back}]C[{r_off}]:RC[{r_off}]"

        formula = f'=IFERROR(AVERAGE({q_range}/{r_range}),"")'
        sheet.range((row, helper_col)).formula2 = formula
        formula_cells[(n_q, row)] = helper_col

    if formula_cells:
        sheet.book.app.calculate()
    return formula_cells


def write_regression_formulas(
    sheet: xw.Sheet,
    row_pairs: list[tuple[int, int]],
    first_data_row: int,
    anchor_col: int,
    y_col: int,
    x_col: int,
) -> tuple[dict[tuple[int, int], int], dict[tuple[int, int], int]]:
    intercept_col = anchor_col + 20
    slope_col = anchor_col + 21

    intercept_cells: dict[tuple[int, int], int] = {}
    slope_cells: dict[tuple[int, int], int] = {}

    for n_q, row in row_pairs:
        points = min(n_q, row - first_data_row + 1)
        if points < 2:
            continue
        back = points - 1

        y_off_intercept = y_col - intercept_col
        x_off_intercept = x_col - intercept_col
        y_off_slope = y_col - slope_col
        x_off_slope = x_col - slope_col

        y_range_intercept = f"R[-{back}]C[{y_off_intercept}]:RC[{y_off_intercept}]"
        x_range_intercept = f"R[-{back}]C[{x_off_intercept}]:RC[{x_off_intercept}]"
        y_range_slope = f"R[-{back}]C[{y_off_slope}]:RC[{y_off_slope}]"
        x_range_slope = f"R[-{back}]C[{x_off_slope}]:RC[{x_off_slope}]"

        intercept_formula = f'=IFERROR(INTERCEPT({y_range_intercept},{x_range_intercept}),"")'
        slope_formula = f'=IFERROR(SLOPE({y_range_slope},{x_range_slope}),"")'

        sheet.range((row, intercept_col)).formula2 = intercept_formula
        sheet.range((row, slope_col)).formula2 = slope_formula
        intercept_cells[(n_q, row)] = intercept_col
        slope_cells[(n_q, row)] = slope_col

    if intercept_cells or slope_cells:
        sheet.book.app.calculate()
    return intercept_cells, slope_cells


def extract_empirical_rows(
    wb: xw.Book, meta: dict[str, str], source_file: str
) -> list[dict[str, Any]]:
    try:
        sheet = wb.sheets["Empirical Model"]
    except Exception:
        return []

    anchor = find_anchor(sheet, "max")
    if not anchor:
        return []
    anchor_row, anchor_col = anchor

    header_columns = find_header_columns(sheet, anchor_row, anchor_col)
    columns = resolve_columns(header_columns, anchor_col, EMPIRICAL_FALLBACK_OFFSETS)

    first_data_row = anchor_row + 1
    row_map = build_row_map_from_n_quarters(sheet, columns["num_quarters_used"], first_data_row)
    row_pairs = [(n_q, row_map.get(n_q, first_data_row + (n_q - 1))) for n_q in range(1, N_QUARTERS + 1)]

    formula_cells = write_empirical_formulas(
        sheet=sheet,
        row_pairs=row_pairs,
        first_data_row=first_data_row,
        anchor_col=anchor_col,
        quarterly_col=columns.get("quarterly_sales"),
        reported_col=columns.get("reported_sales") or columns.get("actual_value"),
    )

    rows: list[dict[str, Any]] = []
    for n_q, row in row_pairs:
        num_quarters_used = read_cell(sheet, row, columns["num_quarters_used"])
        last_quarter_used = read_cell(sheet, row, columns["last_quarter_used"])
        forecast_value = read_cell(sheet, row, columns["forecast_value"])
        actual_value = read_cell(sheet, row, columns["actual_value"])
        forecast_max = read_cell(sheet, row, columns["forecast_max"])
        forecast_min = read_cell(sheet, row, columns["forecast_min"])
        quarterly_sales = read_cell(sheet, row, columns["quarterly_sales"])
        reported_sales = read_cell(sheet, row, columns["reported_sales"])
        growth_rate_pct = read_cell(sheet, row, columns["growth_rate_pct"])
        sales_captured = read_cell(sheet, row, columns["sales_captured_in_db_pct"])

        if (n_q, row) in formula_cells:
            avg_penetration = read_cell(sheet, row, formula_cells[(n_q, row)])
        else:
            avg_penetration = read_cell(sheet, row, columns["avg_penetration_pct"])

        if row_is_empty(
            [
                num_quarters_used,
                forecast_value,
                forecast_max,
                forecast_min,
                avg_penetration,
                quarterly_sales,
                reported_sales,
            ]
        ):
            continue

        row_data = {
            "model": meta["model"],
            "ticker": meta["ticker"],
            "model_period": meta["model_period"],
            "model_date": meta["model_date"],
            "method": "empirical",
            "parameter_name": "avg_penetration_pct",
            "parameter_value": value_or_blank(avg_penetration),
            "num_quarters_used": value_or_blank(num_quarters_used if num_quarters_used not in (None, "") else n_q),
            "last_quarter_used": value_or_blank(last_quarter_used),
            "forecast_value": value_or_blank(forecast_value),
            "actual_value": value_or_blank(actual_value),
            "forecast_max": value_or_blank(forecast_max),
            "forecast_min": value_or_blank(forecast_min),
            "range_width": value_or_blank(calculate_range_width(forecast_max, forecast_min)),
            "avg_penetration_pct": value_or_blank(avg_penetration),
            "quarterly_sales": value_or_blank(quarterly_sales),
            "reported_sales": value_or_blank(reported_sales),
            "growth_rate_pct": value_or_blank(growth_rate_pct),
            "sales_captured_in_db_pct": value_or_blank(sales_captured),
            "source_file": source_file,
        }
        rows.append(row_data)

    return rows


def is_duplicate_regression_row(
    prev_row: dict[str, Any] | None, current_row: dict[str, Any]
) -> bool:
    if prev_row is None:
        return False
    keys = (
        "num_quarters_used",
        "forecast_value",
        "forecast_max",
        "forecast_min",
        "intercept",
        "slope",
    )
    return all(prev_row.get(key) == current_row.get(key) for key in keys)


def extract_regression_rows(
    wb: xw.Book, meta: dict[str, str], source_file: str
) -> list[dict[str, Any]]:
    try:
        sheet = wb.sheets["Regression Model"]
    except Exception:
        return []

    anchor = find_anchor(sheet, "max")
    if not anchor:
        return []
    anchor_row, anchor_col = anchor

    header_columns = find_header_columns(sheet, anchor_row, anchor_col)
    columns = resolve_columns(header_columns, anchor_col, REGRESSION_FALLBACK_OFFSETS)

    y_col = anchor_col - 7
    x_col = anchor_col - 11

    first_data_row = anchor_row + 1
    row_map = build_row_map_from_n_quarters(sheet, columns["num_quarters_used"], first_data_row)
    row_pairs = [(n_q, row_map.get(n_q, first_data_row + (n_q - 1))) for n_q in range(1, N_QUARTERS + 1)]

    intercept_cells, slope_cells = write_regression_formulas(
        sheet=sheet,
        row_pairs=row_pairs,
        first_data_row=first_data_row,
        anchor_col=anchor_col,
        y_col=y_col,
        x_col=x_col,
    )

    rows: list[dict[str, Any]] = []
    prev_row: dict[str, Any] | None = None
    for n_q, row in row_pairs:
        num_quarters_used = read_cell(sheet, row, columns["num_quarters_used"])
        forecast_value = read_cell(sheet, row, columns["forecast_value"])
        actual_value = read_cell(sheet, row, columns["actual_value"])
        forecast_max = read_cell(sheet, row, columns["forecast_max"])
        forecast_min = read_cell(sheet, row, columns["forecast_min"])

        if (n_q, row) in intercept_cells:
            intercept = read_cell(sheet, row, intercept_cells[(n_q, row)])
        else:
            intercept = read_cell(sheet, row, columns["intercept"])
        if (n_q, row) in slope_cells:
            slope = read_cell(sheet, row, slope_cells[(n_q, row)])
        else:
            slope = read_cell(sheet, row, columns["slope"])

        if row_is_empty([num_quarters_used, forecast_value, forecast_max, forecast_min, intercept, slope]):
            continue

        row_data = {
            "model": meta["model"],
            "ticker": meta["ticker"],
            "model_period": meta["model_period"],
            "model_date": meta["model_date"],
            "method": "regression",
            "parameter_name": "num_quarters_used",
            "parameter_value": value_or_blank(num_quarters_used if num_quarters_used not in (None, "") else n_q),
            "num_quarters_used": value_or_blank(num_quarters_used if num_quarters_used not in (None, "") else n_q),
            "forecast_value": value_or_blank(forecast_value),
            "actual_value": value_or_blank(actual_value),
            "forecast_max": value_or_blank(forecast_max),
            "forecast_min": value_or_blank(forecast_min),
            "range_width": value_or_blank(calculate_range_width(forecast_max, forecast_min)),
            "intercept": value_or_blank(intercept),
            "slope": value_or_blank(slope),
            "source_file": source_file,
        }

        if is_duplicate_regression_row(prev_row, row_data):
            continue

        rows.append(row_data)
        prev_row = row_data

    return rows


def write_sheet(
    wb: Workbook, sheet_name: str, columns: list[str], rows: list[dict[str, Any]]
) -> None:
    ws = wb.create_sheet(sheet_name)
    ws.append(columns)
    for row_data in rows:
        ws.append([row_data.get(column, "") for column in columns])

    for cell in ws[1]:
        cell.font = Font(bold=True)

    ws.freeze_panes = "A2"
    end_col = get_column_letter(len(columns))
    ws.auto_filter.ref = f"A1:{end_col}{max(1, ws.max_row)}"

    for col_index, column_name in enumerate(columns, start=1):
        max_length = len(column_name)
        for row_values in ws.iter_rows(
            min_row=2, max_row=min(ws.max_row, 501), min_col=col_index, max_col=col_index
        ):
            cell_value = row_values[0].value
            if cell_value is None:
                continue
            max_length = max(max_length, len(str(cell_value)))
        ws.column_dimensions[get_column_letter(col_index)].width = min(max(12, max_length + 2), 42)


def should_skip_file(path: Path, input_folder_name: str, output_folder: Path) -> tuple[bool, str]:
    if path.suffix.lower() != ".xlsx":
        return True, "not an .xlsx file"
    if path.name.startswith("~"):
        return True, "temporary Excel file"

    if path.parent.resolve() == output_folder.resolve():
        pattern = rf"^{re.escape(input_folder_name)}_PARAM(\.\d+)?\.xlsx$"
        if re.match(pattern, path.name, flags=re.IGNORECASE):
            return True, "output workbook artifact"
    return False, ""


def main() -> None:
    source_dir = Path(input_dir).expanduser().resolve()
    destination_dir = Path(output_dir).expanduser().resolve()
    destination_dir.mkdir(parents=True, exist_ok=True)

    if not source_dir.exists():
        raise SystemExit(f"Input directory does not exist: {source_dir}")

    output_path = select_output_path(source_dir, destination_dir)

    empirical_rows: list[dict[str, Any]] = []
    regression_rows: list[dict[str, Any]] = []
    processed_files = 0

    app: xw.App | None = None
    try:
        app = xw.App(visible=False, add_book=False)
        app.display_alerts = False
        app.screen_updating = False
        app.enable_events = False
        app.calculation = "manual"

        for file_path in sorted(source_dir.iterdir()):
            if not file_path.is_file():
                continue

            skip, reason = should_skip_file(file_path, source_dir.name, destination_dir)
            if skip:
                print(f"Skipped {file_path.name}: {reason}")
                continue

            print(f"Processing {file_path.name}")
            wb: xw.Book | None = None
            try:
                wb = app.books.open(str(file_path), update_links=False)
                meta = parse_file_label(file_path.name)
                empirical_rows.extend(extract_empirical_rows(wb, meta, file_path.name))
                regression_rows.extend(extract_regression_rows(wb, meta, file_path.name))
                processed_files += 1
            except Exception as exc:
                print(f"Skipped {file_path.name}: {exc}")
            finally:
                if wb is not None:
                    close_workbook_safely(wb)
    finally:
        if app is not None:
            try:
                app.quit()
            except Exception:
                pass

    out_wb = Workbook()
    default_sheet = out_wb.active
    out_wb.remove(default_sheet)
    write_sheet(out_wb, "empirical_candidates", EMPIRICAL_COLUMNS, empirical_rows)
    write_sheet(out_wb, "regression_candidates", REGRESSION_COLUMNS, regression_rows)
    out_wb.save(output_path)

    print(f"Output path: {output_path}")
    print(f"Number of files processed: {processed_files}")
    print(f"Number of empirical rows: {len(empirical_rows)}")
    print(f"Number of regression rows: {len(regression_rows)}")


if __name__ == "__main__":
    main()
