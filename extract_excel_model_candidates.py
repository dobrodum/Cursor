#!/usr/bin/env python3
from __future__ import annotations

import re
from datetime import date
from pathlib import Path
from typing import Any

import xlwings as xw
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

# Configure these two paths before running.
input_dir = Path("/workspace/input")
output_dir = Path("/workspace/output")

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

DAY_BY_PERIOD = {
    "early": 5,
    "mid": 15,
    "late": 25,
}

MONTH_BY_TOKEN = {
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


def normalize_label(value: Any) -> str:
    if not isinstance(value, str):
        return ""
    return re.sub(r"[^a-z0-9]+", "", value.strip().lower())


def as_float(value: Any) -> float | None:
    if value is None or value == "":
        return None
    if isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        return float(value)
    try:
        cleaned = str(value).replace(",", "").strip()
        if not cleaned:
            return None
        return float(cleaned)
    except (TypeError, ValueError):
        return None


def as_int(value: Any) -> int | None:
    numeric = as_float(value)
    if numeric is None:
        return None
    return int(round(numeric))


def normalize_matrix(values: Any) -> list[list[Any]]:
    if values is None:
        return []
    if not isinstance(values, list):
        return [[values]]
    if not values:
        return []
    if isinstance(values[0], list):
        return values
    return [values]


def parse_model_metadata(file_path: Path) -> dict[str, str]:
    stem = file_path.stem
    parts = [part.strip() for part in stem.split(" - ") if part.strip()]

    ticker = ""
    if len(parts) >= 2:
        ticker = parts[1]
    if not ticker:
        ticker_match = re.search(r"\b([A-Z]{2,8})\b", stem)
        ticker = ticker_match.group(1) if ticker_match else "UNKNOWN"
    ticker = ticker.strip().upper()

    period_match = re.search(r"(Early|Mid|Late)([A-Za-z]+)(\d{4})", stem, flags=re.IGNORECASE)
    if not period_match:
        raise ValueError(f"Unable to parse period label from filename: {file_path.name}")

    period_prefix = period_match.group(1).capitalize()
    month_token_raw = period_match.group(2)
    year = int(period_match.group(3))

    month_key = month_token_raw.lower()
    month = MONTH_BY_TOKEN.get(month_key) or MONTH_BY_TOKEN.get(month_key[:3])
    if month is None:
        raise ValueError(f"Unsupported month token '{month_token_raw}' in filename: {file_path.name}")

    month_abbrev = date(year, month, 1).strftime("%b")
    model_period = f"{period_prefix}{month_abbrev}_{year}"
    day = DAY_BY_PERIOD[period_prefix.lower()]
    model_date = date(year, month, day).isoformat()

    return {
        "ticker": ticker,
        "model_period": model_period,
        "model_date": model_date,
        "model": f"{ticker}_{model_period}",
    }


def build_output_path(input_folder: Path, output_folder: Path) -> Path:
    base = f"{input_folder.name}_PARAM"
    candidate = output_folder / f"{base}.xlsx"
    counter = 1
    while candidate.exists():
        candidate = output_folder / f"{base}.{counter}.xlsx"
        counter += 1
    return candidate


def close_workbook_no_save(workbook: xw.Book) -> None:
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
    except Exception:
        try:
            workbook.close()
        except Exception:
            pass


def read_sheet_matrix(sheet: xw.Sheet) -> tuple[list[list[Any]], int, int]:
    used = sheet.used_range
    values = normalize_matrix(used.value)
    return values, used.row, used.column


def find_max_anchor(values: list[list[Any]], start_row: int, start_col: int) -> tuple[int, int] | None:
    for r_index, row in enumerate(values):
        for c_index, value in enumerate(row):
            if isinstance(value, str) and value.strip().lower() == "max":
                return start_row + r_index, start_col + c_index
    return None


def get_value(
    sheet: xw.Sheet,
    values: list[list[Any]],
    start_row: int,
    start_col: int,
    row: int,
    col: int,
) -> Any:
    r_index = row - start_row
    c_index = col - start_col
    if 0 <= r_index < len(values):
        row_values = values[r_index]
        if 0 <= c_index < len(row_values):
            return row_values[c_index]
    return sheet.range((row, col)).value


def detect_scenario_columns(
    values: list[list[Any]],
    start_row: int,
    start_col: int,
    anchor_row: int,
    anchor_col: int,
    n_quarters: int,
) -> list[int]:
    row_index = anchor_row - start_row
    if row_index < 0 or row_index >= len(values):
        return [anchor_col + idx + 1 for idx in range(n_quarters)]

    row_values = values[row_index]
    right_numeric: list[int] = []
    left_numeric: list[int] = []

    for c_index, value in enumerate(row_values):
        abs_col = start_col + c_index
        if as_float(value) is None:
            continue
        if abs_col > anchor_col:
            right_numeric.append(abs_col)
        elif abs_col < anchor_col:
            left_numeric.append(abs_col)

    if right_numeric or left_numeric:
        if len(right_numeric) >= len(left_numeric) and right_numeric:
            cols = sorted(right_numeric)[:n_quarters]
        else:
            cols = sorted(left_numeric)[-n_quarters:]
        if len(cols) < n_quarters:
            last_col = cols[-1] if cols else anchor_col
            cols.extend([last_col + idx + 1 for idx in range(n_quarters - len(cols))])
        return cols

    return [anchor_col + idx + 1 for idx in range(n_quarters)]


def collect_text_cells(values: list[list[Any]], start_row: int, start_col: int) -> list[tuple[str, int, int]]:
    labels: list[tuple[str, int, int]] = []
    for r_index, row in enumerate(values):
        for c_index, value in enumerate(row):
            normalized = normalize_label(value)
            if normalized:
                labels.append((normalized, start_row + r_index, start_col + c_index))
    return labels


def find_row_by_label_patterns(
    labels: list[tuple[str, int, int]],
    include_any: tuple[tuple[str, ...], ...],
    exclude: tuple[str, ...] = (),
) -> int | None:
    for normalized, row, _col in labels:
        if exclude and any(token in normalized for token in exclude):
            continue
        for token_group in include_any:
            if all(token in normalized for token in token_group):
                return row
    return None


def get_sheet_by_name(workbook: xw.Book, expected_name: str) -> xw.Sheet | None:
    expected = expected_name.strip().lower()
    for sheet in workbook.sheets:
        if sheet.name.strip().lower() == expected:
            return sheet
    return None


def process_empirical_sheet(
    sheet: xw.Sheet,
    workbook: xw.Book,
    metadata: dict[str, str],
    source_file: str,
) -> list[dict[str, Any]]:
    values, start_row, start_col = read_sheet_matrix(sheet)
    anchor = find_max_anchor(values, start_row, start_col)
    if anchor is None:
        print(f"  empirical skipped: no 'max' anchor in sheet '{sheet.name}'")
        return []

    anchor_row, anchor_col = anchor
    scenario_cols = detect_scenario_columns(values, start_row, start_col, anchor_row, anchor_col, n_quarters=10)
    labels = collect_text_cells(values, start_row, start_col)

    num_quarters_row = find_row_by_label_patterns(labels, include_any=(("num", "quarter"), ("quarter", "used")))
    last_quarter_row = find_row_by_label_patterns(labels, include_any=(("last", "quarter"),))
    forecast_value_row = find_row_by_label_patterns(
        labels,
        include_any=(("estimated", "sold"), ("tot", "fcst"), ("forecast", "total")),
    )
    actual_value_row = find_row_by_label_patterns(labels, include_any=(("reported", "sales"),))
    quarterly_sales_row = find_row_by_label_patterns(labels, include_any=(("quarterly", "sales"),))
    reported_sales_row = actual_value_row
    growth_rate_row = find_row_by_label_patterns(labels, include_any=(("growth", "rate"), ("growth",)))
    captured_pct_row = find_row_by_label_patterns(
        labels,
        include_any=(("captured", "db"), ("sales", "captured")),
    )
    penetration_row = find_row_by_label_patterns(
        labels,
        include_any=(("penetration",),),
        exclude=("avg",),
    )
    avg_penetration_row = find_row_by_label_patterns(
        labels,
        include_any=(("avg", "penetration"),),
    )

    if num_quarters_row is None:
        num_quarters_row = anchor_row - 4
    if last_quarter_row is None:
        last_quarter_row = anchor_row - 3
    if forecast_value_row is None:
        forecast_value_row = anchor_row - 2
    if actual_value_row is None:
        actual_value_row = anchor_row - 1
    if quarterly_sales_row is None:
        quarterly_sales_row = anchor_row - 6
    if growth_rate_row is None:
        growth_rate_row = anchor_row - 5
    if captured_pct_row is None:
        captured_pct_row = anchor_row - 7

    rows: list[dict[str, Any]] = []
    avg_formula_cells: list[tuple[int, int, int, int]] = []

    scratch_row = max(anchor_row + 3, start_row + len(values) + 3)
    scratch_col_start = max(scenario_cols) + 2
    first_scenario_col = min(scenario_cols)

    for index, scenario_col in enumerate(scenario_cols):
        num_quarters_used = as_int(get_value(sheet, values, start_row, start_col, num_quarters_row, scenario_col))
        if not num_quarters_used or num_quarters_used < 1:
            num_quarters_used = index + 1

        forecast_max = as_float(get_value(sheet, values, start_row, start_col, anchor_row, scenario_col))
        forecast_min = as_float(get_value(sheet, values, start_row, start_col, anchor_row + 1, scenario_col))
        forecast_value = as_float(get_value(sheet, values, start_row, start_col, forecast_value_row, scenario_col))
        actual_value = as_float(get_value(sheet, values, start_row, start_col, actual_value_row, scenario_col))
        quarterly_sales = as_float(get_value(sheet, values, start_row, start_col, quarterly_sales_row, scenario_col))
        reported_sales = as_float(get_value(sheet, values, start_row, start_col, reported_sales_row, scenario_col))
        growth_rate_pct = as_float(get_value(sheet, values, start_row, start_col, growth_rate_row, scenario_col))
        sales_captured_pct = as_float(get_value(sheet, values, start_row, start_col, captured_pct_row, scenario_col))
        last_quarter_used = get_value(sheet, values, start_row, start_col, last_quarter_row, scenario_col)

        row_data = {
            "model": metadata["model"],
            "ticker": metadata["ticker"],
            "model_period": metadata["model_period"],
            "model_date": metadata["model_date"],
            "method": "empirical",
            "parameter_name": "avg_penetration_pct",
            "parameter_value": None,
            "num_quarters_used": num_quarters_used,
            "last_quarter_used": last_quarter_used,
            "forecast_value": forecast_value,
            "actual_value": actual_value,
            "forecast_max": forecast_max,
            "forecast_min": forecast_min,
            "range_width": (forecast_max - forecast_min) if forecast_max is not None and forecast_min is not None else None,
            "avg_penetration_pct": None,
            "quarterly_sales": quarterly_sales,
            "reported_sales": reported_sales,
            "growth_rate_pct": growth_rate_pct,
            "sales_captured_in_db_pct": sales_captured_pct,
            "source_file": source_file,
        }

        if penetration_row is not None:
            avg_start_col = max(first_scenario_col, scenario_col - num_quarters_used + 1)
            avg_formula_cells.append((scratch_row, scratch_col_start + index, avg_start_col, scenario_col))
        elif avg_penetration_row is not None:
            avg_pen = as_float(get_value(sheet, values, start_row, start_col, avg_penetration_row, scenario_col))
            row_data["avg_penetration_pct"] = avg_pen
            row_data["parameter_value"] = avg_pen

        rows.append(row_data)

    if avg_formula_cells and penetration_row is not None:
        for formula_row, formula_col, avg_start_col, avg_end_col in avg_formula_cells:
            formula_cell = sheet.range((formula_row, formula_col))
            formula_cell.formula2 = f"=AVERAGE(R{penetration_row}C{avg_start_col}:R{penetration_row}C{avg_end_col})"

        workbook.app.calculate()

        values_range = sheet.range(
            (scratch_row, scratch_col_start),
            (scratch_row, scratch_col_start + len(avg_formula_cells) - 1),
        ).value
        avg_results = values_range if isinstance(values_range, list) else [values_range]

        for idx, avg_value in enumerate(avg_results):
            avg_numeric = as_float(avg_value)
            rows[idx]["avg_penetration_pct"] = avg_numeric
            rows[idx]["parameter_value"] = avg_numeric

        sheet.range(
            (scratch_row, scratch_col_start),
            (scratch_row, scratch_col_start + len(avg_formula_cells) - 1),
        ).clear_contents()

    for row_data in rows:
        if row_data["parameter_value"] is None:
            row_data["parameter_value"] = row_data["avg_penetration_pct"]

    return rows


def process_regression_sheet(
    sheet: xw.Sheet,
    workbook: xw.Book,
    metadata: dict[str, str],
    source_file: str,
) -> list[dict[str, Any]]:
    values, start_row, start_col = read_sheet_matrix(sheet)
    anchor = find_max_anchor(values, start_row, start_col)
    if anchor is None:
        print(f"  regression skipped: no 'max' anchor in sheet '{sheet.name}'")
        return []

    anchor_row, anchor_col = anchor
    y_col = anchor_col - 7
    x_col = anchor_col - 11
    scenario_cols = detect_scenario_columns(values, start_row, start_col, anchor_row, anchor_col, n_quarters=10)
    labels = collect_text_cells(values, start_row, start_col)

    num_quarters_row = find_row_by_label_patterns(labels, include_any=(("num", "quarter"), ("quarter", "used")))
    forecast_value_row = find_row_by_label_patterns(
        labels,
        include_any=(("tot", "fcst", "wosa"), ("total", "forecast"), ("fcst", "wosa")),
    )
    actual_row = find_row_by_label_patterns(labels, include_any=(("actual",), ("reported", "sales")))

    if num_quarters_row is None:
        num_quarters_row = anchor_row - 3
    if forecast_value_row is None:
        forecast_value_row = anchor_row - 2

    xy_pairs: list[tuple[int, float, float]] = []
    for r_index, _row_values in enumerate(values):
        row_number = start_row + r_index
        if row_number >= anchor_row:
            continue
        x_val = as_float(get_value(sheet, values, start_row, start_col, row_number, x_col))
        y_val = as_float(get_value(sheet, values, start_row, start_col, row_number, y_col))
        if x_val is not None and y_val is not None:
            xy_pairs.append((row_number, x_val, y_val))

    rows: list[dict[str, Any]] = []
    scratch_row_intercept = max(anchor_row + 3, start_row + len(values) + 3)
    scratch_row_slope = scratch_row_intercept + 1
    scratch_col_start = max(scenario_cols) + 2

    formula_targets: list[tuple[int, int, int]] = []

    for index, scenario_col in enumerate(scenario_cols):
        num_quarters_used = as_int(get_value(sheet, values, start_row, start_col, num_quarters_row, scenario_col))
        if not num_quarters_used or num_quarters_used < 1:
            num_quarters_used = index + 1

        forecast_max = as_float(get_value(sheet, values, start_row, start_col, anchor_row, scenario_col))
        forecast_min = as_float(get_value(sheet, values, start_row, start_col, anchor_row + 1, scenario_col))
        forecast_value = as_float(get_value(sheet, values, start_row, start_col, forecast_value_row, scenario_col))
        actual_value = (
            as_float(get_value(sheet, values, start_row, start_col, actual_row, scenario_col))
            if actual_row is not None
            else None
        )

        if xy_pairs:
            effective_quarters = min(num_quarters_used, len(xy_pairs))
            first_row = xy_pairs[-effective_quarters][0]
            last_row = xy_pairs[-1][0]
            formula_targets.append((index, first_row, last_row))

        row_data = {
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
            "range_width": (forecast_max - forecast_min) if forecast_max is not None and forecast_min is not None else None,
            "intercept": None,
            "slope": None,
            "source_file": source_file,
        }
        rows.append(row_data)

    if formula_targets:
        for index, first_row, last_row in formula_targets:
            col = scratch_col_start + index
            intercept_cell = sheet.range((scratch_row_intercept, col))
            slope_cell = sheet.range((scratch_row_slope, col))
            intercept_cell.formula2 = (
                f"=INTERCEPT(R{first_row}C{y_col}:R{last_row}C{y_col},R{first_row}C{x_col}:R{last_row}C{x_col})"
            )
            slope_cell.formula2 = f"=SLOPE(R{first_row}C{y_col}:R{last_row}C{y_col},R{first_row}C{x_col}:R{last_row}C{x_col})"

        workbook.app.calculate()

        intercept_values = sheet.range(
            (scratch_row_intercept, scratch_col_start),
            (scratch_row_intercept, scratch_col_start + len(formula_targets) - 1),
        ).value
        slope_values = sheet.range(
            (scratch_row_slope, scratch_col_start),
            (scratch_row_slope, scratch_col_start + len(formula_targets) - 1),
        ).value

        intercept_list = intercept_values if isinstance(intercept_values, list) else [intercept_values]
        slope_list = slope_values if isinstance(slope_values, list) else [slope_values]

        for idx in range(len(rows)):
            rows[idx]["intercept"] = as_float(intercept_list[idx]) if idx < len(intercept_list) else None
            rows[idx]["slope"] = as_float(slope_list[idx]) if idx < len(slope_list) else None

        sheet.range(
            (scratch_row_intercept, scratch_col_start),
            (scratch_row_slope, scratch_col_start + len(formula_targets) - 1),
        ).clear_contents()

    deduped_rows: list[dict[str, Any]] = []
    previous_signature: tuple[Any, ...] | None = None
    for row_data in rows:
        signature = (
            row_data["num_quarters_used"],
            row_data["forecast_value"],
            row_data["forecast_max"],
            row_data["forecast_min"],
            row_data["intercept"],
            row_data["slope"],
        )
        if signature == previous_signature:
            continue
        deduped_rows.append(row_data)
        previous_signature = signature

    return deduped_rows


def write_sheet(
    ws: Any,
    columns: list[str],
    rows: list[dict[str, Any]],
) -> None:
    ws.append(columns)
    for row_data in rows:
        ws.append([row_data.get(column) for column in columns])

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(columns))}{max(1, ws.max_row)}"

    for cell in ws[1]:
        cell.font = Font(bold=True)

    for col_idx, column_name in enumerate(columns, start=1):
        max_len = len(column_name)
        for row_idx in range(2, ws.max_row + 1):
            value = ws.cell(row=row_idx, column=col_idx).value
            value_len = len(str(value)) if value is not None else 0
            if value_len > max_len:
                max_len = value_len
        ws.column_dimensions[get_column_letter(col_idx)].width = min(max_len + 2, 42)


def write_output_workbook(
    output_path: Path,
    empirical_rows: list[dict[str, Any]],
    regression_rows: list[dict[str, Any]],
) -> None:
    workbook = Workbook()
    empirical_ws = workbook.active
    empirical_ws.title = "empirical_candidates"
    regression_ws = workbook.create_sheet("regression_candidates")

    write_sheet(empirical_ws, EMPIRICAL_COLUMNS, empirical_rows)
    write_sheet(regression_ws, REGRESSION_COLUMNS, regression_rows)

    workbook.save(output_path)


def should_skip_file(file_path: Path, input_folder: Path) -> tuple[bool, str]:
    if not file_path.is_file():
        return True, "not a file"
    if file_path.name.startswith("~"):
        return True, "temp file"
    if file_path.suffix.lower() != ".xlsx":
        return True, "not an .xlsx file"
    if re.search(rf"^{re.escape(input_folder.name)}_PARAM(\.\d+)?\.xlsx$", file_path.name, flags=re.IGNORECASE):
        return True, "generated output file"
    return False, ""


def main() -> None:
    in_dir = Path(input_dir)
    out_dir = Path(output_dir)

    if not in_dir.exists():
        raise FileNotFoundError(f"Input directory not found: {in_dir}")
    out_dir.mkdir(parents=True, exist_ok=True)

    output_path = build_output_path(in_dir, out_dir)

    empirical_rows: list[dict[str, Any]] = []
    regression_rows: list[dict[str, Any]] = []
    processed_file_count = 0

    app = xw.App(visible=False, add_book=False)
    app.display_alerts = False
    app.screen_updating = False

    try:
        for file_path in sorted(in_dir.iterdir()):
            skip, reason = should_skip_file(file_path, in_dir)
            if skip:
                print(f"Skipped: {file_path.name} ({reason})")
                continue

            print(f"Processing: {file_path.name}")
            try:
                metadata = parse_model_metadata(file_path)
            except Exception as exc:
                print(f"Skipped: {file_path.name} (filename parse failed: {exc})")
                continue

            workbook = None
            try:
                workbook = app.books.open(str(file_path), update_links=False)
                empirical_sheet = get_sheet_by_name(workbook, "Empirical Model")
                regression_sheet = get_sheet_by_name(workbook, "Regression Model")

                if empirical_sheet is not None:
                    empirical_rows.extend(
                        process_empirical_sheet(
                            empirical_sheet,
                            workbook,
                            metadata,
                            source_file=file_path.name,
                        )
                    )
                else:
                    print("  empirical skipped: missing sheet 'Empirical Model'")

                if regression_sheet is not None:
                    regression_rows.extend(
                        process_regression_sheet(
                            regression_sheet,
                            workbook,
                            metadata,
                            source_file=file_path.name,
                        )
                    )
                else:
                    print("  regression skipped: missing sheet 'Regression Model'")

                processed_file_count += 1
            except Exception as exc:
                print(f"Skipped: {file_path.name} (processing failed: {exc})")
            finally:
                if workbook is not None:
                    close_workbook_no_save(workbook)
    finally:
        app.quit()

    write_output_workbook(output_path, empirical_rows, regression_rows)

    print(f"Output workbook: {output_path}")
    print(f"Files processed: {processed_file_count}")
    print(f"Empirical rows: {len(empirical_rows)}")
    print(f"Regression rows: {len(regression_rows)}")


if __name__ == "__main__":
    main()
