#!/usr/bin/env python3
from __future__ import annotations

import re
from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path
from typing import Any, Optional

import xlwings as xw

# ---------- User inputs ----------
input_dir = r"/path/to/input"
output_dir = r"/path/to/output"
# -------------------------------

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

PHASE_TO_DAY = {"early": 5, "mid": 15, "late": 25}


@dataclass(frozen=True)
class SheetSnapshot:
    values: list[list[Any]]
    base_row: int
    base_col: int
    n_rows: int
    n_cols: int


def to_2d(values: Any) -> list[list[Any]]:
    if values is None:
        return []
    if isinstance(values, list):
        if not values:
            return []
        if isinstance(values[0], list):
            row_lengths = [len(row) for row in values]
            max_len = max(row_lengths) if row_lengths else 0
            if max_len == 0:
                return []
            return [row + [None] * (max_len - len(row)) for row in values]
        return [values]
    return [[values]]


def snapshot_sheet(sheet: xw.Sheet, max_rows: int = 5000, max_cols: int = 256) -> SheetSnapshot:
    used = sheet.used_range
    base_row = used.row
    base_col = used.column
    rows_count = int(used.rows.count)
    cols_count = int(used.columns.count)

    if rows_count <= 0 or cols_count <= 0:
        return SheetSnapshot(values=[], base_row=1, base_col=1, n_rows=0, n_cols=0)

    read_rows = min(rows_count, max_rows)
    read_cols = min(cols_count, max_cols)
    bottom_row = base_row + read_rows - 1
    right_col = base_col + read_cols - 1
    values = to_2d(sheet.range((base_row, base_col), (bottom_row, right_col)).value)

    if not values:
        return SheetSnapshot(values=[], base_row=base_row, base_col=base_col, n_rows=0, n_cols=0)

    return SheetSnapshot(
        values=values,
        base_row=base_row,
        base_col=base_col,
        n_rows=len(values),
        n_cols=len(values[0]),
    )


def snapshot_value(snapshot: SheetSnapshot, row: int, col: int) -> Any:
    row_idx = row - snapshot.base_row
    col_idx = col - snapshot.base_col
    if row_idx < 0 or col_idx < 0:
        return None
    if row_idx >= snapshot.n_rows or col_idx >= snapshot.n_cols:
        return None
    return snapshot.values[row_idx][col_idx]


def normalize_label(value: Any) -> str:
    if not isinstance(value, str):
        return ""
    normalized = value.strip().lower().replace("_", " ")
    return re.sub(r"\s+", " ", normalized)


def to_float(value: Any) -> Optional[float]:
    if value is None or value == "" or isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        return float(value)
    if isinstance(value, str):
        cleaned = value.strip().replace(",", "")
        if cleaned.endswith("%"):
            cleaned = cleaned[:-1].strip()
            try:
                return float(cleaned) / 100.0
            except ValueError:
                return None
        try:
            return float(cleaned)
        except ValueError:
            return None
    return None


def find_anchor(snapshot: SheetSnapshot, anchor_text: str = "max") -> Optional[tuple[int, int]]:
    target = anchor_text.strip().lower()
    for r in range(snapshot.n_rows):
        for c in range(snapshot.n_cols):
            value = snapshot.values[r][c]
            if isinstance(value, str) and value.strip().lower() == target:
                return snapshot.base_row + r, snapshot.base_col + c
    return None


def build_label_entries(
    snapshot: SheetSnapshot,
    anchor_row: int,
    anchor_col: int,
    row_radius: int = 120,
    col_radius: int = 40,
) -> list[tuple[str, int, int, int]]:
    labels: list[tuple[str, int, int, int]] = []
    min_row = max(snapshot.base_row, anchor_row - row_radius)
    max_row = min(snapshot.base_row + snapshot.n_rows - 1, anchor_row + row_radius)
    min_col = max(snapshot.base_col, anchor_col - col_radius)
    max_col = min(snapshot.base_col + snapshot.n_cols - 1, anchor_col + col_radius)

    for row in range(min_row, max_row + 1):
        for col in range(min_col, max_col + 1):
            text = normalize_label(snapshot_value(snapshot, row, col))
            if not text:
                continue
            distance = abs(row - anchor_row) + abs(col - anchor_col)
            labels.append((text, row, col, distance))
    return labels


def locate_label(
    labels: list[tuple[str, int, int, int]],
    patterns: list[str],
) -> Optional[tuple[int, int, str]]:
    if not labels:
        return None
    lowered = [p.lower() for p in patterns]
    matches = [entry for entry in labels if any(p in entry[0] for p in lowered)]
    if not matches:
        return None
    matches.sort(key=lambda item: (item[3], len(item[0])))
    text, row, col, _ = matches[0]
    return row, col, text


def pick_value_cell(snapshot: SheetSnapshot, label_row: int, label_col: int) -> tuple[int, int]:
    candidates = [
        (label_row, label_col + 1),
        (label_row, label_col + 2),
        (label_row + 1, label_col),
        (label_row + 1, label_col + 1),
    ]
    for row, col in candidates:
        value = snapshot_value(snapshot, row, col)
        if value not in (None, ""):
            return row, col
    return label_row, label_col + 1


def resolve_cell_from_label(
    snapshot: SheetSnapshot,
    labels: list[tuple[str, int, int, int]],
    patterns: list[str],
) -> Optional[tuple[int, int]]:
    found = locate_label(labels, patterns)
    if not found:
        return None
    label_row, label_col, _ = found
    return pick_value_cell(snapshot, label_row, label_col)


def sheet_value(sheet: xw.Sheet, coord: Optional[tuple[int, int]]) -> Any:
    if coord is None:
        return None
    return sheet.range(coord).value


def collect_numeric_rows(
    snapshot: SheetSnapshot,
    col: int,
    start_row: int,
    end_row: int,
) -> list[int]:
    if start_row > end_row:
        return []
    rows: list[int] = []
    for row in range(start_row, end_row + 1):
        if to_float(snapshot_value(snapshot, row, col)) is not None:
            rows.append(row)
    return rows


def latest_contiguous_block(rows: list[int]) -> list[int]:
    if not rows:
        return []
    block = [rows[0]]
    for row in rows[1:]:
        if row == block[-1] + 1:
            block.append(row)
        else:
            block = [row]
    return block


def safe_close_workbook(wb: xw.Book) -> None:
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


def parse_file_label(file_path: Path) -> dict[str, str]:
    parts = [part.strip() for part in file_path.stem.split(" - ")]
    if len(parts) < 3:
        raise ValueError("expected at least three filename sections split by ' - '")

    ticker = parts[1]
    period_chunk = parts[2]
    match = re.search(r"(Early|Mid|Late)\s*([A-Za-z]{3,9})\s*(\d{4})", period_chunk, flags=re.IGNORECASE)
    if not match:
        raise ValueError("missing Early/Mid/Late + month + year pattern")

    phase = match.group(1).title()
    month_token = match.group(2).strip()
    year = int(match.group(3))

    month_abbr = datetime.strptime(month_token[:3].title(), "%b").strftime("%b")
    month_num = datetime.strptime(month_abbr, "%b").month
    day = PHASE_TO_DAY[phase.lower()]

    model_period = f"{phase}{month_abbr}_{year}"
    model_date = date(year, month_num, day).isoformat()
    model = f"{ticker}_{model_period}"

    return {
        "model": model,
        "ticker": ticker,
        "model_period": model_period,
        "model_date": model_date,
        "source_file": file_path.name,
    }


def next_output_path(input_folder: Path, output_folder: Path) -> Path:
    base_name = f"{input_folder.name}_PARAM"
    candidate = output_folder / f"{base_name}.xlsx"
    counter = 1
    while candidate.exists():
        candidate = output_folder / f"{base_name}.{counter}.xlsx"
        counter += 1
    return candidate


def process_empirical_sheet(
    wb: xw.Book,
    meta: dict[str, str],
    n_quarters: int = 10,
) -> tuple[list[dict[str, Any]], Optional[str]]:
    try:
        sheet = wb.sheets["Empirical Model"]
    except Exception:
        return [], "Empirical Model sheet not found"

    snapshot = snapshot_sheet(sheet)
    anchor = find_anchor(snapshot, "max")
    if anchor is None:
        return [], "Empirical Model max anchor not found"

    anchor_row, anchor_col = anchor
    labels = build_label_entries(snapshot, anchor_row, anchor_col)
    right_of_anchor = (anchor_row, anchor_col + 1)

    min_label = locate_label(labels, ["min"])
    min_coord = pick_value_cell(snapshot, min_label[0], min_label[1]) if min_label else (anchor_row + 1, anchor_col + 1)

    num_quarters_coord = resolve_cell_from_label(snapshot, labels, ["num quarters used", "num quarters"])
    last_quarter_coord = resolve_cell_from_label(snapshot, labels, ["last quarter used", "last quarter", "last qtr"])
    forecast_coord = resolve_cell_from_label(
        snapshot,
        labels,
        ["estimated total sold", "total sold estimate", "tot fcst", "forecast total"],
    )
    reported_sales_coord = resolve_cell_from_label(snapshot, labels, ["reported sales", "actual sales"])
    quarterly_sales_coord = resolve_cell_from_label(snapshot, labels, ["quarterly sales", "sales in db", "db sales"])
    growth_rate_coord = resolve_cell_from_label(snapshot, labels, ["growth rate", "growth"])
    captured_pct_coord = resolve_cell_from_label(snapshot, labels, ["captured in db", "sales captured"])
    avg_penetration_output_coord = resolve_cell_from_label(snapshot, labels, ["avg penetration", "average penetration"])

    penetration_label = locate_label(labels, ["penetration"])
    candidate_penetration_cols: list[int] = []
    data_start_row = snapshot.base_row
    if penetration_label:
        label_row, label_col, _ = penetration_label
        data_start_row = label_row + 1
        candidate_penetration_cols.extend([label_col, label_col + 1])

    candidate_penetration_cols.extend([anchor_col - 6, anchor_col - 5, anchor_col - 4, anchor_col - 3])
    candidate_penetration_cols = [col for col in candidate_penetration_cols if col >= snapshot.base_col]

    best_penetration_rows: list[int] = []
    penetration_col: Optional[int] = None
    search_start = max(snapshot.base_row, data_start_row)
    search_end = anchor_row - 1
    for col in candidate_penetration_cols:
        rows = collect_numeric_rows(snapshot, col, search_start, search_end)
        if len(rows) > len(best_penetration_rows):
            best_penetration_rows = rows
            penetration_col = col

    temp_col = max(anchor_col + 25, snapshot.base_col + snapshot.n_cols + 2)
    avg_pen_temp_cell = sheet.range((anchor_row, temp_col))

    rows_out: list[dict[str, Any]] = []
    for quarter_count in range(1, n_quarters + 1):
        changed = False

        if num_quarters_coord:
            sheet.range(num_quarters_coord).value = quarter_count
            changed = True

        avg_penetration_pct = None
        if penetration_col is not None and best_penetration_rows:
            use_count = min(quarter_count, len(best_penetration_rows))
            selected_rows = best_penetration_rows[-use_count:]
            start_row = selected_rows[0]
            end_row = selected_rows[-1]
            avg_pen_temp_cell.formula2 = f"=AVERAGE(R{start_row}C{penetration_col}:R{end_row}C{penetration_col})"
            changed = True

        if changed:
            wb.app.calculate()

        if penetration_col is not None and best_penetration_rows:
            avg_penetration_pct = to_float(avg_pen_temp_cell.value)
        elif avg_penetration_output_coord:
            avg_penetration_pct = to_float(sheet_value(sheet, avg_penetration_output_coord))

        forecast_value = to_float(sheet_value(sheet, forecast_coord))
        reported_sales = to_float(sheet_value(sheet, reported_sales_coord))
        actual_value = reported_sales
        forecast_max = to_float(sheet_value(sheet, right_of_anchor))
        forecast_min = to_float(sheet_value(sheet, min_coord))
        quarterly_sales = to_float(sheet_value(sheet, quarterly_sales_coord))
        growth_rate_pct = to_float(sheet_value(sheet, growth_rate_coord))
        sales_captured_in_db_pct = to_float(sheet_value(sheet, captured_pct_coord))
        last_quarter_used = sheet_value(sheet, last_quarter_coord)

        if forecast_value is None and quarterly_sales is not None and avg_penetration_pct not in (None, 0):
            forecast_value = quarterly_sales / avg_penetration_pct

        range_width = None
        if forecast_max is not None and forecast_min is not None:
            range_width = forecast_max - forecast_min

        if all(
            value is None
            for value in [
                forecast_value,
                actual_value,
                forecast_max,
                forecast_min,
                avg_penetration_pct,
                quarterly_sales,
            ]
        ):
            continue

        rows_out.append(
            {
                "model": meta["model"],
                "ticker": meta["ticker"],
                "model_period": meta["model_period"],
                "model_date": meta["model_date"],
                "method": "empirical",
                "parameter_name": "avg_penetration_pct",
                "parameter_value": avg_penetration_pct,
                "num_quarters_used": quarter_count,
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
                "source_file": meta["source_file"],
            }
        )

    return rows_out, None


def process_regression_sheet(
    wb: xw.Book,
    meta: dict[str, str],
    n_quarters: int = 10,
) -> tuple[list[dict[str, Any]], Optional[str]]:
    try:
        sheet = wb.sheets["Regression Model"]
    except Exception:
        return [], "Regression Model sheet not found"

    snapshot = snapshot_sheet(sheet)
    anchor = find_anchor(snapshot, "max")
    if anchor is None:
        return [], "Regression Model max anchor not found"

    anchor_row, anchor_col = anchor
    labels = build_label_entries(snapshot, anchor_row, anchor_col)

    y_col = anchor_col - 7
    x_col = anchor_col - 11
    search_start = max(snapshot.base_row, anchor_row - 500)
    search_end = anchor_row - 1

    candidate_rows: list[int] = []
    for row in range(search_start, search_end + 1):
        x_val = to_float(snapshot_value(snapshot, row, x_col))
        y_val = to_float(snapshot_value(snapshot, row, y_col))
        if x_val is not None and y_val is not None:
            candidate_rows.append(row)

    rows_block = latest_contiguous_block(candidate_rows)
    if len(rows_block) < 2:
        return [], "Regression Model lacks enough contiguous x/y rows"

    right_of_anchor = (anchor_row, anchor_col + 1)
    min_label = locate_label(labels, ["min"])
    min_coord = pick_value_cell(snapshot, min_label[0], min_label[1]) if min_label else (anchor_row + 1, anchor_col + 1)

    num_quarters_coord = resolve_cell_from_label(snapshot, labels, ["num quarters used", "num quarters"])
    forecast_coord = resolve_cell_from_label(
        snapshot,
        labels,
        ["tot fcst w/o sa", "total fcst w/o sa", "tot fcst without sa", "tot fcst wo sa"],
    )
    actual_coord = resolve_cell_from_label(snapshot, labels, ["reported sales", "actual sales"])

    temp_col = max(anchor_col + 25, snapshot.base_col + snapshot.n_cols + 2)
    intercept_cell = sheet.range((anchor_row + 1, temp_col))
    slope_cell = sheet.range((anchor_row + 2, temp_col))

    max_count = min(n_quarters, len(rows_block))
    rows_out: list[dict[str, Any]] = []
    prev_signature: Optional[tuple[Any, ...]] = None

    for quarter_count in range(2, max_count + 1):
        selected_rows = rows_block[-quarter_count:]
        start_row = selected_rows[0]
        end_row = selected_rows[-1]

        intercept_cell.formula2 = (
            f"=INTERCEPT(R{start_row}C{y_col}:R{end_row}C{y_col},R{start_row}C{x_col}:R{end_row}C{x_col})"
        )
        slope_cell.formula2 = f"=SLOPE(R{start_row}C{y_col}:R{end_row}C{y_col},R{start_row}C{x_col}:R{end_row}C{x_col})"

        if num_quarters_coord:
            sheet.range(num_quarters_coord).value = quarter_count

        wb.app.calculate()

        intercept = to_float(intercept_cell.value)
        slope = to_float(slope_cell.value)
        forecast_value = to_float(sheet_value(sheet, forecast_coord))
        actual_value = to_float(sheet_value(sheet, actual_coord))
        forecast_max = to_float(sheet_value(sheet, right_of_anchor))
        forecast_min = to_float(sheet_value(sheet, min_coord))

        if forecast_value is None and intercept is not None and slope is not None:
            last_x_value = to_float(sheet.range((end_row, x_col)).value)
            if last_x_value is not None:
                forecast_value = intercept + slope * (last_x_value + 1)

        range_width = None
        if forecast_max is not None and forecast_min is not None:
            range_width = forecast_max - forecast_min

        signature = (
            quarter_count,
            round(intercept, 10) if intercept is not None else None,
            round(slope, 10) if slope is not None else None,
            round(forecast_value, 10) if forecast_value is not None else None,
            round(forecast_max, 10) if forecast_max is not None else None,
            round(forecast_min, 10) if forecast_min is not None else None,
        )
        if signature == prev_signature:
            continue
        prev_signature = signature

        if all(value is None for value in [intercept, slope, forecast_value, forecast_max, forecast_min]):
            continue

        rows_out.append(
            {
                "model": meta["model"],
                "ticker": meta["ticker"],
                "model_period": meta["model_period"],
                "model_date": meta["model_date"],
                "method": "regression",
                "parameter_name": "num_quarters_used",
                "parameter_value": quarter_count,
                "num_quarters_used": quarter_count,
                "forecast_value": forecast_value,
                "actual_value": actual_value,
                "forecast_max": forecast_max,
                "forecast_min": forecast_min,
                "range_width": range_width,
                "intercept": intercept,
                "slope": slope,
                "source_file": meta["source_file"],
            }
        )

    return rows_out, None


def column_width_for(name: str, rows: list[dict[str, Any]], minimum: int = 12, maximum: int = 48) -> int:
    max_len = len(name)
    for row in rows:
        value = row.get(name)
        if value is None:
            continue
        max_len = max(max_len, len(str(value)))
    return max(minimum, min(maximum, max_len + 2))


def write_output_sheet(sheet: xw.Sheet, columns: list[str], rows: list[dict[str, Any]]) -> None:
    table = [columns]
    for row in rows:
        table.append([row.get(column) for column in columns])

    sheet.range((1, 1)).value = table

    header_range = sheet.range((1, 1), (1, len(columns)))
    header_range.api.Font.Bold = True

    last_row = len(table)
    last_col = len(columns)
    data_range = sheet.range((1, 1), (last_row, last_col))
    data_range.api.AutoFilter()

    sheet.activate()
    sheet.range("A2").select()
    sheet.book.app.api.ActiveWindow.FreezePanes = True

    for idx, column in enumerate(columns, start=1):
        sheet.range((1, idx)).column_width = column_width_for(column, rows)


def write_output_workbook(
    app: xw.App,
    output_path: Path,
    empirical_rows: list[dict[str, Any]],
    regression_rows: list[dict[str, Any]],
) -> None:
    out_wb = app.books.add()
    try:
        empirical_sheet = out_wb.sheets[0]
        empirical_sheet.name = "empirical_candidates"
        regression_sheet = out_wb.sheets.add("regression_candidates", after=empirical_sheet)

        write_output_sheet(empirical_sheet, EMPIRICAL_COLUMNS, empirical_rows)
        write_output_sheet(regression_sheet, REGRESSION_COLUMNS, regression_rows)

        out_wb.save(str(output_path))
    finally:
        safe_close_workbook(out_wb)


def should_skip_file(file_path: Path) -> Optional[str]:
    if not file_path.is_file():
        return "not a file"
    if file_path.name.startswith("~"):
        return "temporary file"
    if file_path.suffix.lower() != ".xlsx":
        return "not .xlsx"
    return None


def main() -> None:
    input_path = Path(input_dir).expanduser().resolve()
    output_path_dir = Path(output_dir).expanduser().resolve()

    if not input_path.exists() or not input_path.is_dir():
        raise FileNotFoundError(f"input_dir does not exist or is not a directory: {input_path}")

    output_path_dir.mkdir(parents=True, exist_ok=True)
    output_path = next_output_path(input_path, output_path_dir)

    processed_files = 0
    empirical_rows: list[dict[str, Any]] = []
    regression_rows: list[dict[str, Any]] = []

    app = xw.App(visible=False, add_book=False)
    app.display_alerts = False
    app.screen_updating = False

    try:
        for file_path in sorted(input_path.iterdir()):
            skip_reason = should_skip_file(file_path)
            if skip_reason:
                if file_path.name:
                    print(f"skipped: {file_path.name} ({skip_reason})")
                continue

            if file_path.resolve() == output_path.resolve():
                print(f"skipped: {file_path.name} (matches output filename)")
                continue

            try:
                meta = parse_file_label(file_path)
            except Exception as exc:
                print(f"skipped: {file_path.name} (filename parse error: {exc})")
                continue

            wb: Optional[xw.Book] = None
            try:
                wb = app.books.open(str(file_path), update_links=False)

                empirical_result, empirical_error = process_empirical_sheet(wb, meta, n_quarters=10)
                regression_result, regression_error = process_regression_sheet(wb, meta, n_quarters=10)

                if empirical_error:
                    print(f"skipped: {file_path.name} ({empirical_error})")
                else:
                    empirical_rows.extend(empirical_result)

                if regression_error:
                    print(f"skipped: {file_path.name} ({regression_error})")
                else:
                    regression_rows.extend(regression_result)

                processed_files += 1
                print(f"processed: {file_path.name}")
            except Exception as exc:
                print(f"skipped: {file_path.name} (processing error: {exc})")
            finally:
                if wb is not None:
                    safe_close_workbook(wb)

        write_output_workbook(app, output_path, empirical_rows, regression_rows)
    finally:
        try:
            app.quit()
        except Exception:
            pass

    print(f"output_path: {output_path}")
    print(f"files_processed: {processed_files}")
    print(f"empirical_rows: {len(empirical_rows)}")
    print(f"regression_rows: {len(regression_rows)}")


if __name__ == "__main__":
    main()
