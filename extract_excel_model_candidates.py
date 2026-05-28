#!/usr/bin/env python3
"""Build empirical/regression candidate rows from model workbooks."""

from __future__ import annotations

from dataclasses import dataclass
from datetime import date
from pathlib import Path
import re
from typing import Any

import xlwings as xw
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

# Configure these two paths for your environment.
input_dir = Path("input")
output_dir = Path("output")

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

PERIOD_PATTERN = re.compile(
    r"(?i)(early|mid|late)(jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)(\d{4})"
)
DAY_BY_PART = {"early": 5, "mid": 15, "late": 25}
MONTH_BY_ABBR = {
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


@dataclass(frozen=True)
class FileLabel:
    model: str
    ticker: str
    model_period: str
    model_date: str


@dataclass
class SheetSnapshot:
    top_row: int
    left_col: int
    values: list[list[Any]]
    labels: dict[str, list[tuple[int, int]]]


def parse_file_label(file_name: str) -> FileLabel:
    stem = Path(file_name).stem
    parts = [part.strip() for part in stem.split(" - ") if part.strip()]

    ticker = parts[1] if len(parts) >= 2 else "UNKNOWN"
    ticker = ticker.replace(" ", "")

    period_source = parts[2] if len(parts) >= 3 else stem
    period_token = period_source.split("_", 1)[0]
    match = PERIOD_PATTERN.search(period_token)

    if not match:
        model_period = "Unknown_0000"
        model_date = ""
    else:
        part_name = match.group(1).lower()
        month_abbr = match.group(2).lower()
        year = int(match.group(3))
        model_period = f"{part_name.capitalize()}{month_abbr.capitalize()}_{year}"
        model_date = date(year, MONTH_BY_ABBR[month_abbr], DAY_BY_PART[part_name]).isoformat()

    model = f"{ticker}_{model_period}"
    return FileLabel(
        model=model,
        ticker=ticker,
        model_period=model_period,
        model_date=model_date,
    )


def build_output_path(in_dir: Path, out_dir: Path) -> Path:
    out_dir.mkdir(parents=True, exist_ok=True)
    base_name = f"{in_dir.name}_PARAM"

    candidate = out_dir / f"{base_name}.xlsx"
    if not candidate.exists():
        return candidate

    index = 1
    while True:
        candidate = out_dir / f"{base_name}.{index}.xlsx"
        if not candidate.exists():
            return candidate
        index += 1


def to_float(value: Any) -> float | None:
    if value is None or value == "":
        return None
    if isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        return float(value)
    if isinstance(value, str):
        stripped = value.strip().replace(",", "")
        if not stripped:
            return None
        try:
            return float(stripped)
        except ValueError:
            return None
    return None


def to_2d(values: Any) -> list[list[Any]]:
    if values is None:
        return []
    if isinstance(values, (list, tuple)):
        if not values:
            return []
        first = values[0]
        if isinstance(first, (list, tuple)):
            return [list(row) for row in values]
        return [list(values)]
    return [[values]]


def to_1d(values: Any) -> list[Any]:
    if values is None:
        return []
    if isinstance(values, (list, tuple)):
        if not values:
            return []
        first = values[0]
        if isinstance(first, (list, tuple)):
            return [row[0] if row else None for row in values]
        return list(values)
    return [values]


def capture_snapshot(sheet: xw.Sheet) -> SheetSnapshot:
    used = sheet.used_range
    values = to_2d(used.value)
    labels: dict[str, list[tuple[int, int]]] = {}

    for row_idx, row_values in enumerate(values):
        for col_idx, value in enumerate(row_values):
            if isinstance(value, str):
                key = value.strip().lower()
                if key:
                    abs_row = used.row + row_idx
                    abs_col = used.column + col_idx
                    labels.setdefault(key, []).append((abs_row, abs_col))

    return SheetSnapshot(
        top_row=used.row,
        left_col=used.column,
        values=values,
        labels=labels,
    )


def find_max_anchor(snapshot: SheetSnapshot) -> tuple[int, int] | None:
    candidates = list(snapshot.labels.get("max", []))
    if not candidates:
        for label, coords in snapshot.labels.items():
            if label.startswith("max"):
                candidates.extend(coords)
    if not candidates:
        return None

    min_candidates = snapshot.labels.get("min", [])
    for candidate in sorted(candidates, key=lambda rc: (rc[1], rc[0]), reverse=True):
        cand_row, cand_col = candidate
        for min_row, min_col in min_candidates:
            if abs(min_col - cand_col) <= 2 and 0 <= (min_row - cand_row) <= 3:
                return candidate
    return sorted(candidates, key=lambda rc: (rc[1], rc[0]), reverse=True)[0]


def find_nearest_label(
    snapshot: SheetSnapshot,
    keywords: list[str],
    anchor: tuple[int, int],
) -> tuple[int, int] | None:
    keyword_lc = [word.lower() for word in keywords]
    candidates: list[tuple[int, int]] = []
    for label, coords in snapshot.labels.items():
        if any(word in label for word in keyword_lc):
            candidates.extend(coords)
    if not candidates:
        return None

    anchor_row, anchor_col = anchor
    return min(candidates, key=lambda rc: abs(rc[0] - anchor_row) + abs(rc[1] - anchor_col))


def right_value(sheet: xw.Sheet, coord: tuple[int, int] | None) -> Any:
    if coord is None:
        return None
    row, col = coord
    value = sheet.cells(row, col + 1).value
    if value not in (None, ""):
        return value
    return None


def read_row_values(sheet: xw.Sheet, row: int, start_col: int, end_col: int) -> list[Any]:
    if row < 1 or start_col < 1 or end_col < start_col:
        return []
    values = sheet.range((row, start_col), (row, end_col)).value
    return to_1d(values)


def extract_numeric_points(values: list[Any], start_col: int) -> list[tuple[int, float]]:
    points: list[tuple[int, float]] = []
    for idx, value in enumerate(values):
        num = to_float(value)
        if num is not None:
            points.append((start_col + idx, num))
    return points


def map_text_by_col(values: list[Any], start_col: int) -> dict[int, str]:
    labels: dict[int, str] = {}
    for idx, value in enumerate(values):
        if isinstance(value, str):
            cleaned = value.strip()
            if cleaned:
                labels[start_col + idx] = cleaned
    return labels


def set_formula2_r1c1(cell: xw.Range, formula_r1c1: str) -> None:
    try:
        cell.formula2 = formula_r1c1
    except Exception:
        try:
            cell.api.Formula2R1C1 = formula_r1c1
        except Exception:
            cell.formula = formula_r1c1


def safe_subtract(left: float | None, right: float | None) -> float | None:
    if left is None or right is None:
        return None
    return left - right


def nearly_equal(left: float | None, right: float | None, tol: float = 1e-9) -> bool:
    if left is None and right is None:
        return True
    if left is None or right is None:
        return False
    return abs(left - right) <= tol


def close_source_workbook(workbook: xw.Book) -> None:
    try:
        workbook.close(save=False)
        return
    except TypeError:
        pass
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
        pass


def process_empirical_sheet(
    workbook: xw.Book,
    sheet: xw.Sheet,
    label: FileLabel,
    source_file: str,
) -> list[dict[str, Any]]:
    snapshot = capture_snapshot(sheet)
    anchor = find_max_anchor(snapshot)
    if anchor is None:
        print(f"  skipped empirical in {source_file}: max anchor not found")
        return []

    anchor_row, anchor_col = anchor
    value_col = anchor_col + 1
    scan_start_col = max(1, anchor_col - 80)
    scan_end_col = max(scan_start_col, anchor_col - 1)

    min_anchor = find_nearest_label(snapshot, ["min"], anchor)
    min_row = min_anchor[0] if min_anchor else anchor_row + 1

    penetration_anchor = find_nearest_label(snapshot, ["penetration"], anchor)
    penetration_row = penetration_anchor[0] if penetration_anchor else max(1, anchor_row - 9)
    quarter_row = max(1, penetration_row - 1)

    quarterly_anchor = find_nearest_label(snapshot, ["quarterly sales"], anchor)
    reported_anchor = find_nearest_label(snapshot, ["reported sales"], anchor)
    growth_anchor = find_nearest_label(snapshot, ["growth rate"], anchor)
    captured_anchor = find_nearest_label(snapshot, ["sales captured in db", "captured in db"], anchor)
    estimated_anchor = find_nearest_label(
        snapshot,
        ["estimated total sold", "estimated sold", "tot fcst"],
        anchor,
    )

    quarterly_row = quarterly_anchor[0] if quarterly_anchor else penetration_row + 1
    reported_row = reported_anchor[0] if reported_anchor else penetration_row + 2

    penetration_values = read_row_values(sheet, penetration_row, scan_start_col, scan_end_col)
    quarterly_values = read_row_values(sheet, quarterly_row, scan_start_col, scan_end_col)
    reported_values = read_row_values(sheet, reported_row, scan_start_col, scan_end_col)
    quarter_labels = read_row_values(sheet, quarter_row, scan_start_col, scan_end_col)

    penetration_points = extract_numeric_points(penetration_values, scan_start_col)
    quarterly_points = dict(extract_numeric_points(quarterly_values, scan_start_col))
    reported_points = dict(extract_numeric_points(reported_values, scan_start_col))
    quarter_by_col = map_text_by_col(quarter_labels, scan_start_col)

    if not penetration_points:
        print(f"  skipped empirical in {source_file}: no penetration series found")
        return []

    static_growth_value = to_float(right_value(sheet, growth_anchor))
    static_captured_value = to_float(right_value(sheet, captured_anchor))
    estimated_total = to_float(right_value(sheet, estimated_anchor))
    static_max = to_float(sheet.cells(anchor_row, value_col).value)
    static_min = to_float(sheet.cells(min_row, value_col).value)

    helper_col = anchor_col + 6
    avg_cell = sheet.cells(anchor_row + 2, helper_col)
    max_cell = sheet.cells(anchor_row + 3, helper_col)
    min_cell = sheet.cells(anchor_row + 4, helper_col)

    rows: list[dict[str, Any]] = []
    max_quarters = min(10, len(penetration_points))
    for n_quarters in range(1, max_quarters + 1):
        subset = penetration_points[-n_quarters:]
        start_col = subset[0][0]
        end_col = subset[-1][0]

        set_formula2_r1c1(
            avg_cell,
            f"=AVERAGE(R{penetration_row}C{start_col}:R{penetration_row}C{end_col})",
        )
        set_formula2_r1c1(
            max_cell,
            f"=MAX(R{penetration_row}C{start_col}:R{penetration_row}C{end_col})",
        )
        set_formula2_r1c1(
            min_cell,
            f"=MIN(R{penetration_row}C{start_col}:R{penetration_row}C{end_col})",
        )
        workbook.app.calculate()

        avg_penetration = to_float(avg_cell.value)
        pen_max = to_float(max_cell.value)
        pen_min = to_float(min_cell.value)
        if avg_penetration is None:
            continue

        last_col = end_col
        quarterly_sales = quarterly_points.get(last_col)
        if quarterly_sales is None:
            quarterly_sales = to_float(right_value(sheet, quarterly_anchor))

        reported_sales = reported_points.get(last_col)
        if reported_sales is None:
            reported_sales = to_float(right_value(sheet, reported_anchor))

        growth_value = static_growth_value
        captured_value = static_captured_value
        if growth_value is None and quarterly_sales not in (None, 0) and reported_sales is not None:
            growth_value = (quarterly_sales - reported_sales) / reported_sales if reported_sales else None
        if captured_value is None and quarterly_sales not in (None, 0) and reported_sales is not None:
            captured_value = reported_sales / quarterly_sales

        if quarterly_sales is not None:
            forecast_value = avg_penetration * quarterly_sales
            dynamic_max = pen_max * quarterly_sales if pen_max is not None else None
            dynamic_min = pen_min * quarterly_sales if pen_min is not None else None
        else:
            forecast_value = estimated_total
            dynamic_max = None
            dynamic_min = None

        forecast_max = static_max if static_max is not None else dynamic_max
        forecast_min = static_min if static_min is not None else dynamic_min
        range_width = safe_subtract(forecast_max, forecast_min)

        rows.append(
            {
                "model": label.model,
                "ticker": label.ticker,
                "model_period": label.model_period,
                "model_date": label.model_date,
                "method": "empirical",
                "parameter_name": "avg_penetration_pct",
                "parameter_value": avg_penetration,
                "num_quarters_used": n_quarters,
                "last_quarter_used": quarter_by_col.get(last_col, f"col_{last_col}"),
                "forecast_value": forecast_value,
                "actual_value": reported_sales,
                "forecast_max": forecast_max,
                "forecast_min": forecast_min,
                "range_width": range_width,
                "avg_penetration_pct": avg_penetration,
                "quarterly_sales": quarterly_sales,
                "reported_sales": reported_sales,
                "growth_rate_pct": growth_value,
                "sales_captured_in_db_pct": captured_value,
                "source_file": source_file,
            }
        )

    return rows


def process_regression_sheet(
    workbook: xw.Book,
    sheet: xw.Sheet,
    label: FileLabel,
    source_file: str,
) -> list[dict[str, Any]]:
    snapshot = capture_snapshot(sheet)
    anchor = find_max_anchor(snapshot)
    if anchor is None:
        print(f"  skipped regression in {source_file}: max anchor not found")
        return []

    anchor_row, anchor_col = anchor
    value_col = anchor_col + 1

    y_col = anchor_col - 7
    x_col = anchor_col - 11
    if y_col < 1 or x_col < 1:
        print(f"  skipped regression in {source_file}: invalid anchor offsets")
        return []

    min_anchor = find_nearest_label(snapshot, ["min"], anchor)
    min_row = min_anchor[0] if min_anchor else anchor_row + 1

    max_row_for_series = anchor_row - 1
    if max_row_for_series < 1:
        print(f"  skipped regression in {source_file}: no x/y source rows")
        return []

    x_values = to_1d(sheet.range((1, x_col), (max_row_for_series, x_col)).value)
    y_values = to_1d(sheet.range((1, y_col), (max_row_for_series, y_col)).value)

    points: list[tuple[int, float, float]] = []
    for idx, (x_raw, y_raw) in enumerate(zip(x_values, y_values), start=1):
        x_num = to_float(x_raw)
        y_num = to_float(y_raw)
        if x_num is not None and y_num is not None:
            points.append((idx, x_num, y_num))

    if not points:
        print(f"  skipped regression in {source_file}: no numeric x/y pairs")
        return []

    forecast_x = to_float(sheet.cells(anchor_row, x_col).value)
    if forecast_x is None:
        forecast_x = points[-1][1]

    static_max = to_float(sheet.cells(anchor_row, value_col).value)
    static_min = to_float(sheet.cells(min_row, value_col).value)

    actual_anchor = find_nearest_label(snapshot, ["actual", "reported sales"], anchor)
    actual_value = to_float(right_value(sheet, actual_anchor))

    helper_col = anchor_col + 6
    intercept_cell = sheet.cells(anchor_row + 2, helper_col)
    slope_cell = sheet.cells(anchor_row + 3, helper_col)

    rows: list[dict[str, Any]] = []
    max_quarters = min(10, len(points))
    for n_quarters in range(1, max_quarters + 1):
        subset = points[-n_quarters:]
        start_row = subset[0][0]
        end_row = subset[-1][0]

        set_formula2_r1c1(
            intercept_cell,
            f"=INTERCEPT(R{start_row}C{y_col}:R{end_row}C{y_col},R{start_row}C{x_col}:R{end_row}C{x_col})",
        )
        set_formula2_r1c1(
            slope_cell,
            f"=SLOPE(R{start_row}C{y_col}:R{end_row}C{y_col},R{start_row}C{x_col}:R{end_row}C{x_col})",
        )
        workbook.app.calculate()

        intercept = to_float(intercept_cell.value)
        slope = to_float(slope_cell.value)
        if intercept is None or slope is None:
            continue

        forecast_total = intercept + (slope * forecast_x) if forecast_x is not None else None
        subset_y = [point[2] for point in subset]
        dynamic_max = max(subset_y) if subset_y else None
        dynamic_min = min(subset_y) if subset_y else None
        forecast_max = static_max if static_max is not None else dynamic_max
        forecast_min = static_min if static_min is not None else dynamic_min
        range_width = safe_subtract(forecast_max, forecast_min)

        new_row = {
            "model": label.model,
            "ticker": label.ticker,
            "model_period": label.model_period,
            "model_date": label.model_date,
            "method": "regression",
            "parameter_name": "num_quarters_used",
            "parameter_value": n_quarters,
            "num_quarters_used": n_quarters,
            "forecast_value": forecast_total,
            "actual_value": actual_value,
            "forecast_max": forecast_max,
            "forecast_min": forecast_min,
            "range_width": range_width,
            "intercept": intercept,
            "slope": slope,
            "source_file": source_file,
        }

        if rows:
            prev_row = rows[-1]
            duplicate = (
                nearly_equal(prev_row["forecast_value"], new_row["forecast_value"])
                and nearly_equal(prev_row["forecast_max"], new_row["forecast_max"])
                and nearly_equal(prev_row["forecast_min"], new_row["forecast_min"])
                and nearly_equal(prev_row["intercept"], new_row["intercept"])
                and nearly_equal(prev_row["slope"], new_row["slope"])
            )
            if duplicate:
                continue

        rows.append(new_row)

    return rows


def write_candidate_sheet(
    workbook: Workbook,
    sheet_name: str,
    columns: list[str],
    rows: list[dict[str, Any]],
) -> None:
    sheet = workbook.create_sheet(sheet_name)
    sheet.append(columns)
    for row in rows:
        sheet.append([row.get(column) for column in columns])

    for header_cell in sheet[1]:
        header_cell.font = Font(bold=True)

    sheet.freeze_panes = "A2"
    sheet.auto_filter.ref = sheet.dimensions

    widths = [len(column) for column in columns]
    for row in rows:
        for idx, column in enumerate(columns):
            value = row.get(column)
            if value is None:
                continue
            text = str(value)
            if len(text) > widths[idx]:
                widths[idx] = len(text)

    for idx, width in enumerate(widths, start=1):
        sheet.column_dimensions[get_column_letter(idx)].width = max(12, min(width + 2, 60))


def write_output_workbook(
    output_path: Path,
    empirical_rows: list[dict[str, Any]],
    regression_rows: list[dict[str, Any]],
) -> None:
    workbook = Workbook()
    default_sheet = workbook.active
    workbook.remove(default_sheet)

    write_candidate_sheet(workbook, "empirical_candidates", EMPIRICAL_COLUMNS, empirical_rows)
    write_candidate_sheet(workbook, "regression_candidates", REGRESSION_COLUMNS, regression_rows)
    workbook.save(output_path)


def main() -> None:
    if not input_dir.exists():
        raise FileNotFoundError(f"input_dir does not exist: {input_dir}")

    output_path = build_output_path(input_dir, output_dir)
    empirical_rows: list[dict[str, Any]] = []
    regression_rows: list[dict[str, Any]] = []
    processed_files = 0

    app: xw.App | None = None
    try:
        app = xw.App(visible=False, add_book=False)
        app.display_alerts = False
        app.screen_updating = False
        try:
            app.calculation = "manual"
        except Exception:
            pass

        generated_output_pattern = re.compile(
            rf"^{re.escape(input_dir.name)}_PARAM(?:\.\d+)?\.xlsx$",
            re.IGNORECASE,
        )

        for file_path in sorted(input_dir.iterdir()):
            if not file_path.is_file():
                print(f"skipped: {file_path.name} (not a file)")
                continue
            if file_path.name.startswith("~"):
                print(f"skipped: {file_path.name} (temporary Excel file)")
                continue
            if file_path.suffix.lower() != ".xlsx":
                print(f"skipped: {file_path.name} (not an .xlsx file)")
                continue
            if generated_output_pattern.match(file_path.name):
                print(f"skipped: {file_path.name} (generated output workbook)")
                continue

            workbook: xw.Book | None = None
            try:
                workbook = app.books.open(str(file_path), update_links=False)
                label = parse_file_label(file_path.name)

                try:
                    empirical_sheet = workbook.sheets["Empirical Model"]
                    empirical_rows.extend(
                        process_empirical_sheet(workbook, empirical_sheet, label, file_path.name)
                    )
                except Exception as exc:
                    print(f"  skipped empirical in {file_path.name}: {exc}")

                try:
                    regression_sheet = workbook.sheets["Regression Model"]
                    regression_rows.extend(
                        process_regression_sheet(workbook, regression_sheet, label, file_path.name)
                    )
                except Exception as exc:
                    print(f"  skipped regression in {file_path.name}: {exc}")

                processed_files += 1
                print(f"processed: {file_path.name}")
            except Exception as exc:
                print(f"skipped: {file_path.name} (error: {exc})")
            finally:
                if workbook is not None:
                    close_source_workbook(workbook)
    finally:
        if app is not None:
            try:
                app.quit()
            except Exception:
                pass

    write_output_workbook(output_path, empirical_rows, regression_rows)

    print(f"output path: {output_path}")
    print(f"files processed: {processed_files}")
    print(f"empirical rows: {len(empirical_rows)}")
    print(f"regression rows: {len(regression_rows)}")


if __name__ == "__main__":
    main()
