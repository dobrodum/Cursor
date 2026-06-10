#!/usr/bin/env python3
from __future__ import annotations

import math
import re
from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path
from typing import Any, Iterable, Sequence

from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

# ===== User-editable paths =====
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

N_QUARTERS = 10
EARLY_MID_LATE_TO_DAY = {"early": 5, "mid": 15, "late": 25}


@dataclass(frozen=True)
class FileLabel:
    model: str
    ticker: str
    model_period: str
    model_date: str


@dataclass
class SheetSnapshot:
    sheet: Any
    top_row: int
    left_col: int
    values: list[list[Any]]

    @property
    def rows(self) -> int:
        return len(self.values)

    @property
    def cols(self) -> int:
        if not self.values:
            return 0
        return max(len(row) for row in self.values)

    @property
    def bottom_row(self) -> int:
        return self.top_row + self.rows - 1

    @property
    def right_col(self) -> int:
        return self.left_col + self.cols - 1

    def get_cached(self, row: int, col: int) -> Any:
        if (
            row < self.top_row
            or col < self.left_col
            or row > self.bottom_row
            or col > self.right_col
        ):
            return None
        r_idx = row - self.top_row
        c_idx = col - self.left_col
        row_values = self.values[r_idx]
        if c_idx >= len(row_values):
            return None
        return row_values[c_idx]

    def iter_cells(self) -> Iterable[tuple[int, int, Any]]:
        for r_offset, row_values in enumerate(self.values):
            abs_row = self.top_row + r_offset
            for c_offset, value in enumerate(row_values):
                abs_col = self.left_col + c_offset
                yield abs_row, abs_col, value


def ensure_2d(values: Any) -> list[list[Any]]:
    if values is None:
        return []
    if not isinstance(values, list):
        return [[values]]
    if not values:
        return []
    if isinstance(values[0], list):
        return values
    return [values]


def to_float(value: Any) -> float | None:
    if value is None:
        return None
    if isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        number = float(value)
        if math.isnan(number) or math.isinf(number):
            return None
        return number
    if isinstance(value, str):
        text = value.strip().replace(",", "")
        if not text:
            return None
        if text.endswith("%"):
            text = text[:-1].strip()
            try:
                return float(text) / 100.0
            except ValueError:
                return None
        try:
            return float(text)
        except ValueError:
            return None
    return None


def parse_file_label(file_name: str) -> FileLabel:
    stem = Path(file_name).stem
    parts = [part.strip() for part in stem.split(" - ")]
    if len(parts) < 3:
        raise ValueError(
            "expected filename format like '... - TICKER - MidJan2026_Send.xlsx'"
        )

    ticker = parts[1]
    period_token = parts[2].split("_")[0]
    token_match = re.match(r"^(Early|Mid|Late)([A-Za-z]+)(\d{4})$", period_token, re.I)
    if token_match is None:
        raise ValueError(f"invalid period token '{period_token}'")

    timing_raw, month_raw, year_raw = token_match.groups()
    timing = timing_raw.capitalize()
    month_abbrev = month_raw[:3].capitalize()
    year = int(year_raw)
    month_num = datetime.strptime(month_abbrev, "%b").month
    day = EARLY_MID_LATE_TO_DAY[timing.lower()]
    model_period = f"{timing}{month_abbrev}_{year}"
    model_date = date(year, month_num, day).isoformat()
    model = f"{ticker}_{model_period}"

    return FileLabel(
        model=model,
        ticker=ticker,
        model_period=model_period,
        model_date=model_date,
    )


def next_output_path(in_dir: Path, out_dir: Path) -> Path:
    folder_name = in_dir.resolve().name
    base_name = f"{folder_name}_PARAM"
    first_path = out_dir / f"{base_name}.xlsx"
    if not first_path.exists():
        return first_path

    suffix = 1
    while True:
        candidate = out_dir / f"{base_name}.{suffix}.xlsx"
        if not candidate.exists():
            return candidate
        suffix += 1


def build_snapshot(sheet: Any) -> SheetSnapshot:
    used = sheet.used_range
    top_row = used.row
    left_col = used.column
    values = ensure_2d(used.value)
    return SheetSnapshot(sheet=sheet, top_row=top_row, left_col=left_col, values=values)


def normalize_text(value: Any) -> str:
    if not isinstance(value, str):
        return ""
    return re.sub(r"\s+", " ", value.strip().lower())


def find_anchor_max(snapshot: SheetSnapshot) -> tuple[int, int]:
    candidates: list[tuple[int, int]] = []
    for row, col, value in snapshot.iter_cells():
        if normalize_text(value) == "max":
            candidates.append((row, col))

    if not candidates:
        raise ValueError("could not find 'max' anchor")
    if len(candidates) == 1:
        return candidates[0]

    # Prefer "max" that has "min" directly below it.
    for row, col in candidates:
        below_text = normalize_text(snapshot.get_cached(row + 1, col))
        if below_text == "min":
            return row, col
    return candidates[-1]


def find_label_cell(
    snapshot: SheetSnapshot,
    keywords: Sequence[str],
    anchor: tuple[int, int] | None = None,
    max_row_distance: int = 120,
) -> tuple[int, int] | None:
    norm_keywords = [k.lower() for k in keywords]
    matches: list[tuple[int, int, int]] = []
    for row, col, value in snapshot.iter_cells():
        text = normalize_text(value)
        if not text:
            continue
        if any(keyword in text for keyword in norm_keywords):
            if anchor is None:
                matches.append((0, row, col))
            else:
                distance = abs(row - anchor[0]) + abs(col - anchor[1])
                if abs(row - anchor[0]) <= max_row_distance:
                    matches.append((distance, row, col))
    if not matches:
        return None
    matches.sort(key=lambda item: item[0])
    _, row, col = matches[0]
    return row, col


def find_header_col(
    snapshot: SheetSnapshot,
    anchor: tuple[int, int],
    keywords: Sequence[str],
    fallback_offset: int,
) -> int | None:
    header_cell = find_label_cell(
        snapshot=snapshot,
        keywords=keywords,
        anchor=anchor,
        max_row_distance=80,
    )
    if header_cell is not None:
        return header_cell[1]
    fallback_col = anchor[1] + fallback_offset
    if fallback_col < 1:
        return None
    return fallback_col


def read_live_value(sheet: Any, row: int | None, col: int | None) -> Any:
    if row is None or col is None:
        return None
    if row < 1 or col < 1:
        return None
    return sheet.range((row, col)).value


def find_value_cell_right(
    snapshot: SheetSnapshot,
    label_cell: tuple[int, int] | None,
    max_right_steps: int = 4,
) -> tuple[int, int] | None:
    if label_cell is None:
        return None
    row, col = label_cell
    for step in range(1, max_right_steps + 1):
        value = snapshot.get_cached(row, col + step)
        if value not in (None, ""):
            return row, col + step
    return row, col + 1


def find_max_min_value_cells(
    snapshot: SheetSnapshot,
    max_anchor: tuple[int, int],
) -> tuple[tuple[int, int] | None, tuple[int, int] | None]:
    max_value_cell = find_value_cell_right(snapshot, max_anchor)
    min_label = find_label_cell(snapshot, ["min"], anchor=max_anchor, max_row_distance=30)
    min_value_cell = find_value_cell_right(snapshot, min_label)
    return max_value_cell, min_value_cell


def collect_data_rows(
    sheet: Any,
    row_start: int,
    row_end: int,
    value_cols: Sequence[int | None],
    max_empty_streak: int = 8,
) -> list[int]:
    usable_cols = [col for col in value_cols if col is not None and col > 0]
    if not usable_cols:
        return []

    rows: list[int] = []
    empty_streak = 0
    for row in range(row_end, row_start - 1, -1):
        has_data = False
        for col in usable_cols:
            if to_float(sheet.range((row, col)).value) is not None:
                has_data = True
                break
        if has_data:
            rows.append(row)
            empty_streak = 0
        else:
            empty_streak += 1
            if rows and empty_streak >= max_empty_streak:
                break
    rows.reverse()
    return rows


def set_formula_r1c1(range_obj: Any, formula: str) -> None:
    try:
        range_obj.formula2 = formula
    except Exception:
        range_obj.formula = formula


def close_source_wb(wb: Any) -> None:
    close_attempts = (
        lambda: wb.close(save=False),
        lambda: wb.close(False),
        lambda: wb.api.Close(SaveChanges=False),
        lambda: wb.api.Close(False),
    )
    for close_fn in close_attempts:
        try:
            close_fn()
            return
        except TypeError:
            continue
        except Exception:
            continue


def process_empirical_model(
    wb: Any,
    label: FileLabel,
    source_file: str,
) -> list[dict[str, Any]]:
    if "Empirical Model" not in [sheet.name for sheet in wb.sheets]:
        return []

    sheet = wb.sheets["Empirical Model"]
    snapshot = build_snapshot(sheet)
    max_anchor = find_anchor_max(snapshot)

    quarter_col = find_header_col(snapshot, max_anchor, ["quarter"], fallback_offset=-12)
    quarterly_sales_col = find_header_col(
        snapshot,
        max_anchor,
        ["quarterly sales", "quarter sales"],
        fallback_offset=-11,
    )
    reported_sales_col = find_header_col(
        snapshot,
        max_anchor,
        ["reported sales", "actual sales"],
        fallback_offset=-10,
    )
    growth_rate_col = find_header_col(
        snapshot,
        max_anchor,
        ["growth rate", "growth"],
        fallback_offset=-9,
    )
    sales_captured_col = find_header_col(
        snapshot,
        max_anchor,
        ["sales captured in db", "captured in db"],
        fallback_offset=-8,
    )
    penetration_col = find_header_col(
        snapshot,
        max_anchor,
        ["penetration"],
        fallback_offset=-7,
    )

    data_rows = collect_data_rows(
        sheet=sheet,
        row_start=snapshot.top_row,
        row_end=max_anchor[0] - 1,
        value_cols=(penetration_col, quarterly_sales_col, reported_sales_col),
    )
    if not data_rows:
        return []

    avg_pen_label = find_label_cell(snapshot, ["avg penetration"], anchor=max_anchor)
    avg_pen_cell = find_value_cell_right(snapshot, avg_pen_label)
    est_total_label = find_label_cell(
        snapshot,
        ["estimated total sold", "total sold", "tot fcst"],
        anchor=max_anchor,
    )
    est_total_cell = find_value_cell_right(snapshot, est_total_label)
    reported_total_label = find_label_cell(
        snapshot,
        ["reported sales", "actual sales"],
        anchor=max_anchor,
    )
    reported_total_cell = find_value_cell_right(snapshot, reported_total_label)
    forecast_max_cell, forecast_min_cell = find_max_min_value_cells(snapshot, max_anchor)

    original_avg_formula = None
    original_avg_value = None
    live_avg_pen_cell = None
    if avg_pen_cell is not None:
        live_avg_pen_cell = sheet.range(avg_pen_cell)
        try:
            original_avg_formula = live_avg_pen_cell.formula2
        except Exception:
            original_avg_formula = None
        original_avg_value = live_avg_pen_cell.value

    empirical_rows: list[dict[str, Any]] = []
    loop_count = min(N_QUARTERS, len(data_rows))

    try:
        for n_used in range(1, loop_count + 1):
            row_start = data_rows[-n_used]
            row_end = data_rows[-1]

            avg_pen = None
            if penetration_col is not None and live_avg_pen_cell is not None:
                avg_formula = (
                    f"=AVERAGE(R{row_start}C{penetration_col}:R{row_end}C{penetration_col})"
                )
                set_formula_r1c1(live_avg_pen_cell, avg_formula)
                wb.app.calculate()
                avg_pen = to_float(live_avg_pen_cell.value)
            elif penetration_col is not None:
                pen_vals = [
                    to_float(sheet.range((row, penetration_col)).value)
                    for row in data_rows[-n_used:]
                ]
                clean_pen_vals = [value for value in pen_vals if value is not None]
                if clean_pen_vals:
                    avg_pen = sum(clean_pen_vals) / len(clean_pen_vals)

            last_data_row = data_rows[-1]
            last_quarter_used = read_live_value(sheet, last_data_row, quarter_col)
            quarterly_sales = to_float(read_live_value(sheet, last_data_row, quarterly_sales_col))
            reported_sales = to_float(read_live_value(sheet, last_data_row, reported_sales_col))
            growth_rate = to_float(read_live_value(sheet, last_data_row, growth_rate_col))
            sales_captured = to_float(read_live_value(sheet, last_data_row, sales_captured_col))

            forecast_value = to_float(
                read_live_value(sheet, *(est_total_cell or (None, None)))
            )
            actual_value = to_float(
                read_live_value(sheet, *(reported_total_cell or (None, None)))
            )
            if actual_value is None:
                actual_value = reported_sales

            if forecast_value is None and avg_pen not in (None, 0.0) and quarterly_sales is not None:
                forecast_value = quarterly_sales / avg_pen

            forecast_max = to_float(
                read_live_value(sheet, *(forecast_max_cell or (None, None)))
            )
            forecast_min = to_float(
                read_live_value(sheet, *(forecast_min_cell or (None, None)))
            )
            if forecast_max is None and forecast_value is not None:
                forecast_max = forecast_value
            if forecast_min is None and forecast_value is not None:
                forecast_min = forecast_value

            range_width = None
            if forecast_max is not None and forecast_min is not None:
                range_width = forecast_max - forecast_min

            empirical_rows.append(
                {
                    "model": label.model,
                    "ticker": label.ticker,
                    "model_period": label.model_period,
                    "model_date": label.model_date,
                    "method": "empirical",
                    "parameter_name": "avg_penetration_pct",
                    "parameter_value": avg_pen,
                    "num_quarters_used": n_used,
                    "last_quarter_used": last_quarter_used,
                    "forecast_value": forecast_value,
                    "actual_value": actual_value,
                    "forecast_max": forecast_max,
                    "forecast_min": forecast_min,
                    "range_width": range_width,
                    "avg_penetration_pct": avg_pen,
                    "quarterly_sales": quarterly_sales,
                    "reported_sales": reported_sales,
                    "growth_rate_pct": growth_rate,
                    "sales_captured_in_db_pct": sales_captured,
                    "source_file": source_file,
                }
            )
    finally:
        if live_avg_pen_cell is not None:
            if original_avg_formula not in (None, ""):
                set_formula_r1c1(live_avg_pen_cell, original_avg_formula)
            else:
                live_avg_pen_cell.value = original_avg_value

    return empirical_rows


def process_regression_model(
    wb: Any,
    label: FileLabel,
    source_file: str,
) -> list[dict[str, Any]]:
    if "Regression Model" not in [sheet.name for sheet in wb.sheets]:
        return []

    sheet = wb.sheets["Regression Model"]
    snapshot = build_snapshot(sheet)
    max_anchor = find_anchor_max(snapshot)

    y_col = max_anchor[1] - 7
    x_col = max_anchor[1] - 11

    data_rows = collect_data_rows(
        sheet=sheet,
        row_start=snapshot.top_row,
        row_end=max_anchor[0] - 1,
        value_cols=(x_col, y_col),
    )
    data_rows = [
        row
        for row in data_rows
        if to_float(sheet.range((row, x_col)).value) is not None
        and to_float(sheet.range((row, y_col)).value) is not None
    ]
    if not data_rows:
        return []

    max_label = max_anchor
    min_label = find_label_cell(snapshot, ["min"], anchor=max_anchor, max_row_distance=30)
    max_value_cell = find_value_cell_right(snapshot, max_label)
    min_value_cell = find_value_cell_right(snapshot, min_label)
    forecast_label = find_label_cell(
        snapshot,
        ["tot fcst w/o sa", "total forecast without sa", "tot fcst"],
        anchor=max_anchor,
    )
    forecast_value_cell = find_value_cell_right(snapshot, forecast_label)
    actual_label = find_label_cell(snapshot, ["actual sales", "reported sales"], anchor=max_anchor)
    actual_value_cell = find_value_cell_right(snapshot, actual_label)

    temp_base_row = max_anchor[0] + 5
    temp_intercept_col = max_anchor[1] + 3
    temp_slope_col = max_anchor[1] + 4

    intercept_cells: list[Any] = []
    slope_cells: list[Any] = []
    loop_count = min(N_QUARTERS, len(data_rows))
    for n_used in range(1, loop_count + 1):
        row_start = data_rows[-n_used]
        row_end = data_rows[-1]
        intercept_formula = (
            f"=INTERCEPT(R{row_start}C{y_col}:R{row_end}C{y_col},"
            f"R{row_start}C{x_col}:R{row_end}C{x_col})"
        )
        slope_formula = (
            f"=SLOPE(R{row_start}C{y_col}:R{row_end}C{y_col},"
            f"R{row_start}C{x_col}:R{row_end}C{x_col})"
        )
        intercept_cell = sheet.range((temp_base_row + n_used, temp_intercept_col))
        slope_cell = sheet.range((temp_base_row + n_used, temp_slope_col))
        set_formula_r1c1(intercept_cell, intercept_formula)
        set_formula_r1c1(slope_cell, slope_formula)
        intercept_cells.append(intercept_cell)
        slope_cells.append(slope_cell)

    wb.app.calculate()

    x_vals = [to_float(sheet.range((row, x_col)).value) for row in data_rows]
    x_vals = [value for value in x_vals if value is not None]
    next_x = (max(x_vals) + 1.0) if x_vals else None

    regression_rows: list[dict[str, Any]] = []
    previous_signature: tuple[Any, ...] | None = None
    for n_used in range(1, loop_count + 1):
        intercept = to_float(intercept_cells[n_used - 1].value)
        slope = to_float(slope_cells[n_used - 1].value)

        forecast_value = to_float(
            read_live_value(sheet, *(forecast_value_cell or (None, None)))
        )
        if forecast_value is None and intercept is not None and slope is not None and next_x is not None:
            forecast_value = intercept + slope * next_x

        forecast_max = to_float(read_live_value(sheet, *(max_value_cell or (None, None))))
        forecast_min = to_float(read_live_value(sheet, *(min_value_cell or (None, None))))
        if forecast_max is None and forecast_value is not None:
            forecast_max = forecast_value
        if forecast_min is None and forecast_value is not None:
            forecast_min = forecast_value

        range_width = None
        if forecast_max is not None and forecast_min is not None:
            range_width = forecast_max - forecast_min

        actual_value = to_float(read_live_value(sheet, *(actual_value_cell or (None, None))))
        signature = (
            round(intercept, 10) if intercept is not None else None,
            round(slope, 10) if slope is not None else None,
            round(forecast_value, 10) if forecast_value is not None else None,
            round(forecast_max, 10) if forecast_max is not None else None,
            round(forecast_min, 10) if forecast_min is not None else None,
        )
        if previous_signature is not None and signature == previous_signature:
            continue
        previous_signature = signature

        regression_rows.append(
            {
                "model": label.model,
                "ticker": label.ticker,
                "model_period": label.model_period,
                "model_date": label.model_date,
                "method": "regression",
                "parameter_name": "num_quarters_used",
                "parameter_value": n_used,
                "num_quarters_used": n_used,
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

    return regression_rows


def write_output_workbook(
    destination: Path,
    empirical_rows: list[dict[str, Any]],
    regression_rows: list[dict[str, Any]],
) -> None:
    wb = Workbook()
    default_sheet = wb.active
    wb.remove(default_sheet)

    def write_sheet(name: str, headers: list[str], rows: list[dict[str, Any]]) -> None:
        ws = wb.create_sheet(title=name)
        ws.append(headers)
        for row_data in rows:
            ws.append([row_data.get(col_name) for col_name in headers])

        for header_cell in ws[1]:
            header_cell.font = Font(bold=True)
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions

        for col_index, header in enumerate(headers, start=1):
            max_len = len(header)
            for row_index in range(2, ws.max_row + 1):
                value = ws.cell(row=row_index, column=col_index).value
                if value is None:
                    continue
                max_len = max(max_len, len(str(value)))
            ws.column_dimensions[get_column_letter(col_index)].width = min(max_len + 2, 48)

    write_sheet("empirical_candidates", EMPIRICAL_HEADERS, empirical_rows)
    write_sheet("regression_candidates", REGRESSION_HEADERS, regression_rows)
    wb.save(destination)


def run() -> None:
    in_dir = Path(input_dir)
    out_dir = Path(output_dir)

    if not in_dir.exists():
        raise FileNotFoundError(f"input_dir does not exist: {in_dir}")
    out_dir.mkdir(parents=True, exist_ok=True)

    output_path = next_output_path(in_dir, out_dir)
    source_files = sorted(in_dir.iterdir(), key=lambda path: path.name.lower())

    empirical_rows: list[dict[str, Any]] = []
    regression_rows: list[dict[str, Any]] = []
    processed_files = 0

    try:
        import xlwings as xw
    except ImportError as exc:
        raise RuntimeError(
            "xlwings is required for this workflow. Install with: pip install xlwings"
        ) from exc

    app = xw.App(visible=False, add_book=False)
    app.display_alerts = False
    app.screen_updating = False
    try:
        app.calculation = "manual"
    except Exception:
        pass

    try:
        for file_path in source_files:
            if not file_path.is_file():
                print(f"Skipped: {file_path.name} (not a file)")
                continue
            if file_path.name.startswith("~"):
                print(f"Skipped: {file_path.name} (temporary file)")
                continue
            if file_path.suffix.lower() != ".xlsx":
                print(f"Skipped: {file_path.name} (not .xlsx)")
                continue

            try:
                label = parse_file_label(file_path.name)
            except Exception as exc:
                print(f"Skipped: {file_path.name} (filename parse failed: {exc})")
                continue

            print(f"Processing: {file_path.name}")
            wb = app.books.open(str(file_path), update_links=False)
            try:
                empirical_rows.extend(process_empirical_model(wb, label, file_path.name))
                regression_rows.extend(process_regression_model(wb, label, file_path.name))
                processed_files += 1
            except Exception as exc:
                print(f"Skipped: {file_path.name} (processing failed: {exc})")
            finally:
                close_source_wb(wb)
    finally:
        app.quit()

    write_output_workbook(output_path, empirical_rows, regression_rows)
    print(f"Output: {output_path}")
    print(f"Files processed: {processed_files}")
    print(f"Empirical rows: {len(empirical_rows)}")
    print(f"Regression rows: {len(regression_rows)}")


if __name__ == "__main__":
    run()
