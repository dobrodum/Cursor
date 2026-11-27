#!/usr/bin/env python3
from __future__ import annotations

import calendar
import re
from collections import defaultdict
from dataclasses import dataclass
from datetime import date
from pathlib import Path
from typing import Any, Iterable

import xlwings as xw
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

# ---------------------------------------------------------------------------
# Configure these two paths before running.
# ---------------------------------------------------------------------------
input_dir = Path("/path/to/input")
output_dir = Path("/path/to/output")

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

PHASE_TO_DAY = {"early": 5, "mid": 15, "late": 25}


@dataclass
class ModelMeta:
    model: str
    ticker: str
    model_period: str
    model_date: str


def normalize_label(value: str) -> str:
    lowered = value.strip().lower()
    lowered = lowered.replace("%", " pct ").replace("/", " ")
    lowered = re.sub(r"[_\-]+", " ", lowered)
    lowered = re.sub(r"\s+", " ", lowered)
    return lowered.strip()


def to_float(value: Any) -> float | None:
    if isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        return float(value)
    if isinstance(value, str):
        text = value.strip()
        if not text:
            return None
        pct = text.endswith("%")
        if pct:
            text = text[:-1]
        text = text.replace(",", "")
        try:
            parsed = float(text)
        except ValueError:
            return None
        return parsed / 100.0 if pct else parsed
    return None


def safe_round(value: float | None, digits: int = 10) -> float | None:
    if value is None:
        return None
    return round(value, digits)


class SheetIndexer:
    def __init__(self, sheet: xw.Sheet) -> None:
        self.sheet = sheet
        self.last_row = max(1, int(sheet.used_range.last_cell.row))
        self.last_col = max(1, int(sheet.used_range.last_cell.column))
        raw_values = sheet.range((1, 1), (self.last_row, self.last_col)).value
        self.values = self._to_2d(raw_values, self.last_row, self.last_col)
        self.label_index: dict[str, list[tuple[int, int]]] = defaultdict(list)

        for row_idx, row in enumerate(self.values, start=1):
            for col_idx, value in enumerate(row, start=1):
                if isinstance(value, str):
                    key = normalize_label(value)
                    if key:
                        self.label_index[key].append((row_idx, col_idx))

    @staticmethod
    def _to_2d(raw_values: Any, rows: int, cols: int) -> list[list[Any]]:
        if rows == 1 and cols == 1:
            return [[raw_values]]
        if rows == 1:
            return [list(raw_values)]
        if cols == 1:
            return [[row_value] for row_value in raw_values]
        return [list(row_values) for row_values in raw_values]

    def get(self, row: int, col: int) -> Any:
        if row < 1 or col < 1 or row > self.last_row or col > self.last_col:
            return None
        return self.values[row - 1][col - 1]

    def find_labels_containing(self, token: str) -> list[tuple[int, int]]:
        token_norm = normalize_label(token)
        if not token_norm:
            return []
        matches: list[tuple[int, int]] = []
        for key, positions in self.label_index.items():
            if token_norm in key:
                matches.extend(positions)
        return matches

    def numeric_cells_to_right(self, row: int, start_col: int) -> list[tuple[int, int, float]]:
        results: list[tuple[int, int, float]] = []
        for col in range(start_col + 1, self.last_col + 1):
            parsed = to_float(self.get(row, col))
            if parsed is not None:
                results.append((row, col, parsed))
        return results


def parse_model_metadata(file_name: str) -> ModelMeta:
    base_name = Path(file_name).stem
    ticker_match = re.search(r"\s-\s([A-Za-z0-9]+)\s-\s", base_name)
    ticker = ticker_match.group(1).upper() if ticker_match else "UNKNOWN"

    period_match = re.search(r"\b(Early|Mid|Late)([A-Za-z]+)(\d{4})\b", base_name, flags=re.IGNORECASE)
    if period_match:
        phase_raw, month_raw, year_raw = period_match.groups()
        phase = phase_raw.capitalize()
        month_abbrev = month_raw[:3].capitalize()
        year = int(year_raw)
        month_number = parse_month_to_number(month_raw)
        day = PHASE_TO_DAY[phase.lower()]
        model_period = f"{phase}{month_abbrev}_{year}"
        model_date = date(year, month_number, day).isoformat()
    else:
        model_period = "UnknownPeriod"
        model_date = ""

    model = f"{ticker}_{model_period}" if model_period else ticker
    return ModelMeta(model=model, ticker=ticker, model_period=model_period, model_date=model_date)


def parse_month_to_number(month_text: str) -> int:
    month_key = month_text[:3].capitalize()
    month_map = {calendar.month_abbr[i]: i for i in range(1, 13)}
    if month_key not in month_map:
        raise ValueError(f"Unrecognized month token: {month_text}")
    return month_map[month_key]


def next_output_path(source_input_dir: Path, source_output_dir: Path) -> Path:
    base_name = f"{source_input_dir.name}_PARAM"
    candidate = source_output_dir / f"{base_name}.xlsx"
    if not candidate.exists():
        return candidate

    suffix = 1
    while True:
        candidate = source_output_dir / f"{base_name}.{suffix}.xlsx"
        if not candidate.exists():
            return candidate
        suffix += 1


def safe_close_workbook(workbook: xw.Book) -> None:
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


def get_sheet_if_exists(workbook: xw.Book, sheet_name: str) -> xw.Sheet | None:
    for sheet in workbook.sheets:
        if sheet.name.strip().lower() == sheet_name.strip().lower():
            return sheet
    return None


def find_max_anchor(indexer: SheetIndexer) -> tuple[int, int] | None:
    candidates = indexer.find_labels_containing("max")
    if not candidates:
        return None

    def score(candidate: tuple[int, int]) -> tuple[int, int]:
        row, col = candidate
        score_value = 0
        if normalize_label(str(indexer.get(row + 1, col) or "")) == "min":
            score_value += 3
        if to_float(indexer.get(row, col + 1)) is not None:
            score_value += 2
        if to_float(indexer.get(row + 1, col + 1)) is not None:
            score_value += 1
        return score_value, row

    return max(candidates, key=score)


def find_labeled_numeric(
    indexer: SheetIndexer,
    label_tokens: Iterable[str],
    anchor: tuple[int, int] | None = None,
) -> float | None:
    candidate_cells: list[tuple[int, int]] = []
    for token in label_tokens:
        candidate_cells.extend(indexer.find_labels_containing(token))

    if not candidate_cells:
        return None

    if anchor is not None:
        anchor_row, anchor_col = anchor
        candidate_cells.sort(key=lambda rc: abs(rc[0] - anchor_row) + abs(rc[1] - anchor_col))

    for row, col in candidate_cells:
        for row_offset, col_offset in ((0, 1), (1, 1), (0, 2), (1, 0), (-1, 1), (0, -1)):
            numeric = to_float(indexer.get(row + row_offset, col + col_offset))
            if numeric is not None:
                return numeric
    return None


def get_anchor_numeric(
    indexer: SheetIndexer,
    anchor: tuple[int, int],
    offsets: Iterable[tuple[int, int]],
) -> float | None:
    anchor_row, anchor_col = anchor
    for row_offset, col_offset in offsets:
        numeric = to_float(indexer.get(anchor_row + row_offset, anchor_col + col_offset))
        if numeric is not None:
            return numeric
    return None


def get_series_cells_from_label(
    indexer: SheetIndexer,
    tokens: Iterable[str],
) -> list[tuple[int, int, float]]:
    for token in tokens:
        for row, col in indexer.find_labels_containing(token):
            row_cells = indexer.numeric_cells_to_right(row, col)
            if row_cells:
                return row_cells
            if row + 1 <= indexer.last_row:
                row_cells = indexer.numeric_cells_to_right(row + 1, col)
                if row_cells:
                    return row_cells
    return []


def get_series_cells_from_anchor_window(
    indexer: SheetIndexer,
    anchor: tuple[int, int],
    row_offsets: Iterable[int],
    value_filter: str = "any",
) -> list[tuple[int, int, float]]:
    anchor_row, anchor_col = anchor
    left_bound = max(1, anchor_col - 30)
    right_bound = max(1, anchor_col - 1)

    best: list[tuple[int, int, float]] = []
    for row_offset in row_offsets:
        row = anchor_row + row_offset
        if row < 1 or row > indexer.last_row:
            continue
        series: list[tuple[int, int, float]] = []
        for col in range(left_bound, right_bound + 1):
            numeric = to_float(indexer.get(row, col))
            if numeric is None:
                continue
            if value_filter == "pct" and not (0.0 <= numeric <= 1.5):
                continue
            series.append((row, col, numeric))
        if len(series) > len(best):
            best = series

    return best


def get_last_value(series_cells: list[tuple[int, int, float]]) -> float | None:
    if not series_cells:
        return None
    return series_cells[-1][2]


def get_last_quarter_label(indexer: SheetIndexer, series_cells: list[tuple[int, int, float]]) -> str:
    if not series_cells:
        return ""
    row, col, _ = series_cells[-1]
    header = indexer.get(row - 1, col)
    if isinstance(header, str) and header.strip():
        return header.strip()
    return f"col_{col}"


def calc_range_width(max_value: float | None, min_value: float | None) -> float | None:
    if max_value is None or min_value is None:
        return None
    return max_value - min_value


def extract_empirical_rows(
    workbook: xw.Book,
    meta: ModelMeta,
    source_file: str,
) -> list[dict[str, Any]]:
    sheet = get_sheet_if_exists(workbook, "Empirical Model")
    if sheet is None:
        print(f"Skipped empirical extraction for {source_file}: sheet 'Empirical Model' not found")
        return []

    indexer = SheetIndexer(sheet)
    anchor = find_max_anchor(indexer)
    if anchor is None:
        print(f"Skipped empirical extraction for {source_file}: 'max' anchor not found")
        return []

    penetration_cells = get_series_cells_from_label(indexer, ("sales captured in db", "penetration"))
    if not penetration_cells:
        penetration_cells = get_series_cells_from_anchor_window(
            indexer=indexer,
            anchor=anchor,
            row_offsets=(-1, -2, -3, -4, -5, 1, 2),
            value_filter="pct",
        )

    penetration_cells = penetration_cells[-N_QUARTERS:]
    if not penetration_cells:
        print(f"Skipped empirical extraction for {source_file}: no penetration series found")
        return []

    quarterly_sales_cells = get_series_cells_from_label(indexer, ("quarterly sales", "db sales", "sales in db"))
    reported_sales_cells = get_series_cells_from_label(indexer, ("reported sales", "actual sales", "actual"))
    growth_cells = get_series_cells_from_label(indexer, ("growth rate",))

    forecast_max = get_anchor_numeric(indexer, anchor, offsets=((0, 1), (1, 1), (0, 2)))
    if forecast_max is None:
        forecast_max = find_labeled_numeric(indexer, ("max",), anchor=anchor)

    forecast_min = get_anchor_numeric(indexer, anchor, offsets=((1, 1), (2, 1), (1, 2)))
    if forecast_min is None:
        forecast_min = find_labeled_numeric(indexer, ("min",), anchor=anchor)

    helper_col = anchor[1] + 2
    helper_start_row = indexer.last_row + 2
    formula_cells: list[tuple[int, xw.Range]] = []

    for idx, quarters_used in enumerate(range(1, min(N_QUARTERS, len(penetration_cells)) + 1)):
        start_row, start_col, _ = penetration_cells[-quarters_used]
        end_row, end_col, _ = penetration_cells[-1]
        formula_cell = sheet.cells(helper_start_row + idx, helper_col)
        formula_cell.formula2 = f"=AVERAGE(R{start_row}C{start_col}:R{end_row}C{end_col})"
        formula_cells.append((quarters_used, formula_cell))

    if formula_cells:
        workbook.app.calculate()

    latest_quarterly_sales = get_last_value(quarterly_sales_cells)
    latest_reported_sales = get_last_value(reported_sales_cells)
    latest_growth = get_last_value(growth_cells)
    latest_captured = get_last_value(penetration_cells)
    last_quarter_label = get_last_quarter_label(indexer, penetration_cells)

    estimated_total_from_sheet = find_labeled_numeric(
        indexer,
        ("estimated total sold", "total sold", "tot fcst", "forecast"),
        anchor=anchor,
    )

    rows: list[dict[str, Any]] = []
    for quarters_used, formula_cell in formula_cells:
        avg_penetration_pct = to_float(formula_cell.value)
        forecast_value: float | None = None
        if avg_penetration_pct not in (None, 0.0) and latest_quarterly_sales is not None:
            forecast_value = latest_quarterly_sales / avg_penetration_pct
        elif estimated_total_from_sheet is not None:
            forecast_value = estimated_total_from_sheet

        row = {
            "model": meta.model,
            "ticker": meta.ticker,
            "model_period": meta.model_period,
            "model_date": meta.model_date,
            "method": "empirical",
            "parameter_name": "avg_penetration_pct",
            "parameter_value": avg_penetration_pct,
            "num_quarters_used": quarters_used,
            "last_quarter_used": last_quarter_label,
            "forecast_value": forecast_value,
            "actual_value": latest_reported_sales,
            "forecast_max": forecast_max,
            "forecast_min": forecast_min,
            "range_width": calc_range_width(forecast_max, forecast_min),
            "avg_penetration_pct": avg_penetration_pct,
            "quarterly_sales": latest_quarterly_sales,
            "reported_sales": latest_reported_sales,
            "growth_rate_pct": latest_growth,
            "sales_captured_in_db_pct": latest_captured,
            "source_file": source_file,
        }
        rows.append(row)

    return rows


def extract_regression_rows(
    workbook: xw.Book,
    meta: ModelMeta,
    source_file: str,
) -> list[dict[str, Any]]:
    sheet = get_sheet_if_exists(workbook, "Regression Model")
    if sheet is None:
        print(f"Skipped regression extraction for {source_file}: sheet 'Regression Model' not found")
        return []

    indexer = SheetIndexer(sheet)
    anchor = find_max_anchor(indexer)
    if anchor is None:
        print(f"Skipped regression extraction for {source_file}: 'max' anchor not found")
        return []

    anchor_row, anchor_col = anchor
    y_col = anchor_col - 7
    x_col = anchor_col - 11
    if x_col < 1 or y_col < 1:
        print(f"Skipped regression extraction for {source_file}: invalid anchor offsets")
        return []

    paired_points: list[tuple[int, float, float]] = []
    for row in range(1, anchor_row):
        x_value = to_float(indexer.get(row, x_col))
        y_value = to_float(indexer.get(row, y_col))
        if x_value is None or y_value is None:
            continue
        paired_points.append((row, x_value, y_value))

    if len(paired_points) < 2:
        print(f"Skipped regression extraction for {source_file}: insufficient x/y pairs")
        return []

    forecast_total_without_sa = find_labeled_numeric(
        indexer,
        ("tot fcst w/o sa", "tot fcst wo sa", "forecast total without sa", "tot fcst"),
        anchor=anchor,
    )
    actual_value = find_labeled_numeric(indexer, ("actual", "reported sales"), anchor=anchor)

    forecast_max = get_anchor_numeric(indexer, anchor, offsets=((0, 1), (1, 1), (0, 2)))
    if forecast_max is None:
        forecast_max = find_labeled_numeric(indexer, ("max",), anchor=anchor)

    forecast_min = get_anchor_numeric(indexer, anchor, offsets=((1, 1), (2, 1), (1, 2)))
    if forecast_min is None:
        forecast_min = find_labeled_numeric(indexer, ("min",), anchor=anchor)

    helper_start_row = indexer.last_row + 2
    intercept_col = anchor_col + 2
    slope_col = anchor_col + 3

    formula_cells: list[tuple[int, list[tuple[int, float, float]], xw.Range, xw.Range]] = []
    max_quarters = min(N_QUARTERS, len(paired_points))
    for idx, quarters_used in enumerate(range(2, max_quarters + 1)):
        selected = paired_points[-quarters_used:]
        start_row = selected[0][0]
        end_row = selected[-1][0]

        intercept_cell = sheet.cells(helper_start_row + idx, intercept_col)
        slope_cell = sheet.cells(helper_start_row + idx, slope_col)

        intercept_cell.formula2 = (
            f"=INTERCEPT(R{start_row}C{y_col}:R{end_row}C{y_col},R{start_row}C{x_col}:R{end_row}C{x_col})"
        )
        slope_cell.formula2 = (
            f"=SLOPE(R{start_row}C{y_col}:R{end_row}C{y_col},R{start_row}C{x_col}:R{end_row}C{x_col})"
        )
        formula_cells.append((quarters_used, selected, intercept_cell, slope_cell))

    if formula_cells:
        workbook.app.calculate()

    rows: list[dict[str, Any]] = []
    previous_signature: tuple[float | None, ...] | None = None
    for quarters_used, selected, intercept_cell, slope_cell in formula_cells:
        intercept = to_float(intercept_cell.value)
        slope = to_float(slope_cell.value)

        forecast_value = forecast_total_without_sa
        if forecast_value is None and intercept is not None and slope is not None:
            next_x = selected[-1][1] + 1.0
            forecast_value = intercept + slope * next_x

        current_max = forecast_max if forecast_max is not None else forecast_value
        current_min = forecast_min if forecast_min is not None else forecast_value
        range_width = calc_range_width(current_max, current_min)

        signature = (
            safe_round(intercept),
            safe_round(slope),
            safe_round(forecast_value),
            safe_round(current_max),
            safe_round(current_min),
        )
        if previous_signature == signature:
            continue
        previous_signature = signature

        row = {
            "model": meta.model,
            "ticker": meta.ticker,
            "model_period": meta.model_period,
            "model_date": meta.model_date,
            "method": "regression",
            "parameter_name": "num_quarters_used",
            "parameter_value": quarters_used,
            "num_quarters_used": quarters_used,
            "forecast_value": forecast_value,
            "actual_value": actual_value,
            "forecast_max": current_max,
            "forecast_min": current_min,
            "range_width": range_width,
            "intercept": intercept,
            "slope": slope,
            "source_file": source_file,
        }
        rows.append(row)

    return rows


def write_sheet(
    workbook: Workbook,
    sheet_name: str,
    columns: list[str],
    rows: list[dict[str, Any]],
) -> None:
    worksheet = workbook.create_sheet(title=sheet_name)
    worksheet.append(columns)

    for row in rows:
        worksheet.append([row.get(column) for column in columns])

    for cell in worksheet[1]:
        cell.font = Font(bold=True)

    worksheet.freeze_panes = "A2"
    worksheet.auto_filter.ref = f"A1:{get_column_letter(worksheet.max_column)}{max(worksheet.max_row, 1)}"

    for col_idx, column_name in enumerate(columns, start=1):
        max_length = len(column_name)
        for row_idx in range(2, worksheet.max_row + 1):
            value = worksheet.cell(row=row_idx, column=col_idx).value
            if value is None:
                continue
            value_length = len(str(value))
            if value_length > max_length:
                max_length = value_length
        worksheet.column_dimensions[get_column_letter(col_idx)].width = min(max(12, max_length + 2), 45)


def write_output_workbook(
    output_path: Path,
    empirical_rows: list[dict[str, Any]],
    regression_rows: list[dict[str, Any]],
) -> None:
    workbook = Workbook()
    default_sheet = workbook.active
    workbook.remove(default_sheet)

    write_sheet(workbook, "empirical_candidates", EMPIRICAL_COLUMNS, empirical_rows)
    write_sheet(workbook, "regression_candidates", REGRESSION_COLUMNS, regression_rows)

    workbook.save(output_path)


def collect_input_files(source_dir: Path) -> list[Path]:
    input_files: list[Path] = []
    for path in sorted(source_dir.iterdir()):
        if not path.is_file():
            continue
        if path.name.startswith("~"):
            print(f"Skipped {path.name}: temp file")
            continue
        if path.suffix.lower() != ".xlsx":
            print(f"Skipped {path.name}: not .xlsx")
            continue
        input_files.append(path)
    return input_files


def configure_excel_app() -> xw.App:
    app = xw.App(visible=False, add_book=False)
    try:
        app.display_alerts = False
    except Exception:
        pass
    try:
        app.screen_updating = False
    except Exception:
        pass
    try:
        app.calculation = "manual"
    except Exception:
        pass
    return app


def main() -> None:
    source_input_dir = Path(input_dir).expanduser().resolve()
    source_output_dir = Path(output_dir).expanduser().resolve()

    if not source_input_dir.exists():
        raise FileNotFoundError(f"Input directory does not exist: {source_input_dir}")
    if not source_input_dir.is_dir():
        raise NotADirectoryError(f"Input path is not a directory: {source_input_dir}")

    source_output_dir.mkdir(parents=True, exist_ok=True)
    output_path = next_output_path(source_input_dir, source_output_dir)

    files = collect_input_files(source_input_dir)
    if not files:
        print(f"No .xlsx files found in {source_input_dir}")

    app = configure_excel_app()
    processed_file_count = 0
    empirical_rows: list[dict[str, Any]] = []
    regression_rows: list[dict[str, Any]] = []

    try:
        for file_path in files:
            print(f"Processing file: {file_path.name}")
            workbook: xw.Book | None = None
            try:
                metadata = parse_model_metadata(file_path.name)
                workbook = app.books.open(str(file_path), update_links=False)

                empirical_rows.extend(
                    extract_empirical_rows(
                        workbook=workbook,
                        meta=metadata,
                        source_file=file_path.name,
                    )
                )
                regression_rows.extend(
                    extract_regression_rows(
                        workbook=workbook,
                        meta=metadata,
                        source_file=file_path.name,
                    )
                )

                processed_file_count += 1
            except Exception as exc:
                print(f"Skipped {file_path.name}: {exc}")
            finally:
                if workbook is not None:
                    safe_close_workbook(workbook)
    finally:
        try:
            app.quit()
        except Exception:
            pass

    write_output_workbook(
        output_path=output_path,
        empirical_rows=empirical_rows,
        regression_rows=regression_rows,
    )

    print(f"Output path: {output_path}")
    print(f"Files processed: {processed_file_count}")
    print(f"Empirical rows: {len(empirical_rows)}")
    print(f"Regression rows: {len(regression_rows)}")


if __name__ == "__main__":
    main()
