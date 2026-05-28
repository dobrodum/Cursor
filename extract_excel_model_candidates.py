from __future__ import annotations

import re
from dataclasses import dataclass
from datetime import date
from pathlib import Path
from typing import Any, Iterable, Optional

import xlwings as xw
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

# Configure these two paths before running.
input_dir = Path("./input")
output_dir = Path("./output")

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

EARLY_MID_LATE_DAY = {"early": 5, "mid": 15, "late": 25}
MONTH_TO_NUM = {
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
class SheetIndex:
    top_row: int
    left_col: int
    values: list[list[Any]]
    text_cells: list[tuple[int, int, str]]

    @property
    def row_count(self) -> int:
        return len(self.values)

    @property
    def col_count(self) -> int:
        return max((len(row) for row in self.values), default=0)

    def get_abs_value(self, row: int, col: int) -> Any:
        row_idx = row - self.top_row
        col_idx = col - self.left_col
        if row_idx < 0 or col_idx < 0:
            return None
        if row_idx >= len(self.values):
            return None
        current_row = self.values[row_idx]
        if col_idx >= len(current_row):
            return None
        return current_row[col_idx]


def normalize_text(value: Any) -> str:
    if value is None:
        return ""
    text = str(value).strip().lower()
    text = text.replace("_", " ")
    text = re.sub(r"\s+", " ", text)
    return text


def to_2d(values: Any) -> list[list[Any]]:
    if values is None:
        return []
    if isinstance(values, list):
        if not values:
            return []
        if isinstance(values[0], list):
            return values
        return [values]
    return [[values]]


def to_float(value: Any) -> Optional[float]:
    if value is None or value == "":
        return None
    if isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        return float(value)
    if isinstance(value, str):
        cleaned = value.replace(",", "").replace("%", "").strip()
        if cleaned == "":
            return None
        try:
            return float(cleaned)
        except ValueError:
            return None
    return None


def is_numeric(value: Any) -> bool:
    return to_float(value) is not None


def build_sheet_index(sheet: xw.Sheet) -> SheetIndex:
    used = sheet.used_range
    values = to_2d(used.value)
    text_cells: list[tuple[int, int, str]] = []
    for r_idx, row_values in enumerate(values):
        for c_idx, value in enumerate(row_values):
            if isinstance(value, str) and value.strip():
                abs_row = used.row + r_idx
                abs_col = used.column + c_idx
                text_cells.append((abs_row, abs_col, normalize_text(value)))
    return SheetIndex(
        top_row=used.row,
        left_col=used.column,
        values=values,
        text_cells=text_cells,
    )


def find_sheet(wb: xw.Book, sheet_name: str) -> Optional[xw.Sheet]:
    for sheet in wb.sheets:
        if sheet.name.strip().lower() == sheet_name.lower():
            return sheet
    return None


def find_anchor_cell(sheet: xw.Sheet, index: SheetIndex, anchor_text: str = "max") -> Optional[xw.Range]:
    exact_matches = [
        (row, col)
        for row, col, text in index.text_cells
        if text == anchor_text or text.startswith(f"{anchor_text} ")
    ]
    if exact_matches:
        row, col = exact_matches[0]
        return sheet.range((row, col))

    contains_matches = [(row, col) for row, col, text in index.text_cells if anchor_text in text]
    if not contains_matches:
        return None
    row, col = contains_matches[0]
    return sheet.range((row, col))


def closest_text_cell(
    index: SheetIndex,
    anchor_row: int,
    anchor_col: int,
    labels: Iterable[str],
    max_row_distance: int = 160,
    max_col_distance: int = 40,
) -> Optional[tuple[int, int]]:
    normalized_labels = [normalize_text(label) for label in labels]
    best: Optional[tuple[int, int, int]] = None
    for row, col, text in index.text_cells:
        if not any(label in text for label in normalized_labels):
            continue
        row_dist = abs(row - anchor_row)
        col_dist = abs(col - anchor_col)
        if row_dist > max_row_distance or col_dist > max_col_distance:
            continue
        score = row_dist * 3 + col_dist
        if best is None or score < best[2]:
            best = (row, col, score)
    if best is None:
        return None
    return best[0], best[1]


def value_cell_near_label(index: SheetIndex, label_row: int, label_col: int) -> Optional[tuple[int, int]]:
    # Try same row, right side first (most common dashboard layout).
    for col in range(label_col + 1, label_col + 8):
        value = index.get_abs_value(label_row, col)
        if value not in (None, ""):
            return label_row, col

    # Then same column, below the label.
    for row in range(label_row + 1, label_row + 6):
        value = index.get_abs_value(row, label_col)
        if value not in (None, ""):
            return row, label_col

    # Lastly, check right-below positions.
    for row in range(label_row + 1, label_row + 6):
        for col in range(label_col + 1, label_col + 8):
            value = index.get_abs_value(row, col)
            if value not in (None, ""):
                return row, col
    return None


def locate_metric_cell(
    sheet: xw.Sheet,
    index: SheetIndex,
    anchor_row: int,
    anchor_col: int,
    label_candidates: Iterable[str],
    default_offset: Optional[tuple[int, int]] = None,
) -> Optional[xw.Range]:
    label_cell = closest_text_cell(index, anchor_row, anchor_col, label_candidates)
    if label_cell is not None:
        value_cell = value_cell_near_label(index, label_cell[0], label_cell[1])
        if value_cell is not None:
            return sheet.range(value_cell)
    if default_offset is None:
        return None
    return sheet.range((anchor_row + default_offset[0], anchor_col + default_offset[1]))


def parse_file_label(file_path: Path) -> FileLabel:
    # Example: MedMiner_Model - AORT - MidJan2026_Send.xlsx
    match = re.search(
        r"-\s*(?P<ticker>[A-Za-z0-9]+)\s*-\s*(?P<period>(?:Early|Mid|Late)[A-Za-z]{3}\d{4})",
        file_path.stem,
        flags=re.IGNORECASE,
    )
    if not match:
        raise ValueError("filename does not match expected model naming convention")

    ticker = match.group("ticker").upper()
    raw_period = match.group("period")
    period_match = re.match(
        r"(?P<phase>Early|Mid|Late)(?P<month>[A-Za-z]{3})(?P<year>\d{4})",
        raw_period,
        flags=re.IGNORECASE,
    )
    if not period_match:
        raise ValueError("filename model period could not be parsed")

    phase = period_match.group("phase").title()
    month_abbr = period_match.group("month").title()
    month_num = MONTH_TO_NUM.get(month_abbr.lower())
    if month_num is None:
        raise ValueError(f"unsupported month abbreviation: {month_abbr}")
    year_num = int(period_match.group("year"))
    day_num = EARLY_MID_LATE_DAY[phase.lower()]

    model_period = f"{phase}{month_abbr}_{year_num}"
    model_date = date(year_num, month_num, day_num).isoformat()
    model = f"{ticker}_{model_period}"
    return FileLabel(model=model, ticker=ticker, model_period=model_period, model_date=model_date)


def safe_close_workbook(wb: xw.Book) -> None:
    try:
        wb.close(save=False)
        return
    except TypeError:
        pass
    except Exception:
        pass

    # Safe fallback for environments where save kwarg isn't accepted.
    try:
        wb.api.Close(SaveChanges=False)
        return
    except Exception:
        pass

    try:
        wb.api.Close(False)
    except Exception:
        try:
            wb.close()
        except Exception:
            pass


def next_output_path(source_input_dir: Path, target_output_dir: Path) -> Path:
    base_name = f"{source_input_dir.name}_PARAM"
    candidate = target_output_dir / f"{base_name}.xlsx"
    suffix_number = 1
    while candidate.exists():
        candidate = target_output_dir / f"{base_name}.{suffix_number}.xlsx"
        suffix_number += 1
    return candidate


def find_penetration_series(
    index: SheetIndex,
    anchor_row: int,
    anchor_col: int,
) -> list[tuple[int, int, float]]:
    candidate_rows: list[tuple[int, list[tuple[int, int, float]]]] = []

    for label_row, label_col, label_text in index.text_cells:
        if "penetration" not in label_text or "avg" in label_text:
            continue
        if abs(label_row - anchor_row) > 160:
            continue

        row_values: list[tuple[int, int, float]] = []
        for col in range(label_col + 1, anchor_col + 1):
            value = to_float(index.get_abs_value(label_row, col))
            if value is not None:
                row_values.append((label_row, col, value))
        if len(row_values) >= 3:
            candidate_rows.append((len(row_values), row_values))

    if candidate_rows:
        candidate_rows.sort(key=lambda x: x[0], reverse=True)
        return candidate_rows[0][1]

    # Fallback: choose the longest numeric run in rows near anchor.
    best_series: list[tuple[int, int, float]] = []
    for abs_row in range(anchor_row - 80, anchor_row + 1):
        row_series: list[tuple[int, int, float]] = []
        for abs_col in range(index.left_col, anchor_col + 1):
            value = to_float(index.get_abs_value(abs_row, abs_col))
            if value is not None:
                row_series.append((abs_row, abs_col, value))
        if len(row_series) > len(best_series):
            best_series = row_series
    return best_series


def quarter_label_for_col(index: SheetIndex, value_row: int, value_col: int) -> Optional[str]:
    for lookup_row in range(value_row - 1, max(index.top_row - 1, value_row - 4), -1):
        label = index.get_abs_value(lookup_row, value_col)
        if label not in (None, ""):
            return str(label)
    return None


def process_empirical_sheet(
    wb: xw.Book,
    file_label: FileLabel,
    source_file: str,
) -> list[dict[str, Any]]:
    sheet = find_sheet(wb, "Empirical Model")
    if sheet is None:
        return []

    index = build_sheet_index(sheet)
    anchor = find_anchor_cell(sheet, index, anchor_text="max")
    if anchor is None:
        return []

    anchor_row = anchor.row
    anchor_col = anchor.column

    num_quarters_cell = locate_metric_cell(
        sheet,
        index,
        anchor_row,
        anchor_col,
        ["num quarters used", "quarters used", "num quarters"],
    )
    avg_penetration_cell = locate_metric_cell(
        sheet,
        index,
        anchor_row,
        anchor_col,
        ["avg penetration", "average penetration"],
    )
    forecast_value_cell = locate_metric_cell(
        sheet,
        index,
        anchor_row,
        anchor_col,
        ["estimated total sold", "estimated total", "total sold"],
        default_offset=(-1, 1),
    )
    actual_value_cell = locate_metric_cell(
        sheet,
        index,
        anchor_row,
        anchor_col,
        ["reported sales", "actual sales", "actual value"],
        default_offset=(2, 1),
    )
    forecast_max_cell = locate_metric_cell(
        sheet,
        index,
        anchor_row,
        anchor_col,
        ["max"],
        default_offset=(0, 1),
    )
    forecast_min_cell = locate_metric_cell(
        sheet,
        index,
        anchor_row,
        anchor_col,
        ["min"],
        default_offset=(1, 1),
    )
    quarterly_sales_cell = locate_metric_cell(
        sheet,
        index,
        anchor_row,
        anchor_col,
        ["quarterly sales"],
    )
    reported_sales_cell = locate_metric_cell(
        sheet,
        index,
        anchor_row,
        anchor_col,
        ["reported sales"],
    )
    growth_rate_cell = locate_metric_cell(
        sheet,
        index,
        anchor_row,
        anchor_col,
        ["growth rate"],
    )
    sales_captured_cell = locate_metric_cell(
        sheet,
        index,
        anchor_row,
        anchor_col,
        ["sales captured in db", "captured in db", "captured %"],
    )

    penetration_series = find_penetration_series(index, anchor_row, anchor_col)

    # Helper cell used only if no dedicated avg penetration input/output cell is detected.
    helper_avg_cell = sheet.range((anchor_row + 25, anchor_col + 8))
    avg_formula_target = avg_penetration_cell or helper_avg_cell

    empirical_rows: list[dict[str, Any]] = []
    n_quarters = 10
    for n in range(1, n_quarters + 1):
        if num_quarters_cell is not None:
            num_quarters_cell.value = n

        if len(penetration_series) >= n:
            selected = penetration_series[-n:]
            start_row, start_col, _ = selected[0]
            end_row, end_col, _ = selected[-1]
            avg_formula_target.formula2 = (
                f"=AVERAGE(R{start_row}C{start_col}:R{end_row}C{end_col})"
            )

        wb.app.calculate()

        avg_penetration_pct = to_float(avg_formula_target.value)
        forecast_value = to_float(forecast_value_cell.value if forecast_value_cell else None)
        actual_value = to_float(actual_value_cell.value if actual_value_cell else None)
        forecast_max = to_float(forecast_max_cell.value if forecast_max_cell else None)
        forecast_min = to_float(forecast_min_cell.value if forecast_min_cell else None)

        if (
            avg_penetration_pct is None
            and forecast_value is None
            and forecast_max is None
            and forecast_min is None
        ):
            continue

        range_width = (
            forecast_max - forecast_min
            if forecast_max is not None and forecast_min is not None
            else None
        )

        last_quarter_used = None
        if penetration_series:
            _, last_col, _ = penetration_series[-1]
            last_quarter_used = quarter_label_for_col(index, penetration_series[-1][0], last_col)

        empirical_rows.append(
            {
                "model": file_label.model,
                "ticker": file_label.ticker,
                "model_period": file_label.model_period,
                "model_date": file_label.model_date,
                "method": "empirical",
                "parameter_name": "avg_penetration_pct",
                "parameter_value": avg_penetration_pct,
                "num_quarters_used": int(num_quarters_cell.value) if num_quarters_cell else n,
                "last_quarter_used": last_quarter_used,
                "forecast_value": forecast_value,
                "actual_value": actual_value,
                "forecast_max": forecast_max,
                "forecast_min": forecast_min,
                "range_width": range_width,
                "avg_penetration_pct": avg_penetration_pct,
                "quarterly_sales": to_float(quarterly_sales_cell.value if quarterly_sales_cell else None),
                "reported_sales": to_float(reported_sales_cell.value if reported_sales_cell else None),
                "growth_rate_pct": to_float(growth_rate_cell.value if growth_rate_cell else None),
                "sales_captured_in_db_pct": to_float(
                    sales_captured_cell.value if sales_captured_cell else None
                ),
                "source_file": source_file,
            }
        )

    return empirical_rows


def collect_regression_rows(index: SheetIndex, y_col: int, x_col: int, max_row: int) -> list[int]:
    valid_rows: list[int] = []
    for row in range(index.top_row, max_row + 1):
        y_val = index.get_abs_value(row, y_col)
        x_val = index.get_abs_value(row, x_col)
        if is_numeric(y_val) and is_numeric(x_val):
            valid_rows.append(row)
    return valid_rows


def process_regression_sheet(
    wb: xw.Book,
    file_label: FileLabel,
    source_file: str,
) -> list[dict[str, Any]]:
    sheet = find_sheet(wb, "Regression Model")
    if sheet is None:
        return []

    index = build_sheet_index(sheet)
    anchor = find_anchor_cell(sheet, index, anchor_text="max")
    if anchor is None:
        return []

    anchor_row = anchor.row
    anchor_col = anchor.column

    y_col = anchor_col - 7
    x_col = anchor_col - 11
    data_rows = collect_regression_rows(index, y_col=y_col, x_col=x_col, max_row=anchor_row)
    if len(data_rows) < 2:
        return []

    num_quarters_cell = locate_metric_cell(
        sheet,
        index,
        anchor_row,
        anchor_col,
        ["num quarters used", "quarters used", "num quarters"],
    )
    forecast_total_cell = locate_metric_cell(
        sheet,
        index,
        anchor_row,
        anchor_col,
        ["tot fcst w/o sa", "tot fcst wo sa", "total fcst w/o sa", "total forecast w/o sa"],
        default_offset=(-1, 1),
    )
    actual_value_cell = locate_metric_cell(
        sheet,
        index,
        anchor_row,
        anchor_col,
        ["actual value", "actual sales", "reported sales"],
    )
    forecast_max_cell = locate_metric_cell(
        sheet,
        index,
        anchor_row,
        anchor_col,
        ["max"],
        default_offset=(0, 1),
    )
    forecast_min_cell = locate_metric_cell(
        sheet,
        index,
        anchor_row,
        anchor_col,
        ["min"],
        default_offset=(1, 1),
    )

    helper_intercept = sheet.range((anchor_row + 25, anchor_col + 8))
    helper_slope = sheet.range((anchor_row + 26, anchor_col + 8))

    regression_rows: list[dict[str, Any]] = []
    previous_signature: Optional[tuple[Any, ...]] = None
    n_quarters = min(10, len(data_rows))

    for n in range(1, n_quarters + 1):
        rows_subset = data_rows[-n:]
        start_row = rows_subset[0]
        end_row = rows_subset[-1]

        helper_intercept.formula2 = (
            f"=INTERCEPT(R{start_row}C{y_col}:R{end_row}C{y_col},"
            f"R{start_row}C{x_col}:R{end_row}C{x_col})"
        )
        helper_slope.formula2 = (
            f"=SLOPE(R{start_row}C{y_col}:R{end_row}C{y_col},"
            f"R{start_row}C{x_col}:R{end_row}C{x_col})"
        )

        if num_quarters_cell is not None:
            num_quarters_cell.value = n

        wb.app.calculate()

        intercept = to_float(helper_intercept.value)
        slope = to_float(helper_slope.value)
        forecast_max = to_float(forecast_max_cell.value if forecast_max_cell else None)
        forecast_min = to_float(forecast_min_cell.value if forecast_min_cell else None)
        forecast_value = to_float(forecast_total_cell.value if forecast_total_cell else None)
        actual_value = to_float(actual_value_cell.value if actual_value_cell else None)

        if forecast_value is None and intercept is not None and slope is not None:
            latest_x = to_float(index.get_abs_value(rows_subset[-1], x_col))
            if latest_x is not None:
                forecast_value = intercept + slope * latest_x

        if (
            intercept is None
            and slope is None
            and forecast_value is None
            and forecast_max is None
            and forecast_min is None
        ):
            continue

        range_width = (
            forecast_max - forecast_min
            if forecast_max is not None and forecast_min is not None
            else None
        )

        num_quarters_used = int(num_quarters_cell.value) if num_quarters_cell else n
        signature = (
            num_quarters_used,
            round(intercept, 8) if intercept is not None else None,
            round(slope, 8) if slope is not None else None,
            round(forecast_value, 8) if forecast_value is not None else None,
            round(forecast_max, 8) if forecast_max is not None else None,
            round(forecast_min, 8) if forecast_min is not None else None,
        )
        if signature == previous_signature:
            continue
        previous_signature = signature

        regression_rows.append(
            {
                "model": file_label.model,
                "ticker": file_label.ticker,
                "model_period": file_label.model_period,
                "model_date": file_label.model_date,
                "method": "regression",
                "parameter_name": "num_quarters_used",
                "parameter_value": num_quarters_used,
                "num_quarters_used": num_quarters_used,
                "forecast_value": forecast_value,
                "actual_value": actual_value if actual_value is not None else "",
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
    output_path: Path,
    empirical_rows: list[dict[str, Any]],
    regression_rows: list[dict[str, Any]],
) -> None:
    out_wb = Workbook()
    default_sheet = out_wb.active
    out_wb.remove(default_sheet)

    empirical_ws = out_wb.create_sheet("empirical_candidates")
    regression_ws = out_wb.create_sheet("regression_candidates")

    write_sheet(empirical_ws, EMPIRICAL_COLUMNS, empirical_rows)
    write_sheet(regression_ws, REGRESSION_COLUMNS, regression_rows)

    out_wb.save(output_path)


def write_sheet(
    ws: Any,
    columns: list[str],
    rows: list[dict[str, Any]],
) -> None:
    ws.append(columns)
    for row_data in rows:
        ws.append([row_data.get(col, "") for col in columns])

    for header_cell in ws[1]:
        header_cell.font = Font(bold=True)

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions

    for col_idx, col_name in enumerate(columns, start=1):
        max_len = len(col_name)
        for row_idx in range(2, ws.max_row + 1):
            value = ws.cell(row=row_idx, column=col_idx).value
            if value is None:
                continue
            text = str(value)
            if len(text) > max_len:
                max_len = len(text)
        ws.column_dimensions[get_column_letter(col_idx)].width = min(max(max_len + 2, 12), 42)


def main() -> None:
    source_dir = Path(input_dir).expanduser().resolve()
    target_dir = Path(output_dir).expanduser().resolve()

    if not source_dir.exists():
        print(f"Input directory not found: {source_dir}")
        return

    target_dir.mkdir(parents=True, exist_ok=True)
    output_path = next_output_path(source_dir, target_dir)

    files_processed = 0
    empirical_rows: list[dict[str, Any]] = []
    regression_rows: list[dict[str, Any]] = []

    with xw.App(visible=False, add_book=False) as app:
        app.display_alerts = False
        app.screen_updating = False
        try:
            app.calculation = "manual"
        except Exception:
            pass

        for file_path in sorted(source_dir.iterdir()):
            if not file_path.is_file():
                print(f"skipped: {file_path.name} (not a file)")
                continue
            if file_path.name.startswith("~"):
                print(f"skipped: {file_path.name} (temporary file)")
                continue
            if file_path.suffix.lower() != ".xlsx":
                print(f"skipped: {file_path.name} (not .xlsx)")
                continue

            try:
                file_label = parse_file_label(file_path)
            except Exception as exc:
                print(f"skipped: {file_path.name} ({exc})")
                continue

            wb = None
            try:
                wb = app.books.open(str(file_path), update_links=False)
                emp_rows = process_empirical_sheet(wb, file_label, file_path.name)
                reg_rows = process_regression_sheet(wb, file_label, file_path.name)
                empirical_rows.extend(emp_rows)
                regression_rows.extend(reg_rows)
                files_processed += 1
                print(f"processed: {file_path.name}")
            except Exception as exc:
                print(f"skipped: {file_path.name} (processing error: {exc})")
            finally:
                if wb is not None:
                    safe_close_workbook(wb)

    write_output_workbook(output_path, empirical_rows, regression_rows)

    print(f"output: {output_path}")
    print(f"files_processed: {files_processed}")
    print(f"empirical_rows: {len(empirical_rows)}")
    print(f"regression_rows: {len(regression_rows)}")


if __name__ == "__main__":
    main()
