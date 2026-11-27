#!/usr/bin/env python3
from __future__ import annotations

import re
from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Sequence, Tuple

import xlwings as xw
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

# Required inputs
input_dir = Path("input")
output_dir = Path("output")

N_QUARTERS = 10
HEADER_FONT = Font(bold=True)

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

PERIOD_TO_DAY = {"early": 5, "mid": 15, "late": 25}


@dataclass
class FileLabel:
    model: str
    ticker: str
    model_period: str
    model_date: str
    source_file: str


@dataclass
class SheetContext:
    sheet: Any
    top_row: int
    left_col: int
    values: List[List[Any]]
    labels: Dict[str, List[Tuple[int, int]]]
    normalized_by_cell: Dict[Tuple[int, int], str]


def normalize_text(value: Any) -> str:
    if value is None:
        return ""
    text = str(value).strip().lower()
    if not text:
        return ""
    text = re.sub(r"[^a-z0-9]+", " ", text)
    text = re.sub(r"\s+", " ", text).strip()
    return text


def as_2d_list(value: Any) -> List[List[Any]]:
    if value is None:
        return []
    if not isinstance(value, list):
        return [[value]]
    if value and isinstance(value[0], list):
        return value
    return [value]


def as_number(value: Any) -> Optional[float]:
    if value is None or isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        return float(value)
    text = str(value).strip().replace(",", "")
    if not text:
        return None
    if text.endswith("%"):
        text = text[:-1]
        try:
            return float(text) / 100.0
        except ValueError:
            return None
    try:
        return float(text)
    except ValueError:
        return None


def clean_for_output(value: Any) -> Any:
    if value is None:
        return ""
    if isinstance(value, float):
        if value != value:
            return ""
        return value
    if isinstance(value, date):
        return value.isoformat()
    return value


def month_to_number(month_token: str) -> Optional[int]:
    token = month_token.strip()
    if not token:
        return None
    for pattern in ("%b", "%B"):
        try:
            return datetime.strptime(token.title(), pattern).month
        except ValueError:
            continue
    if len(token) >= 3:
        try:
            return datetime.strptime(token[:3].title(), "%b").month
        except ValueError:
            return None
    return None


def parse_file_label(file_name: str) -> FileLabel:
    stem = Path(file_name).stem
    parts = [part.strip() for part in stem.split(" - ")]

    ticker = parts[1] if len(parts) >= 2 else ""
    period_piece = parts[2] if len(parts) >= 3 else ""
    period_token = re.split(r"[_\s]", period_piece)[0]

    model_period = ""
    model_date = ""

    match = re.match(r"(?i)(early|mid|late)([a-z]+)(\d{4})$", period_token)
    if match:
        period_word = match.group(1).lower()
        month_word = match.group(2)
        year = int(match.group(3))
        month_number = month_to_number(month_word)
        if month_number is not None:
            month_label = datetime(2000, month_number, 1).strftime("%b")
            model_period = f"{period_word.title()}{month_label}_{year}"
            model_day = PERIOD_TO_DAY[period_word]
            model_date = date(year, month_number, model_day).isoformat()

    model = f"{ticker}_{model_period}" if ticker and model_period else ticker
    return FileLabel(
        model=model,
        ticker=ticker,
        model_period=model_period,
        model_date=model_date,
        source_file=file_name,
    )


def make_output_path(input_path: Path, output_path: Path) -> Path:
    output_path.mkdir(parents=True, exist_ok=True)
    base_name = f"{input_path.name}_PARAM"
    candidate = output_path / f"{base_name}.xlsx"
    if not candidate.exists():
        return candidate
    suffix = 1
    while True:
        candidate = output_path / f"{base_name}.{suffix}.xlsx"
        if not candidate.exists():
            return candidate
        suffix += 1


def build_sheet_context(sheet: Any) -> SheetContext:
    used = sheet.used_range
    values = as_2d_list(used.value)
    top_row = int(used.row)
    left_col = int(used.column)

    labels: Dict[str, List[Tuple[int, int]]] = {}
    normalized_by_cell: Dict[Tuple[int, int], str] = {}

    for row_index, row_values in enumerate(values):
        row_number = top_row + row_index
        for col_index, cell_value in enumerate(row_values):
            col_number = left_col + col_index
            normalized = normalize_text(cell_value)
            if not normalized:
                continue
            labels.setdefault(normalized, []).append((row_number, col_number))
            normalized_by_cell[(row_number, col_number)] = normalized

    return SheetContext(
        sheet=sheet,
        top_row=top_row,
        left_col=left_col,
        values=values,
        labels=labels,
        normalized_by_cell=normalized_by_cell,
    )


def get_cell_value(sheet: Any, row: int, col: int) -> Any:
    if row < 1 or col < 1:
        return None
    return sheet.cells(row, col).value


def find_value_cell_right_of_label(
    context: SheetContext, aliases: Sequence[str]
) -> Optional[Tuple[int, int]]:
    normalized_aliases = [normalize_text(alias) for alias in aliases]
    for alias in normalized_aliases:
        if not alias:
            continue
        positions = context.labels.get(alias)
        if positions:
            row, col = positions[0]
            return row, col + 1

    # Fallback to fuzzy label matching in the already-scanned label map.
    for alias in normalized_aliases:
        if not alias:
            continue
        for label, positions in context.labels.items():
            if alias in label or label in alias:
                row, col = positions[0]
                return row, col + 1
    return None


def find_max_anchor(context: SheetContext) -> Optional[Tuple[int, int]]:
    candidates = context.labels.get("max", [])
    if not candidates:
        return None

    def score(candidate: Tuple[int, int]) -> Tuple[int, int]:
        row, col = candidate
        below_label = context.normalized_by_cell.get((row + 1, col), "")
        has_min_below = 1 if below_label == "min" else 0
        # Prefer the "max" that looks paired with "min", then the lower one.
        return (has_min_below, row)

    return max(candidates, key=score)


def safe_close_workbook(workbook: Any) -> None:
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
        pass


def read_value(sheet: Any, location: Optional[Tuple[int, int]]) -> Any:
    if location is None:
        return None
    row, col = location
    return get_cell_value(sheet, row, col)


def rounded_signature(values: Iterable[Optional[float]]) -> Tuple[Optional[float], ...]:
    signature: List[Optional[float]] = []
    for value in values:
        if value is None:
            signature.append(None)
        else:
            signature.append(round(value, 8))
    return tuple(signature)


def extract_empirical_rows(workbook: Any, file_label: FileLabel) -> List[Dict[str, Any]]:
    rows: List[Dict[str, Any]] = []
    try:
        sheet = workbook.sheets["Empirical Model"]
    except Exception:
        print(f"Skipped empirical for {file_label.source_file}: missing 'Empirical Model' sheet")
        return rows

    context = build_sheet_context(sheet)
    anchor = find_max_anchor(context)
    if anchor is None:
        print(f"Skipped empirical for {file_label.source_file}: could not find 'max' anchor")
        return rows

    anchor_row, anchor_col = anchor
    x_col = max(1, anchor_col - 11)
    y_col = max(1, anchor_col - 7)
    quarter_col = max(1, x_col - 1)

    num_quarters_input = find_value_cell_right_of_label(
        context,
        ["num quarters used", "num quarters", "quarters used"],
    )
    estimated_total_sold_cell = find_value_cell_right_of_label(
        context,
        ["estimated total sold", "est total sold", "total sold estimate", "tot fcst", "forecast"],
    )
    reported_sales_cell = find_value_cell_right_of_label(
        context,
        ["reported sales", "actual sales", "reported"],
    )
    growth_rate_cell = find_value_cell_right_of_label(
        context,
        ["growth rate pct", "growth rate", "growth pct", "growth"],
    )
    sales_captured_cell = find_value_cell_right_of_label(
        context,
        ["sales captured in db pct", "sales captured in db", "captured in db"],
    )
    last_quarter_cell = find_value_cell_right_of_label(
        context,
        ["last quarter used", "last qtr used", "last quarter"],
    )

    min_value_cell = find_value_cell_right_of_label(context, ["min"])
    forecast_max_cell = (anchor_row, anchor_col + 1)
    forecast_min_cell = min_value_cell or (anchor_row + 1, anchor_col + 1)

    avg_pen_helper = sheet.cells(2, 500)

    for n_quarters in range(1, N_QUARTERS + 1):
        start_row = anchor_row - n_quarters + 1
        if start_row < 1:
            break

        if num_quarters_input is not None:
            nr, nc = num_quarters_input
            sheet.cells(nr, nc).value = n_quarters

        avg_pen_helper.formula2 = (
            f'=IFERROR(AVERAGE(R{start_row}C{y_col}:R{anchor_row}C{y_col}/'
            f'R{start_row}C{x_col}:R{anchor_row}C{x_col}),"")'
        )
        workbook.app.calculate()

        avg_penetration = as_number(avg_pen_helper.value)
        forecast_value = as_number(read_value(sheet, estimated_total_sold_cell))
        actual_value = as_number(read_value(sheet, reported_sales_cell))
        forecast_max = as_number(read_value(sheet, forecast_max_cell))
        forecast_min = as_number(read_value(sheet, forecast_min_cell))

        quarterly_sales = as_number(get_cell_value(sheet, anchor_row, x_col))
        reported_sales = as_number(get_cell_value(sheet, anchor_row, y_col))

        if actual_value is None:
            actual_value = reported_sales

        if forecast_value is None and avg_penetration not in (None, 0) and actual_value is not None:
            forecast_value = actual_value / avg_penetration

        last_quarter_used = read_value(sheet, last_quarter_cell)
        if last_quarter_used is None:
            last_quarter_used = get_cell_value(sheet, start_row, quarter_col)

        growth_rate_pct = as_number(read_value(sheet, growth_rate_cell))
        sales_captured_in_db_pct = as_number(read_value(sheet, sales_captured_cell))
        range_width = (
            (forecast_max - forecast_min)
            if forecast_max is not None and forecast_min is not None
            else None
        )

        row = {
            "model": file_label.model,
            "ticker": file_label.ticker,
            "model_period": file_label.model_period,
            "model_date": file_label.model_date,
            "method": "empirical",
            "parameter_name": "avg_penetration_pct",
            "parameter_value": avg_penetration,
            "num_quarters_used": n_quarters,
            "last_quarter_used": last_quarter_used,
            "forecast_value": forecast_value,
            "actual_value": actual_value,
            "forecast_max": forecast_max,
            "forecast_min": forecast_min,
            "range_width": range_width,
            "avg_penetration_pct": avg_penetration,
            "quarterly_sales": quarterly_sales,
            "reported_sales": reported_sales,
            "growth_rate_pct": growth_rate_pct,
            "sales_captured_in_db_pct": sales_captured_in_db_pct,
            "source_file": file_label.source_file,
        }

        has_signal = any(
            row[key] not in ("", None)
            for key in (
                "parameter_value",
                "forecast_value",
                "actual_value",
                "forecast_max",
                "forecast_min",
            )
        )
        if has_signal:
            rows.append(row)

    avg_pen_helper.clear_contents()
    return rows


def extract_regression_rows(workbook: Any, file_label: FileLabel) -> List[Dict[str, Any]]:
    rows: List[Dict[str, Any]] = []
    try:
        sheet = workbook.sheets["Regression Model"]
    except Exception:
        print(f"Skipped regression for {file_label.source_file}: missing 'Regression Model' sheet")
        return rows

    context = build_sheet_context(sheet)
    anchor = find_max_anchor(context)
    if anchor is None:
        print(f"Skipped regression for {file_label.source_file}: could not find 'max' anchor")
        return rows

    anchor_row, anchor_col = anchor
    y_col = max(1, anchor_col - 7)
    x_col = max(1, anchor_col - 11)

    num_quarters_input = find_value_cell_right_of_label(
        context,
        ["num quarters used", "num quarters", "quarters used"],
    )
    forecast_total_cell = find_value_cell_right_of_label(
        context,
        ["tot fcst w o sa", "tot fcst w/o sa", "total forecast without sa", "tot fcst without sa"],
    )
    actual_value_cell = find_value_cell_right_of_label(
        context,
        ["actual value", "actual sales", "reported sales", "actual"],
    )
    min_value_cell = find_value_cell_right_of_label(context, ["min"])

    forecast_max_cell = (anchor_row, anchor_col + 1)
    forecast_min_cell = min_value_cell or (anchor_row + 1, anchor_col + 1)

    intercept_helper = sheet.cells(2, 501)
    slope_helper = sheet.cells(2, 502)

    previous_signature: Optional[Tuple[Optional[float], ...]] = None

    for n_quarters in range(1, N_QUARTERS + 1):
        start_row = anchor_row - n_quarters + 1
        if start_row < 1:
            break

        if num_quarters_input is not None:
            nr, nc = num_quarters_input
            sheet.cells(nr, nc).value = n_quarters

        intercept_helper.formula2 = (
            f'=IFERROR(INTERCEPT(R{start_row}C{y_col}:R{anchor_row}C{y_col},'
            f'R{start_row}C{x_col}:R{anchor_row}C{x_col}),"")'
        )
        slope_helper.formula2 = (
            f'=IFERROR(SLOPE(R{start_row}C{y_col}:R{anchor_row}C{y_col},'
            f'R{start_row}C{x_col}:R{anchor_row}C{x_col}),"")'
        )
        workbook.app.calculate()

        intercept = as_number(intercept_helper.value)
        slope = as_number(slope_helper.value)
        forecast_value = as_number(read_value(sheet, forecast_total_cell))
        actual_value = as_number(read_value(sheet, actual_value_cell))
        forecast_max = as_number(read_value(sheet, forecast_max_cell))
        forecast_min = as_number(read_value(sheet, forecast_min_cell))

        if forecast_value is None and intercept is not None and slope is not None:
            current_x = as_number(get_cell_value(sheet, anchor_row, x_col))
            if current_x is not None:
                forecast_value = intercept + slope * current_x

        range_width = (
            (forecast_max - forecast_min)
            if forecast_max is not None and forecast_min is not None
            else None
        )

        signature = rounded_signature([intercept, slope, forecast_value, forecast_max, forecast_min])
        if previous_signature == signature:
            continue
        previous_signature = signature

        row = {
            "model": file_label.model,
            "ticker": file_label.ticker,
            "model_period": file_label.model_period,
            "model_date": file_label.model_date,
            "method": "regression",
            "parameter_name": "num_quarters_used",
            "parameter_value": n_quarters,
            "num_quarters_used": n_quarters,
            "forecast_value": forecast_value,
            "actual_value": actual_value,
            "forecast_max": forecast_max,
            "forecast_min": forecast_min,
            "range_width": range_width,
            "intercept": intercept,
            "slope": slope,
            "source_file": file_label.source_file,
        }

        has_signal = any(
            row[key] not in ("", None)
            for key in ("forecast_value", "forecast_max", "forecast_min", "intercept", "slope")
        )
        if has_signal:
            rows.append(row)

    intercept_helper.clear_contents()
    slope_helper.clear_contents()
    return rows


def write_sheet(
    workbook: Workbook,
    sheet_name: str,
    columns: Sequence[str],
    rows: Sequence[Dict[str, Any]],
) -> None:
    sheet = workbook.create_sheet(title=sheet_name)
    sheet.append(list(columns))

    for row in rows:
        sheet.append([clean_for_output(row.get(column)) for column in columns])

    for cell in sheet[1]:
        cell.font = HEADER_FONT

    sheet.freeze_panes = "A2"
    sheet.auto_filter.ref = sheet.dimensions

    for column_index, column_name in enumerate(columns, start=1):
        max_length = len(column_name)
        for value, in sheet.iter_rows(
            min_row=2,
            min_col=column_index,
            max_col=column_index,
            values_only=True,
        ):
            if value is None:
                continue
            max_length = max(max_length, len(str(value)))
        sheet.column_dimensions[get_column_letter(column_index)].width = min(max_length + 2, 40)


def build_output_workbook(
    output_path: Path,
    empirical_rows: Sequence[Dict[str, Any]],
    regression_rows: Sequence[Dict[str, Any]],
) -> None:
    workbook = Workbook()
    workbook.remove(workbook.active)

    write_sheet(workbook, "empirical_candidates", EMPIRICAL_COLUMNS, empirical_rows)
    write_sheet(workbook, "regression_candidates", REGRESSION_COLUMNS, regression_rows)
    workbook.save(output_path)


def iter_input_files(path: Path) -> Iterable[Path]:
    for file_path in sorted(path.iterdir()):
        if not file_path.is_file():
            print(f"Skipped {file_path.name}: not a file")
            continue
        if file_path.name.startswith("~"):
            print(f"Skipped {file_path.name}: temporary file")
            continue
        if file_path.suffix.lower() != ".xlsx":
            print(f"Skipped {file_path.name}: not an .xlsx file")
            continue
        yield file_path


def main() -> None:
    input_path = Path(input_dir)
    output_path = Path(output_dir)

    if not input_path.exists() or not input_path.is_dir():
        raise FileNotFoundError(f"input_dir does not exist or is not a folder: {input_path}")

    final_output_path = make_output_path(input_path, output_path)
    empirical_rows: List[Dict[str, Any]] = []
    regression_rows: List[Dict[str, Any]] = []
    processed_files = 0

    app = xw.App(visible=False, add_book=False)
    app.display_alerts = False
    app.screen_updating = False

    try:
        for file_path in iter_input_files(input_path):
            print(f"Processing {file_path.name}")
            workbook = None
            try:
                workbook = app.books.open(str(file_path), update_links=False)
                file_label = parse_file_label(file_path.name)

                empirical_rows.extend(extract_empirical_rows(workbook, file_label))
                regression_rows.extend(extract_regression_rows(workbook, file_label))
                processed_files += 1
            except Exception as exc:
                print(f"Skipped {file_path.name}: failed to process ({exc})")
            finally:
                if workbook is not None:
                    safe_close_workbook(workbook)
    finally:
        app.quit()

    build_output_workbook(final_output_path, empirical_rows, regression_rows)

    print(f"Output written to: {final_output_path}")
    print(f"Files processed: {processed_files}")
    print(f"Empirical rows: {len(empirical_rows)}")
    print(f"Regression rows: {len(regression_rows)}")


if __name__ == "__main__":
    main()
