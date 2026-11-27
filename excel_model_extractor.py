#!/usr/bin/env python3
from __future__ import annotations

import re
from dataclasses import dataclass
from pathlib import Path
from typing import Any

import xlwings as xw
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

# ------------------------------- user inputs -------------------------------- #
input_dir = Path("/path/to/input")
output_dir = Path("/path/to/output")
# --------------------------------------------------------------------------- #

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

MONTH_TO_NUMBER = {
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

PERIOD_DAY = {"early": 5, "mid": 15, "late": 25}


@dataclass(frozen=True)
class FileLabels:
    model: str
    ticker: str
    model_period: str
    model_date: str


@dataclass
class SheetSnapshot:
    start_row: int
    start_col: int
    values: list[list[Any]]

    @property
    def row_count(self) -> int:
        return len(self.values)

    @property
    def col_count(self) -> int:
        return len(self.values[0]) if self.values else 0

    @property
    def end_row(self) -> int:
        return self.start_row + self.row_count - 1

    @property
    def end_col(self) -> int:
        return self.start_col + self.col_count - 1

    def get(self, row: int, col: int) -> Any:
        row_idx = row - self.start_row
        col_idx = col - self.start_col
        if row_idx < 0 or col_idx < 0:
            return None
        if row_idx >= self.row_count or col_idx >= self.col_count:
            return None
        return self.values[row_idx][col_idx]


def to_2d(values: Any) -> list[list[Any]]:
    if values is None:
        return [[None]]
    if isinstance(values, (list, tuple)):
        if not values:
            return [[None]]
        first = values[0]
        if isinstance(first, (list, tuple)):
            rows = [list(row) if isinstance(row, (list, tuple)) else [row] for row in values]
        else:
            rows = [list(values)]
    else:
        rows = [[values]]

    width = max(len(row) for row in rows) if rows else 1
    for row in rows:
        if len(row) < width:
            row.extend([None] * (width - len(row)))
    return rows


def normalize_text(value: Any) -> str:
    if value is None:
        return ""
    text = str(value).strip().lower()
    text = text.replace("\n", " ")
    text = re.sub(r"\s+", " ", text)
    return text


def to_float(value: Any) -> float | None:
    if value is None:
        return None
    if isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        return float(value)
    text = str(value).strip()
    if not text:
        return None
    if text.startswith("#"):
        return None
    try:
        return float(text.replace(",", ""))
    except ValueError:
        return None


def to_int(value: Any) -> int | None:
    as_float = to_float(value)
    if as_float is None:
        return None
    return int(as_float)


def subtract_numbers(a: Any, b: Any) -> float | None:
    num_a = to_float(a)
    num_b = to_float(b)
    if num_a is None or num_b is None:
        return None
    return num_a - num_b


def is_blank(value: Any) -> bool:
    if value is None:
        return True
    if isinstance(value, str):
        return not value.strip()
    return False


def find_anchor(snapshot: SheetSnapshot, text: str = "max") -> tuple[int, int] | None:
    target = text.lower()
    for row_idx, row in enumerate(snapshot.values):
        for col_idx, value in enumerate(row):
            if isinstance(value, str) and value.strip().lower() == target:
                return snapshot.start_row + row_idx, snapshot.start_col + col_idx
    return None


def build_snapshot(sheet: xw.Sheet) -> SheetSnapshot:
    used = sheet.used_range
    values = to_2d(used.value)
    return SheetSnapshot(start_row=used.row, start_col=used.column, values=values)


def parse_file_labels(file_path: Path) -> FileLabels:
    stem = file_path.stem
    parts = [part.strip() for part in stem.split(" - ")]

    ticker = parts[1] if len(parts) > 1 else ""
    period_source = parts[2] if len(parts) > 2 else stem
    period_source = period_source.split("_")[0]

    period_match = re.search(
        r"(Early|Mid|Late)(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)(\d{4})",
        period_source,
        flags=re.IGNORECASE,
    )
    if not period_match:
        period_match = re.search(
            r"(Early|Mid|Late)(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)(\d{4})",
            stem,
            flags=re.IGNORECASE,
        )

    if period_match:
        bucket = period_match.group(1).lower()
        month_text = period_match.group(2).lower()
        year_text = period_match.group(3)

        month_num = MONTH_TO_NUMBER[month_text]
        day = PERIOD_DAY[bucket]
        month_title = month_text.capitalize()
        bucket_title = bucket.capitalize()

        model_period = f"{bucket_title}{month_title}_{year_text}"
        model_date = f"{year_text}-{month_num:02d}-{day:02d}"
    else:
        model_period = "UnknownPeriod"
        model_date = ""

    if not ticker:
        ticker_match = re.search(r"-\s*([A-Za-z0-9]+)\s*-", stem)
        ticker = ticker_match.group(1) if ticker_match else "UNKNOWN"

    model = f"{ticker}_{model_period}"
    return FileLabels(model=model, ticker=ticker, model_period=model_period, model_date=model_date)


def next_output_path(input_folder: Path, target_dir: Path) -> Path:
    base_name = f"{input_folder.name}_PARAM"
    candidate = target_dir / f"{base_name}.xlsx"
    idx = 1
    while candidate.exists():
        candidate = target_dir / f"{base_name}.{idx}.xlsx"
        idx += 1
    return candidate


def close_workbook_safely(wb: xw.Book) -> None:
    try:
        wb.close(save=False)
        return
    except TypeError:
        pass
    except Exception:
        pass

    try:
        wb.close(SaveChanges=False)
        return
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
        pass


def header_cells(snapshot: SheetSnapshot, header_row: int) -> list[tuple[int, str]]:
    cells: list[tuple[int, str]] = []
    for col in range(snapshot.start_col, snapshot.end_col + 1):
        cells.append((col, normalize_text(snapshot.get(header_row, col))))
    return cells


def find_column(
    headers: list[tuple[int, str]],
    patterns: list[tuple[str, ...]],
    anchor_col: int,
) -> int | None:
    matches: list[int] = []
    for col, header_text in headers:
        if not header_text:
            continue
        for pattern in patterns:
            if all(token in header_text for token in pattern):
                matches.append(col)
                break
    if not matches:
        return None
    return min(matches, key=lambda col: abs(col - anchor_col))


def nearest_header(headers: list[tuple[int, str]], keyword: str, anchor_col: int) -> int | None:
    matches = [col for col, text in headers if keyword in text]
    if not matches:
        return None
    return min(matches, key=lambda col: abs(col - anchor_col))


def as_list(values: Any, expected_size: int) -> list[Any]:
    if isinstance(values, (list, tuple)):
        if values and isinstance(values[0], (list, tuple)):
            flat = [row[0] if row else None for row in values]
        else:
            flat = list(values)
    else:
        flat = [values]

    if len(flat) < expected_size:
        flat.extend([None] * (expected_size - len(flat)))
    return flat[:expected_size]


def row_has_payload(*values: Any) -> bool:
    return any(not is_blank(value) for value in values)


def extract_empirical_rows(wb: xw.Book, file_path: Path, labels: FileLabels) -> list[dict[str, Any]]:
    rows: list[dict[str, Any]] = []
    source_file = file_path.name

    try:
        sheet = wb.sheets["Empirical Model"]
    except Exception:
        print(f"Skipped empirical extraction for {source_file}: missing sheet 'Empirical Model'")
        return rows

    snapshot = build_snapshot(sheet)
    anchor = find_anchor(snapshot, text="max")
    if not anchor:
        print(f"Skipped empirical extraction for {source_file}: 'max' anchor not found")
        return rows

    anchor_row, anchor_col = anchor
    headers = header_cells(snapshot, anchor_row)
    data_start_row = anchor_row + 1
    n_quarters = 10

    max_col = anchor_col
    min_col = nearest_header(headers, "min", anchor_col) or (anchor_col + 1)

    num_quarters_col = find_column(headers, [("num", "quarter"), ("quarters", "used")], anchor_col) or (anchor_col - 10)
    last_quarter_col = find_column(headers, [("last", "quarter"), ("quarter", "used")], anchor_col) or (anchor_col - 9)
    forecast_col = find_column(
        headers,
        [
            ("estimated", "total", "sold"),
            ("forecast", "value"),
            ("forecast",),
        ],
        anchor_col,
    ) or (anchor_col - 1)
    actual_col = find_column(headers, [("reported", "sales"), ("actual", "value"), ("actual",)], anchor_col) or (anchor_col - 2)
    quarterly_sales_col = find_column(headers, [("quarterly", "sales"), ("quarter", "sales")], anchor_col) or (anchor_col - 6)
    reported_sales_col = find_column(headers, [("reported", "sales")], anchor_col) or actual_col
    growth_rate_col = find_column(headers, [("growth", "rate")], anchor_col) or (anchor_col - 4)
    sales_captured_col = find_column(headers, [("sales", "captured"), ("captured", "db")], anchor_col) or (anchor_col - 3)
    penetration_col = find_column(headers, [("penetration",)], anchor_col)

    scratch_col = snapshot.end_col + 2
    scratch_start_row = data_start_row

    for n in range(1, n_quarters + 1):
        end_row = data_start_row + n - 1
        if end_row > snapshot.end_row:
            formula = '=""'
        elif penetration_col is not None:
            formula = (
                f'=IFERROR(AVERAGE(R{data_start_row}C{penetration_col}:'
                f"R{end_row}C{penetration_col}),\"\")"
            )
        else:
            formula = (
                f'=IFERROR(SUM(R{data_start_row}C{quarterly_sales_col}:R{end_row}C{quarterly_sales_col})/'
                f'SUM(R{data_start_row}C{reported_sales_col}:R{end_row}C{reported_sales_col}),"")'
            )
        sheet.range((scratch_start_row + n - 1, scratch_col)).formula2 = formula

    wb.app.calculate()
    avg_penetration_values = as_list(
        sheet.range((scratch_start_row, scratch_col), (scratch_start_row + n_quarters - 1, scratch_col)).value,
        n_quarters,
    )

    for n in range(1, n_quarters + 1):
        row_num = data_start_row + n - 1

        num_quarters_used = to_int(snapshot.get(row_num, num_quarters_col))
        if num_quarters_used is None:
            num_quarters_used = n

        last_quarter_used = snapshot.get(row_num, last_quarter_col)
        forecast_value = snapshot.get(row_num, forecast_col)
        actual_value = snapshot.get(row_num, actual_col)
        forecast_max = snapshot.get(row_num, max_col)
        forecast_min = snapshot.get(row_num, min_col)
        quarterly_sales = snapshot.get(row_num, quarterly_sales_col)
        reported_sales = snapshot.get(row_num, reported_sales_col)
        growth_rate_pct = snapshot.get(row_num, growth_rate_col)
        sales_captured_pct = snapshot.get(row_num, sales_captured_col)
        avg_penetration_pct = avg_penetration_values[n - 1]

        if not row_has_payload(
            forecast_value,
            actual_value,
            forecast_max,
            forecast_min,
            quarterly_sales,
            reported_sales,
            growth_rate_pct,
            sales_captured_pct,
        ):
            continue

        range_width = subtract_numbers(forecast_max, forecast_min)
        rows.append(
            {
                "model": labels.model,
                "ticker": labels.ticker,
                "model_period": labels.model_period,
                "model_date": labels.model_date,
                "method": "empirical",
                "parameter_name": "avg_penetration_pct",
                "parameter_value": avg_penetration_pct,
                "num_quarters_used": num_quarters_used,
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
                "sales_captured_in_db_pct": sales_captured_pct,
                "source_file": source_file,
            }
        )

    return rows


def dedupe_signature(*values: Any) -> tuple[Any, ...]:
    signature: list[Any] = []
    for value in values:
        number = to_float(value)
        if number is None:
            signature.append(value)
        else:
            signature.append(round(number, 8))
    return tuple(signature)


def extract_regression_rows(wb: xw.Book, file_path: Path, labels: FileLabels) -> list[dict[str, Any]]:
    rows: list[dict[str, Any]] = []
    source_file = file_path.name

    try:
        sheet = wb.sheets["Regression Model"]
    except Exception:
        print(f"Skipped regression extraction for {source_file}: missing sheet 'Regression Model'")
        return rows

    snapshot = build_snapshot(sheet)
    anchor = find_anchor(snapshot, text="max")
    if not anchor:
        print(f"Skipped regression extraction for {source_file}: 'max' anchor not found")
        return rows

    anchor_row, anchor_col = anchor
    headers = header_cells(snapshot, anchor_row)
    data_start_row = anchor_row + 1
    n_quarters = 10

    y_col = anchor_col - 7
    x_col = anchor_col - 11

    max_col = anchor_col
    min_col = nearest_header(headers, "min", anchor_col) or (anchor_col + 1)
    num_quarters_col = find_column(headers, [("num", "quarter"), ("quarters", "used")], anchor_col) or (anchor_col - 10)
    forecast_col = find_column(
        headers,
        [
            ("tot", "fcst", "w/o", "sa"),
            ("tot", "fcst", "wo", "sa"),
            ("forecast", "without", "sa"),
            ("forecast", "value"),
        ],
        anchor_col,
    ) or (anchor_col - 1)
    actual_col = find_column(headers, [("actual", "value"), ("reported", "sales"), ("actual",)], anchor_col)

    scratch_intercept_col = snapshot.end_col + 2
    scratch_slope_col = snapshot.end_col + 3
    scratch_start_row = data_start_row

    for n in range(2, n_quarters + 1):
        end_row = data_start_row + n - 1
        scratch_row = scratch_start_row + n - 2

        if end_row > snapshot.end_row:
            intercept_formula = '=""'
            slope_formula = '=""'
        else:
            intercept_formula = (
                f'=IFERROR(INTERCEPT(R{data_start_row}C{y_col}:R{end_row}C{y_col},'
                f'R{data_start_row}C{x_col}:R{end_row}C{x_col}),"")'
            )
            slope_formula = (
                f'=IFERROR(SLOPE(R{data_start_row}C{y_col}:R{end_row}C{y_col},'
                f'R{data_start_row}C{x_col}:R{end_row}C{x_col}),"")'
            )

        sheet.range((scratch_row, scratch_intercept_col)).formula2 = intercept_formula
        sheet.range((scratch_row, scratch_slope_col)).formula2 = slope_formula

    wb.app.calculate()

    intercept_values = as_list(
        sheet.range(
            (scratch_start_row, scratch_intercept_col),
            (scratch_start_row + n_quarters - 2, scratch_intercept_col),
        ).value,
        n_quarters - 1,
    )
    slope_values = as_list(
        sheet.range(
            (scratch_start_row, scratch_slope_col),
            (scratch_start_row + n_quarters - 2, scratch_slope_col),
        ).value,
        n_quarters - 1,
    )

    previous_signature: tuple[Any, ...] | None = None
    for n in range(2, n_quarters + 1):
        row_num = data_start_row + n - 1
        if row_num > snapshot.end_row:
            break

        num_quarters_used = to_int(snapshot.get(row_num, num_quarters_col))
        if num_quarters_used is None:
            num_quarters_used = n

        forecast_value = snapshot.get(row_num, forecast_col)
        forecast_max = snapshot.get(row_num, max_col)
        forecast_min = snapshot.get(row_num, min_col)
        actual_value = snapshot.get(row_num, actual_col) if actual_col is not None else None
        intercept = intercept_values[n - 2]
        slope = slope_values[n - 2]

        if is_blank(forecast_value):
            x_value = to_float(snapshot.get(row_num, x_col))
            intercept_num = to_float(intercept)
            slope_num = to_float(slope)
            if x_value is not None and intercept_num is not None and slope_num is not None:
                forecast_value = intercept_num + (slope_num * x_value)

        if not row_has_payload(forecast_value, forecast_max, forecast_min, intercept, slope):
            continue

        range_width = subtract_numbers(forecast_max, forecast_min)
        signature = dedupe_signature(
            num_quarters_used,
            forecast_value,
            forecast_max,
            forecast_min,
            intercept,
            slope,
        )
        if signature == previous_signature:
            continue
        previous_signature = signature

        rows.append(
            {
                "model": labels.model,
                "ticker": labels.ticker,
                "model_period": labels.model_period,
                "model_date": labels.model_date,
                "method": "regression",
                "parameter_name": "num_quarters_used",
                "parameter_value": num_quarters_used,
                "num_quarters_used": num_quarters_used,
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

    return rows


def set_column_widths(ws) -> None:
    sample_limit = min(ws.max_row, 200)
    for col_idx in range(1, ws.max_column + 1):
        header_text = str(ws.cell(row=1, column=col_idx).value or "")
        max_len = len(header_text)

        for row_idx in range(2, sample_limit + 1):
            value = ws.cell(row=row_idx, column=col_idx).value
            if value is None:
                continue
            max_len = max(max_len, len(str(value)))

        ws.column_dimensions[get_column_letter(col_idx)].width = min(max(12, max_len + 2), 42)


def write_output(output_path: Path, empirical_rows: list[dict[str, Any]], regression_rows: list[dict[str, Any]]) -> None:
    out_wb = Workbook()
    default_sheet = out_wb.active
    out_wb.remove(default_sheet)

    sheets = [
        ("empirical_candidates", EMPIRICAL_COLUMNS, empirical_rows),
        ("regression_candidates", REGRESSION_COLUMNS, regression_rows),
    ]

    for sheet_name, columns, rows in sheets:
        ws = out_wb.create_sheet(title=sheet_name)
        ws.append(columns)
        for col_idx in range(1, len(columns) + 1):
            ws.cell(row=1, column=col_idx).font = Font(bold=True)

        for row in rows:
            ws.append([row.get(col) for col in columns])

        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions
        set_column_widths(ws)

    out_wb.save(output_path)


def main() -> None:
    if not input_dir.exists():
        raise FileNotFoundError(f"Input folder does not exist: {input_dir}")

    output_dir.mkdir(parents=True, exist_ok=True)
    output_path = next_output_path(input_dir, output_dir)

    empirical_rows: list[dict[str, Any]] = []
    regression_rows: list[dict[str, Any]] = []
    processed_files = 0

    app = xw.App(visible=False, add_book=False)
    app.display_alerts = False
    app.screen_updating = False
    app.calculation = "manual"

    try:
        for file_path in sorted(input_dir.iterdir(), key=lambda p: p.name.lower()):
            if not file_path.is_file():
                continue
            if file_path.name.startswith("~"):
                print(f"Skipped {file_path.name}: temporary file")
                continue
            if file_path.suffix.lower() != ".xlsx":
                print(f"Skipped {file_path.name}: not an .xlsx file")
                continue

            print(f"Processing {file_path.name}")
            wb = None
            try:
                labels = parse_file_labels(file_path)
                wb = app.books.open(str(file_path), update_links=False)
                empirical_rows.extend(extract_empirical_rows(wb, file_path, labels))
                regression_rows.extend(extract_regression_rows(wb, file_path, labels))
                processed_files += 1
            except Exception as exc:
                print(f"Skipped {file_path.name}: processing error ({exc})")
            finally:
                if wb is not None:
                    close_workbook_safely(wb)
    finally:
        app.quit()

    write_output(output_path, empirical_rows, regression_rows)

    print(f"Output written: {output_path}")
    print(f"Files processed: {processed_files}")
    print(f"Empirical rows: {len(empirical_rows)}")
    print(f"Regression rows: {len(regression_rows)}")


if __name__ == "__main__":
    main()
