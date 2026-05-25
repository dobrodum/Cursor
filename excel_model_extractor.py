#!/usr/bin/env python3
"""Extract empirical and regression candidates from model workbooks."""

from __future__ import annotations

from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path
import re
from typing import Any, Iterable, Optional

import xlwings as xw
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter


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

PERIOD_TO_DAY = {
    "early": 5,
    "mid": 15,
    "late": 25,
}

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


def normalize_text(value: Any) -> str:
    if value is None:
        return ""
    text = str(value).strip().lower()
    return re.sub(r"\s+", " ", text)


def is_blank(value: Any) -> bool:
    return value is None or (isinstance(value, str) and not value.strip())


def is_number(value: Any) -> bool:
    try:
        if value is None or value == "":
            return False
        float(value)
        return True
    except (TypeError, ValueError):
        return False


def maybe_float(value: Any) -> Optional[float]:
    if not is_number(value):
        return None
    return float(value)


def rounded_signature(*values: Any) -> tuple[Any, ...]:
    signature: list[Any] = []
    for item in values:
        number = maybe_float(item)
        if number is None:
            signature.append(item)
        else:
            signature.append(round(number, 10))
    return tuple(signature)


def normalize_2d(values: Any) -> list[list[Any]]:
    if isinstance(values, list):
        if not values:
            return []
        if isinstance(values[0], list):
            return values
        return [values]
    return [[values]]


def close_source_workbook(wb: xw.Book) -> None:
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
        return
    except Exception:
        pass

    try:
        wb.close()
    except Exception:
        pass


def set_formula2(rng: xw.Range, formula_r1c1: str) -> None:
    try:
        rng.formula2 = formula_r1c1
    except Exception:
        rng.formula = formula_r1c1


def get_output_path(src_dir: Path, dst_dir: Path) -> Path:
    base_name = f"{src_dir.name}_PARAM"
    target = dst_dir / f"{base_name}.xlsx"
    if not target.exists():
        return target

    counter = 1
    while True:
        candidate = dst_dir / f"{base_name}.{counter}.xlsx"
        if not candidate.exists():
            return candidate
        counter += 1


def parse_file_metadata(file_name: str) -> dict[str, str]:
    stem = Path(file_name).stem

    ticker_match = re.search(
        r"-\s*([A-Za-z0-9]+)\s*-\s*(?:Early|Mid|Late)",
        stem,
        flags=re.IGNORECASE,
    )
    if ticker_match:
        ticker = ticker_match.group(1).upper()
    else:
        parts = [segment.strip() for segment in stem.split("-") if segment.strip()]
        ticker = parts[-2].upper() if len(parts) >= 2 else "UNKNOWN"

    period_match = re.search(
        r"(Early|Mid|Late)\s*([A-Za-z]{3,9})\s*(\d{4})",
        stem,
        flags=re.IGNORECASE,
    )
    if period_match:
        period_name = period_match.group(1).lower()
        month_raw = period_match.group(2)[:3].lower()
        year = int(period_match.group(3))
        day = PERIOD_TO_DAY.get(period_name, 15)
        month_number = MONTH_TO_NUMBER.get(month_raw, 1)
        month_token = datetime(year, month_number, day).strftime("%b")
        model_period = f"{period_name.title()}{month_token}_{year}"
        model_date = date(year, month_number, day).isoformat()
    else:
        model_period = "Unknown_0000"
        model_date = ""

    return {
        "model": f"{ticker}_{model_period}",
        "ticker": ticker,
        "model_period": model_period,
        "model_date": model_date,
    }


@dataclass
class SheetSnapshot:
    sheet: xw.Sheet
    values: list[list[Any]]
    first_row: int
    first_col: int

    @classmethod
    def from_sheet(cls, sheet: xw.Sheet) -> "SheetSnapshot":
        used = sheet.used_range
        values = normalize_2d(used.value)
        return cls(sheet=sheet, values=values, first_row=used.row, first_col=used.column)

    @property
    def row_count(self) -> int:
        return len(self.values)

    @property
    def col_count(self) -> int:
        return max((len(row) for row in self.values), default=0)

    @property
    def last_row(self) -> int:
        return self.first_row + self.row_count - 1

    @property
    def last_col(self) -> int:
        return self.first_col + self.col_count - 1

    def get(self, row: int, col: int) -> Any:
        r_idx = row - self.first_row
        c_idx = col - self.first_col
        if r_idx < 0 or c_idx < 0 or r_idx >= self.row_count:
            return None
        row_values = self.values[r_idx]
        if c_idx >= len(row_values):
            return None
        return row_values[c_idx]

    def get_row_values(self, row: int) -> list[Any]:
        r_idx = row - self.first_row
        if r_idx < 0 or r_idx >= self.row_count:
            return []
        return self.values[r_idx]

    def find_max_anchor(self) -> Optional[tuple[int, int]]:
        candidates: list[tuple[int, int, int]] = []
        for r_idx, row_values in enumerate(self.values):
            normalized_row = [normalize_text(item) for item in row_values]
            for c_idx, text in enumerate(normalized_row):
                if text != "max":
                    continue
                score = 0
                if "min" in normalized_row[c_idx : c_idx + 4]:
                    score += 2
                for check_row in range(r_idx + 1, min(r_idx + 8, self.row_count)):
                    if c_idx < len(self.values[check_row]) and is_number(self.values[check_row][c_idx]):
                        score += 1
                        break
                candidates.append((score, self.first_row + r_idx, self.first_col + c_idx))

        if not candidates:
            return None
        candidates.sort(key=lambda item: (item[0], item[1], item[2]), reverse=True)
        _, row, col = candidates[0]
        return row, col

    def build_header_map(self, header_row: int) -> dict[str, int]:
        header_map: dict[str, int] = {}
        for c_idx, value in enumerate(self.get_row_values(header_row)):
            text = normalize_text(value)
            if text and text not in header_map:
                header_map[text] = self.first_col + c_idx
        return header_map

    def find_data_rows(self, header_row: int, candidate_cols: Iterable[Optional[int]]) -> list[int]:
        columns = [column for column in candidate_cols if column is not None]
        if not columns:
            return []

        rows: list[int] = []
        started = False
        blank_streak = 0
        for row in range(header_row + 1, self.last_row + 1):
            values = [self.get(row, col) for col in columns]
            has_value = any(not is_blank(value) for value in values)
            has_numeric = any(is_number(value) for value in values)

            if has_numeric or (started and has_value):
                rows.append(row)
                started = True
                blank_streak = 0
            elif started:
                blank_streak += 1
                if blank_streak >= 12:
                    break
        return rows


def find_column(header_map: dict[str, int], aliases: list[str]) -> Optional[int]:
    normalized_aliases = [normalize_text(alias) for alias in aliases]
    for alias in normalized_aliases:
        for header_name, column in header_map.items():
            if alias in header_name:
                return column
    return None


def safe_subtract(left: Any, right: Any) -> Optional[float]:
    left_num = maybe_float(left)
    right_num = maybe_float(right)
    if left_num is None or right_num is None:
        return None
    return left_num - right_num


def process_empirical(
    wb: xw.Book,
    metadata: dict[str, str],
    source_file: str,
) -> list[list[Any]]:
    try:
        sheet = wb.sheets["Empirical Model"]
    except Exception:
        return []

    snapshot = SheetSnapshot.from_sheet(sheet)
    anchor = snapshot.find_max_anchor()
    if anchor is None:
        return []

    anchor_row, anchor_col = anchor
    header_map = snapshot.build_header_map(anchor_row)

    max_col = anchor_col
    min_col = find_column(header_map, ["min"]) or (anchor_col + 1)
    forecast_col = find_column(
        header_map,
        ["estimated total sold", "tot fcst w/o sa", "tot fcst", "forecast"],
    ) or (anchor_col - 1)
    reported_col = find_column(
        header_map,
        ["reported sales", "reported"],
    ) or (anchor_col - 2)
    quarter_col = find_column(header_map, ["quarter", "qtr", "period"]) or (anchor_col - 11)
    quarterly_sales_col = find_column(
        header_map,
        ["quarterly sales", "quarter sales", "sales"],
    ) or (anchor_col - 8)
    growth_col = find_column(header_map, ["growth rate", "growth"]) or (anchor_col - 6)
    captured_col = find_column(
        header_map,
        ["sales captured in db", "captured in db", "penetration"],
    ) or (anchor_col - 5)

    data_rows = snapshot.find_data_rows(
        anchor_row,
        [max_col, min_col, forecast_col, reported_col, quarterly_sales_col, captured_col],
    )
    if not data_rows:
        return []

    iterations = min(10, len(data_rows))
    scratch_row = anchor_row
    scratch_col = snapshot.last_col + 2
    scratch_cell = sheet.range((scratch_row, scratch_col))

    rows: list[list[Any]] = []
    for num_quarters_used in range(1, iterations + 1):
        selected_rows = data_rows[-num_quarters_used:]
        start_row = selected_rows[0]
        end_row = selected_rows[-1]

        avg_penetration_pct = None
        if captured_col is not None:
            avg_formula = (
                f"=AVERAGE(R{start_row}C{captured_col}:R{end_row}C{captured_col})"
            )
            set_formula2(scratch_cell, avg_formula)
            wb.app.calculate()
            avg_penetration_pct = scratch_cell.value

        forecast_max = snapshot.get(end_row, max_col)
        forecast_min = snapshot.get(end_row, min_col)
        forecast_value = snapshot.get(end_row, forecast_col)
        actual_value = snapshot.get(end_row, reported_col)
        last_quarter_used = snapshot.get(start_row, quarter_col)
        quarterly_sales = snapshot.get(end_row, quarterly_sales_col)
        growth_rate_pct = snapshot.get(end_row, growth_col)
        sales_captured_pct = snapshot.get(end_row, captured_col)

        rows.append(
            [
                metadata["model"],
                metadata["ticker"],
                metadata["model_period"],
                metadata["model_date"],
                "empirical",
                "avg_penetration_pct",
                avg_penetration_pct,
                num_quarters_used,
                last_quarter_used,
                forecast_value,
                actual_value,
                forecast_max,
                forecast_min,
                safe_subtract(forecast_max, forecast_min),
                avg_penetration_pct,
                quarterly_sales,
                actual_value,
                growth_rate_pct,
                sales_captured_pct,
                source_file,
            ]
        )

    scratch_cell.value = None
    return rows


def process_regression(
    wb: xw.Book,
    metadata: dict[str, str],
    source_file: str,
) -> list[list[Any]]:
    try:
        sheet = wb.sheets["Regression Model"]
    except Exception:
        return []

    snapshot = SheetSnapshot.from_sheet(sheet)
    anchor = snapshot.find_max_anchor()
    if anchor is None:
        return []

    anchor_row, anchor_col = anchor
    header_map = snapshot.build_header_map(anchor_row)

    max_col = anchor_col
    min_col = find_column(header_map, ["min"]) or (anchor_col + 1)
    forecast_col = find_column(
        header_map,
        ["tot fcst w/o sa", "tot fcst without sa", "forecast total"],
    ) or (anchor_col - 1)
    actual_col = find_column(header_map, ["actual", "reported sales"])
    num_quarters_col = find_column(header_map, ["num quarters", "quarters used", "n_quarters"])

    y_col = anchor_col - 7
    x_col = anchor_col - 11

    data_rows = snapshot.find_data_rows(anchor_row, [y_col, x_col, max_col, min_col, forecast_col])
    if not data_rows:
        return []

    iterations = min(10, len(data_rows))
    scratch_row = anchor_row
    intercept_col = snapshot.last_col + 2
    slope_col = snapshot.last_col + 3
    intercept_cell = sheet.range((scratch_row, intercept_col))
    slope_cell = sheet.range((scratch_row, slope_col))

    rows: list[list[Any]] = []
    previous_signature: Optional[tuple[Any, ...]] = None

    for iteration in range(1, iterations + 1):
        selected_rows = data_rows[-iteration:]
        start_row = selected_rows[0]
        end_row = selected_rows[-1]

        intercept_formula = (
            f"=INTERCEPT(R{start_row}C{y_col}:R{end_row}C{y_col},"
            f"R{start_row}C{x_col}:R{end_row}C{x_col})"
        )
        slope_formula = (
            f"=SLOPE(R{start_row}C{y_col}:R{end_row}C{y_col},"
            f"R{start_row}C{x_col}:R{end_row}C{x_col})"
        )
        set_formula2(intercept_cell, intercept_formula)
        set_formula2(slope_cell, slope_formula)
        wb.app.calculate()

        intercept = intercept_cell.value
        slope = slope_cell.value
        num_quarters_used = snapshot.get(end_row, num_quarters_col) if num_quarters_col else iteration
        if not is_number(num_quarters_used):
            num_quarters_used = iteration

        forecast_value = snapshot.get(end_row, forecast_col)
        if is_blank(forecast_value) and is_number(intercept) and is_number(slope):
            x_value = snapshot.get(end_row, x_col)
            if is_number(x_value):
                forecast_value = maybe_float(intercept) + maybe_float(slope) * maybe_float(x_value)

        forecast_max = snapshot.get(end_row, max_col)
        forecast_min = snapshot.get(end_row, min_col)
        actual_value = snapshot.get(end_row, actual_col) if actual_col is not None else ""

        row_signature = rounded_signature(
            num_quarters_used,
            intercept,
            slope,
            forecast_value,
            forecast_max,
            forecast_min,
        )
        if row_signature == previous_signature:
            continue
        previous_signature = row_signature

        rows.append(
            [
                metadata["model"],
                metadata["ticker"],
                metadata["model_period"],
                metadata["model_date"],
                "regression",
                "num_quarters_used",
                num_quarters_used,
                num_quarters_used,
                forecast_value,
                actual_value,
                forecast_max,
                forecast_min,
                safe_subtract(forecast_max, forecast_min),
                intercept,
                slope,
                source_file,
            ]
        )

    intercept_cell.value = None
    slope_cell.value = None
    return rows


def write_sheet(
    workbook: Workbook,
    title: str,
    columns: list[str],
    rows: list[list[Any]],
) -> None:
    ws = workbook.create_sheet(title=title)
    ws.append(columns)
    for row in rows:
        ws.append(row)

    for header_cell in ws[1]:
        header_cell.font = Font(bold=True)

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions

    for col_idx, column_name in enumerate(columns, start=1):
        width = len(column_name)
        for row_idx in range(2, ws.max_row + 1):
            value = ws.cell(row=row_idx, column=col_idx).value
            if value is None:
                continue
            width = max(width, len(str(value)))
        ws.column_dimensions[get_column_letter(col_idx)].width = min(60, width + 2)


def main() -> None:
    src_dir = Path(input_dir)
    dst_dir = Path(output_dir)
    dst_dir.mkdir(parents=True, exist_ok=True)

    if not src_dir.exists():
        print(f"Input directory does not exist: {src_dir.resolve()}")
        return

    output_path = get_output_path(src_dir.resolve(), dst_dir.resolve())

    empirical_rows: list[list[Any]] = []
    regression_rows: list[list[Any]] = []
    files_processed = 0

    app = xw.App(visible=False, add_book=False)
    app.display_alerts = False
    app.screen_updating = False

    try:
        for file_path in sorted(src_dir.iterdir()):
            if not file_path.is_file():
                continue

            if file_path.name.startswith("~"):
                print(f"Skipped file: {file_path.name} (temp file)")
                continue

            if file_path.suffix.lower() != ".xlsx":
                print(f"Skipped file: {file_path.name} (not .xlsx)")
                continue

            wb: Optional[xw.Book] = None
            try:
                wb = app.books.open(str(file_path), update_links=False)
                metadata = parse_file_metadata(file_path.name)

                empirical_rows.extend(process_empirical(wb, metadata, file_path.name))
                regression_rows.extend(process_regression(wb, metadata, file_path.name))
                files_processed += 1
                print(f"Processed file: {file_path.name}")
            except Exception as exc:
                print(f"Skipped file: {file_path.name} (error: {exc})")
            finally:
                if wb is not None:
                    close_source_workbook(wb)
    finally:
        app.quit()

    output_book = Workbook()
    default_sheet = output_book.active
    output_book.remove(default_sheet)

    write_sheet(output_book, "empirical_candidates", EMPIRICAL_COLUMNS, empirical_rows)
    write_sheet(output_book, "regression_candidates", REGRESSION_COLUMNS, regression_rows)
    output_book.save(output_path)

    print(f"Output path: {output_path}")
    print(f"Number of files processed: {files_processed}")
    print(f"Number of empirical rows: {len(empirical_rows)}")
    print(f"Number of regression rows: {len(regression_rows)}")


if __name__ == "__main__":
    main()
