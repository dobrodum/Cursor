#!/usr/bin/env python3
from __future__ import annotations

import calendar
import re
from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path
from typing import Any

import xlwings as xw
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

# Update these paths before running.
input_dir = Path("input")
output_dir = Path("output")

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

PERIOD_RE = re.compile(
    r"(?P<phase>Early|Mid|Late)(?P<month>[A-Za-z]{3,4})(?P<year>\d{4})",
    re.IGNORECASE,
)

DAY_BY_PHASE = {"early": 5, "mid": 15, "late": 25}
MONTH_BY_NAME = {name.lower(): idx for idx, name in enumerate(calendar.month_abbr) if name}
MONTH_BY_NAME["sept"] = 9


@dataclass(frozen=True)
class FileMetadata:
    model: str
    ticker: str
    model_period: str
    model_date: str


def normalize_2d(values: Any) -> list[list[Any]]:
    if values is None:
        return []
    if not isinstance(values, list):
        return [[values]]
    if values and not isinstance(values[0], list):
        return [values]
    return values


def normalize_label(value: Any) -> str:
    if not isinstance(value, str):
        return ""
    label = re.sub(r"\s+", " ", value.strip().lower())
    return label


def to_float(value: Any) -> float | None:
    if value is None or isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        return float(value)
    if isinstance(value, str):
        text = value.strip().replace(",", "")
        if not text:
            return None
        percent = text.endswith("%")
        if percent:
            text = text[:-1]
        try:
            parsed = float(text)
        except ValueError:
            return None
        if percent and abs(parsed) > 1:
            parsed /= 100.0
        return parsed
    return None


def quarter_text(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, (datetime, date)):
        return value.isoformat()
    return str(value).strip()


def parse_file_metadata(file_path: Path) -> FileMetadata:
    stem = file_path.stem
    parts = [part.strip() for part in stem.split(" - ") if part.strip()]
    ticker = ""
    if len(parts) >= 2:
        ticker = parts[1]
    elif parts:
        ticker = parts[0]
    ticker = re.sub(r"\s+", "", ticker).upper()

    period_source = parts[2] if len(parts) >= 3 else stem
    period_source = re.sub(r"[_-]?send$", "", period_source, flags=re.IGNORECASE)

    match = PERIOD_RE.search(period_source) or PERIOD_RE.search(stem)
    if not match:
        raise ValueError("filename does not contain Early/Mid/Late period token")

    phase = match.group("phase").title()
    month_token = match.group("month").lower()
    if month_token not in MONTH_BY_NAME and month_token[:3] in MONTH_BY_NAME:
        month_token = month_token[:3]
    month_number = MONTH_BY_NAME.get(month_token)
    if month_number is None:
        raise ValueError(f"unsupported month token: {match.group('month')}")

    year = int(match.group("year"))
    day = DAY_BY_PHASE[phase.lower()]
    model_period = f"{phase}{calendar.month_abbr[month_number]}_{year}"
    model_date = date(year, month_number, day).isoformat()
    model = f"{ticker}_{model_period}"

    return FileMetadata(
        model=model,
        ticker=ticker,
        model_period=model_period,
        model_date=model_date,
    )


def next_output_path(input_folder: Path, destination_dir: Path) -> Path:
    destination_dir.mkdir(parents=True, exist_ok=True)
    base_name = f"{input_folder.name}_PARAM"
    candidate = destination_dir / f"{base_name}.xlsx"
    suffix = 1
    while candidate.exists():
        candidate = destination_dir / f"{base_name}.{suffix}.xlsx"
        suffix += 1
    return candidate


def close_workbook_safely(workbook: xw.Book) -> None:
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
        return
    except Exception:
        pass

    try:
        workbook.api.Close(False)
    except Exception:
        pass


def build_label_index(sheet: xw.Sheet) -> tuple[dict[str, list[tuple[int, int]]], int]:
    used = sheet.used_range
    values = normalize_2d(used.value)
    label_index: dict[str, list[tuple[int, int]]] = {}
    for row_offset, row_values in enumerate(values):
        if not isinstance(row_values, list):
            row_values = [row_values]
        for col_offset, cell_value in enumerate(row_values):
            label = normalize_label(cell_value)
            if label:
                row = used.row + row_offset
                col = used.column + col_offset
                label_index.setdefault(label, []).append((row, col))
    helper_col = used.last_cell.column + 3
    return label_index, helper_col


def find_anchor_max(label_index: dict[str, list[tuple[int, int]]]) -> tuple[int, int] | None:
    if "max" in label_index and label_index["max"]:
        return sorted(label_index["max"])[0]
    candidates: list[tuple[int, int]] = []
    for label, cells in label_index.items():
        if re.match(r"^max(\b|$)", label):
            candidates.extend(cells)
    return sorted(candidates)[0] if candidates else None


def read_neighbor(sheet: xw.Sheet, row: int, col: int, offsets: list[tuple[int, int]]) -> Any:
    for row_off, col_off in offsets:
        try:
            value = sheet.cells(row + row_off, col + col_off).value
        except Exception:
            value = None
        if value not in (None, ""):
            return value
    return None


def get_value_by_label(
    sheet: xw.Sheet,
    label_index: dict[str, list[tuple[int, int]]],
    labels: list[str],
) -> Any:
    for label in labels:
        normalized = normalize_label(label)
        for row, col in label_index.get(normalized, []):
            value = read_neighbor(sheet, row, col, [(0, 1), (1, 0), (0, -1), (-1, 0)])
            if value not in (None, ""):
                return value
    return None


def extract_history_points(
    sheet: xw.Sheet,
    anchor_row: int,
    anchor_col: int,
    n_quarters: int,
) -> tuple[list[dict[str, Any]], int, int]:
    y_col = anchor_col - 7
    x_col = anchor_col - 11
    quarter_col = x_col - 1
    if x_col < 1 or y_col < 1 or quarter_col < 1:
        return [], x_col, y_col

    start_row = max(1, anchor_row - 120)
    end_row = anchor_row - 1
    points: list[dict[str, Any]] = []
    for row in range(start_row, end_row + 1):
        x_value = to_float(sheet.cells(row, x_col).value)
        y_value = to_float(sheet.cells(row, y_col).value)
        if x_value is None or y_value is None:
            continue
        quarter_value = sheet.cells(row, quarter_col).value
        points.append(
            {
                "row": row,
                "x": x_value,
                "y": y_value,
                "quarter": quarter_text(quarter_value),
            }
        )

    if len(points) > n_quarters:
        points = points[-n_quarters:]
    return points, x_col, y_col


def numeric_signature(*values: Any) -> tuple[Any, ...]:
    signature: list[Any] = []
    for value in values:
        if isinstance(value, float):
            signature.append(round(value, 12))
        else:
            signature.append(value)
    return tuple(signature)


def process_empirical_sheet(
    workbook: xw.Book,
    metadata: FileMetadata,
    source_file: str,
    n_quarters: int = 10,
) -> list[dict[str, Any]]:
    try:
        sheet = workbook.sheets["Empirical Model"]
    except Exception:
        return []

    label_index, helper_col = build_label_index(sheet)
    anchor = find_anchor_max(label_index)
    if anchor is None:
        return []
    anchor_row, anchor_col = anchor

    points, x_col, _ = extract_history_points(sheet, anchor_row, anchor_col, n_quarters=n_quarters)
    if not points:
        return []

    helper_row = anchor_row + 2
    max_value = to_float(read_neighbor(sheet, anchor_row, anchor_col, [(0, 1), (0, 2)]))
    min_value = to_float(read_neighbor(sheet, anchor_row + 1, anchor_col, [(0, 1), (0, 2)]))
    if min_value is None:
        min_loc = label_index.get("min", [])
        if min_loc:
            min_value = to_float(read_neighbor(sheet, min_loc[0][0], min_loc[0][1], [(0, 1), (0, 2)]))

    reported_sales = to_float(
        get_value_by_label(sheet, label_index, ["reported sales", "actual sales", "actual value"])
    )
    growth_rate_pct = to_float(
        get_value_by_label(sheet, label_index, ["growth rate", "growth rate %", "growth %"])
    )
    sales_captured_pct = to_float(
        get_value_by_label(
            sheet,
            label_index,
            ["sales captured in db %", "sales captured in db", "captured in db %"],
        )
    )
    estimated_total_sold = to_float(
        get_value_by_label(
            sheet,
            label_index,
            ["estimated total sold", "est total sold", "total sold estimate"],
        )
    )

    rows: list[dict[str, Any]] = []
    max_loop = min(n_quarters, len(points))
    helper_cell = sheet.cells(helper_row, helper_col)

    for quarters_used in range(1, max_loop + 1):
        selected = points[-quarters_used:]
        first_row = selected[0]["row"]
        last_row = selected[-1]["row"]

        helper_cell.formula2 = f"=AVERAGE(R{first_row}C{x_col}:R{last_row}C{x_col})"
        workbook.app.calculate()

        avg_penetration_pct = to_float(helper_cell.value)
        quarterly_sales = selected[-1]["y"]
        forecast_value = estimated_total_sold
        if forecast_value is None and avg_penetration_pct not in (None, 0):
            forecast_value = quarterly_sales / avg_penetration_pct

        forecast_max = max_value if max_value is not None else forecast_value
        forecast_min = min_value if min_value is not None else forecast_value
        range_width = (
            forecast_max - forecast_min
            if forecast_max is not None and forecast_min is not None
            else None
        )

        rows.append(
            {
                "model": metadata.model,
                "ticker": metadata.ticker,
                "model_period": metadata.model_period,
                "model_date": metadata.model_date,
                "method": "empirical",
                "parameter_name": "avg_penetration_pct",
                "parameter_value": avg_penetration_pct,
                "num_quarters_used": quarters_used,
                "last_quarter_used": selected[-1]["quarter"],
                "forecast_value": forecast_value,
                "actual_value": reported_sales,
                "forecast_max": forecast_max,
                "forecast_min": forecast_min,
                "range_width": range_width,
                "avg_penetration_pct": avg_penetration_pct,
                "quarterly_sales": quarterly_sales,
                "reported_sales": reported_sales,
                "growth_rate_pct": growth_rate_pct,
                "sales_captured_in_db_pct": (
                    sales_captured_pct if sales_captured_pct is not None else selected[-1]["x"]
                ),
                "source_file": source_file,
            }
        )

    return rows


def process_regression_sheet(
    workbook: xw.Book,
    metadata: FileMetadata,
    source_file: str,
    n_quarters: int = 10,
) -> list[dict[str, Any]]:
    try:
        sheet = workbook.sheets["Regression Model"]
    except Exception:
        return []

    label_index, helper_col = build_label_index(sheet)
    anchor = find_anchor_max(label_index)
    if anchor is None:
        return []
    anchor_row, anchor_col = anchor

    points, x_col, y_col = extract_history_points(sheet, anchor_row, anchor_col, n_quarters=n_quarters)
    if not points:
        return []

    max_value = to_float(read_neighbor(sheet, anchor_row, anchor_col, [(0, 1), (0, 2)]))
    min_value = to_float(read_neighbor(sheet, anchor_row + 1, anchor_col, [(0, 1), (0, 2)]))
    if min_value is None:
        min_loc = label_index.get("min", [])
        if min_loc:
            min_value = to_float(read_neighbor(sheet, min_loc[0][0], min_loc[0][1], [(0, 1), (0, 2)]))

    forecast_total_without_sa = to_float(
        get_value_by_label(
            sheet,
            label_index,
            ["tot fcst w/o sa", "total forecast w/o sa", "tot fcst without sa"],
        )
    )
    actual_value = to_float(get_value_by_label(sheet, label_index, ["actual value", "reported sales"]))

    intercept_cell = sheet.cells(anchor_row + 2, helper_col)
    slope_cell = sheet.cells(anchor_row + 3, helper_col)

    rows: list[dict[str, Any]] = []
    previous_signature: tuple[Any, ...] | None = None
    max_loop = min(n_quarters, len(points))

    for quarters_used in range(1, max_loop + 1):
        selected = points[-quarters_used:]
        first_row = selected[0]["row"]
        last_row = selected[-1]["row"]

        intercept_cell.formula2 = (
            f"=INTERCEPT(R{first_row}C{y_col}:R{last_row}C{y_col},"
            f"R{first_row}C{x_col}:R{last_row}C{x_col})"
        )
        slope_cell.formula2 = (
            f"=SLOPE(R{first_row}C{y_col}:R{last_row}C{y_col},"
            f"R{first_row}C{x_col}:R{last_row}C{x_col})"
        )
        workbook.app.calculate()

        intercept = to_float(intercept_cell.value)
        slope = to_float(slope_cell.value)
        forecast_value = forecast_total_without_sa
        if forecast_value is None and intercept is not None and slope is not None:
            forecast_value = intercept + (slope * selected[-1]["x"])

        forecast_max = max_value if max_value is not None else forecast_value
        forecast_min = min_value if min_value is not None else forecast_value
        range_width = (
            forecast_max - forecast_min
            if forecast_max is not None and forecast_min is not None
            else None
        )

        signature = numeric_signature(
            forecast_value,
            intercept,
            slope,
            forecast_max,
            forecast_min,
        )
        if signature == previous_signature:
            continue
        previous_signature = signature

        rows.append(
            {
                "model": metadata.model,
                "ticker": metadata.ticker,
                "model_period": metadata.model_period,
                "model_date": metadata.model_date,
                "method": "regression",
                "parameter_name": "num_quarters_used",
                "parameter_value": quarters_used,
                "num_quarters_used": quarters_used,
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


def autosize_columns(worksheet) -> None:
    for col_idx, col_cells in enumerate(worksheet.iter_cols(), start=1):
        max_len = 0
        for cell in col_cells:
            value = "" if cell.value is None else str(cell.value)
            if len(value) > max_len:
                max_len = len(value)
        worksheet.column_dimensions[get_column_letter(col_idx)].width = min(max(12, max_len + 2), 42)


def write_sheet(worksheet, headers: list[str], rows: list[dict[str, Any]]) -> None:
    worksheet.append(headers)
    for row in rows:
        worksheet.append([row.get(header) for header in headers])

    for header_cell in worksheet[1]:
        header_cell.font = Font(bold=True)
    worksheet.freeze_panes = "A2"
    worksheet.auto_filter.ref = worksheet.dimensions
    autosize_columns(worksheet)


def write_output_workbook(
    output_path: Path,
    empirical_rows: list[dict[str, Any]],
    regression_rows: list[dict[str, Any]],
) -> None:
    output_book = Workbook()
    empirical_sheet = output_book.active
    empirical_sheet.title = "empirical_candidates"
    regression_sheet = output_book.create_sheet("regression_candidates")

    write_sheet(empirical_sheet, EMPIRICAL_HEADERS, empirical_rows)
    write_sheet(regression_sheet, REGRESSION_HEADERS, regression_rows)
    output_book.save(output_path)


def main() -> None:
    if not input_dir.exists():
        raise FileNotFoundError(f"input_dir does not exist: {input_dir}")

    candidate_files = sorted(input_dir.iterdir(), key=lambda p: p.name.lower())
    empirical_rows: list[dict[str, Any]] = []
    regression_rows: list[dict[str, Any]] = []
    files_processed = 0

    with xw.App(visible=False, add_book=False) as app:
        app.display_alerts = False
        app.screen_updating = False
        app.calculation = "manual"

        for file_path in candidate_files:
            if not file_path.is_file():
                continue
            if file_path.name.startswith("~"):
                print(f"Skipped file: {file_path.name} (temporary file)")
                continue
            if file_path.suffix.lower() != ".xlsx":
                print(f"Skipped file: {file_path.name} (not .xlsx)")
                continue

            try:
                metadata = parse_file_metadata(file_path)
            except Exception as exc:
                print(f"Skipped file: {file_path.name} (metadata parse error: {exc})")
                continue

            workbook: xw.Book | None = None
            try:
                workbook = app.books.open(str(file_path), update_links=False)
                empirical_rows.extend(
                    process_empirical_sheet(
                        workbook=workbook,
                        metadata=metadata,
                        source_file=file_path.name,
                    )
                )
                regression_rows.extend(
                    process_regression_sheet(
                        workbook=workbook,
                        metadata=metadata,
                        source_file=file_path.name,
                    )
                )
                files_processed += 1
                print(f"Processed file: {file_path.name}")
            except Exception as exc:
                print(f"Skipped file: {file_path.name} (processing error: {exc})")
            finally:
                if workbook is not None:
                    close_workbook_safely(workbook)

    output_path = next_output_path(input_folder=input_dir, destination_dir=output_dir)
    write_output_workbook(output_path, empirical_rows, regression_rows)

    print(f"Output path: {output_path}")
    print(f"Files processed: {files_processed}")
    print(f"Empirical rows: {len(empirical_rows)}")
    print(f"Regression rows: {len(regression_rows)}")


if __name__ == "__main__":
    main()
