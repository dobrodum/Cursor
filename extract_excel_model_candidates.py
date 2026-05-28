#!/usr/bin/env python3
from __future__ import annotations

import re
from dataclasses import dataclass
from datetime import date
from pathlib import Path
from typing import Any, Iterable

import xlwings as xw
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

# Configure these two paths before running.
input_dir = Path("input")
output_dir = Path("output")

N_QUARTERS = 10

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

MONTH_MAP = {
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

PERIOD_DAY_MAP = {"early": 5, "mid": 15, "late": 25}


@dataclass(frozen=True)
class FileLabels:
    model: str
    ticker: str
    model_period: str
    model_date: str


def normalize_header(value: Any) -> str:
    text = str(value or "").strip().lower()
    return re.sub(r"[^a-z0-9]+", " ", text).strip()


def as_matrix(values: Any) -> list[list[Any]]:
    if values is None:
        return []
    if not isinstance(values, list):
        return [[values]]
    if values and not isinstance(values[0], list):
        return [values]
    return values


def is_blank(value: Any) -> bool:
    return value is None or (isinstance(value, str) and not value.strip())


def as_float(value: Any) -> float | None:
    if value is None:
        return None
    if isinstance(value, bool):
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


def as_int(value: Any) -> int | None:
    numeric = as_float(value)
    if numeric is None:
        return None
    return int(round(numeric))


def parse_file_labels(filename: str) -> FileLabels | None:
    match = re.search(
        r"Model\s*-\s*(?P<ticker>[A-Za-z0-9]+)\s*-\s*(?P<period>(Early|Mid|Late)[A-Za-z]+\d{4})",
        filename,
        re.IGNORECASE,
    )
    if not match:
        return None

    ticker = match.group("ticker").upper()
    period_token = match.group("period")

    token_match = re.match(r"(?i)^(Early|Mid|Late)([A-Za-z]+)(\d{4})$", period_token)
    if not token_match:
        return None

    period_word = token_match.group(1).capitalize()
    month_token = token_match.group(2)[:3].lower()
    year = int(token_match.group(3))
    month = MONTH_MAP.get(month_token)
    if month is None:
        return None

    day = PERIOD_DAY_MAP[period_word.lower()]
    model_period = f"{period_word}{month_token.capitalize()}_{year}"
    model_date = date(year, month, day).isoformat()
    model = f"{ticker}_{model_period}"
    return FileLabels(model=model, ticker=ticker, model_period=model_period, model_date=model_date)


def build_output_path(input_folder: Path, output_folder: Path) -> Path:
    base = f"{input_folder.name}_PARAM"
    output_path = output_folder / f"{base}.xlsx"
    if not output_path.exists():
        return output_path

    index = 1
    while True:
        candidate = output_folder / f"{base}.{index}.xlsx"
        if not candidate.exists():
            return candidate
        index += 1


def list_input_files(folder: Path) -> list[Path]:
    if not folder.exists():
        return []

    files: list[Path] = []
    for file_path in sorted(folder.iterdir()):
        if not file_path.is_file():
            continue
        if file_path.name.startswith("~"):
            print(f"Skipped {file_path.name}: temporary file")
            continue
        if file_path.suffix.lower() != ".xlsx":
            print(f"Skipped {file_path.name}: not an .xlsx file")
            continue
        files.append(file_path)
    return files


def close_workbook_safe(wb: xw.Book) -> None:
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


def find_sheet_case_insensitive(wb: xw.Book, sheet_name: str) -> xw.Sheet | None:
    target = sheet_name.strip().lower()
    for sheet in wb.sheets:
        if sheet.name.strip().lower() == target:
            return sheet
    return None


def find_anchor(sheet: xw.Sheet, anchor_text: str = "max") -> tuple[int, int] | None:
    used = sheet.used_range
    values = as_matrix(used.value)
    base_row = used.row
    base_col = used.column

    for row_offset, row_values in enumerate(values):
        for col_offset, cell_value in enumerate(row_values):
            if isinstance(cell_value, str) and cell_value.strip().lower() == anchor_text:
                return base_row + row_offset, base_col + col_offset
    return None


def build_header_map(sheet: xw.Sheet, header_row: int, anchor_col: int) -> dict[str, int]:
    start_col = max(1, anchor_col - 40)
    end_col = anchor_col + 20
    values = sheet.range((header_row, start_col), (header_row, end_col)).value
    if not isinstance(values, list):
        values = [values]

    mapping: dict[str, int] = {}
    for idx, value in enumerate(values):
        normalized = normalize_header(value)
        if normalized:
            mapping.setdefault(normalized, start_col + idx)
    return mapping


def resolve_column(
    header_map: dict[str, int], aliases: Iterable[str], anchor_col: int, fallback_offset: int
) -> int:
    normalized_aliases = [normalize_header(alias) for alias in aliases]

    for alias in normalized_aliases:
        if alias in header_map:
            return header_map[alias]

    for alias in normalized_aliases:
        for existing_header, column in header_map.items():
            if alias and (alias in existing_header or existing_header in alias):
                return column

    return max(1, anchor_col + fallback_offset)


def read_rows(sheet: xw.Sheet, start_row: int, count: int, columns: dict[str, int]) -> list[dict[str, Any]]:
    if count <= 0 or not columns:
        return []

    min_col = min(columns.values())
    max_col = max(columns.values())
    end_row = start_row + count - 1
    matrix = as_matrix(sheet.range((start_row, min_col), (end_row, max_col)).value)

    rows: list[dict[str, Any]] = []
    for row_index in range(count):
        current = matrix[row_index] if row_index < len(matrix) else []
        row_values: dict[str, Any] = {}
        for key, col in columns.items():
            matrix_idx = col - min_col
            row_values[key] = current[matrix_idx] if matrix_idx < len(current) else None
        rows.append(row_values)
    return rows


def apply_empirical_avg_formulas(
    sheet: xw.Sheet, start_row: int, count: int, avg_col: int, quarterly_col: int, reported_col: int
) -> int:
    updates = 0
    q_rel = quarterly_col - avg_col
    r_rel = reported_col - avg_col
    formula = f'=IFERROR(IF(RC[{r_rel}]=0,"",RC[{q_rel}]/RC[{r_rel}]),"")'

    for row in range(start_row, start_row + count):
        target = sheet.cells(row, avg_col)
        if is_blank(target.value):
            target.formula2 = formula
            updates += 1
    return updates


def apply_regression_formulas(
    sheet: xw.Sheet,
    start_row: int,
    count: int,
    num_quarters_col: int,
    intercept_col: int,
    slope_col: int,
    x_col: int,
    y_col: int,
) -> int:
    updates = 0
    for idx in range(count):
        row = start_row + idx
        num_quarters = as_int(sheet.cells(row, num_quarters_col).value)
        if num_quarters is None:
            num_quarters = idx + 1

        if num_quarters < 2:
            continue

        range_start = row - num_quarters + 1
        if range_start < start_row:
            range_start = start_row
        if row - range_start + 1 < 2:
            continue

        intercept_formula = (
            f'=IFERROR(INTERCEPT(R{range_start}C{y_col}:R{row}C{y_col},'
            f'R{range_start}C{x_col}:R{row}C{x_col}),"")'
        )
        slope_formula = (
            f'=IFERROR(SLOPE(R{range_start}C{y_col}:R{row}C{y_col},'
            f'R{range_start}C{x_col}:R{row}C{x_col}),"")'
        )
        sheet.cells(row, intercept_col).formula2 = intercept_formula
        sheet.cells(row, slope_col).formula2 = slope_formula
        updates += 2
    return updates


def extract_empirical_rows(wb: xw.Book, labels: FileLabels, source_file: str) -> list[dict[str, Any]]:
    sheet = find_sheet_case_insensitive(wb, "Empirical Model")
    if sheet is None:
        print(f"Skipped {source_file}: missing 'Empirical Model' sheet")
        return []

    anchor = find_anchor(sheet, "max")
    if anchor is None:
        print(f"Skipped empirical extraction for {source_file}: missing 'max' anchor")
        return []

    anchor_row, anchor_col = anchor
    header_map = build_header_map(sheet, anchor_row, anchor_col)
    columns = {
        "num_quarters_used": resolve_column(
            header_map,
            ("num_quarters_used", "num quarters used", "n quarters", "quarters used", "num quarters"),
            anchor_col,
            -5,
        ),
        "last_quarter_used": resolve_column(
            header_map,
            ("last_quarter_used", "last quarter used", "last quarter"),
            anchor_col,
            -6,
        ),
        "avg_penetration_pct": resolve_column(
            header_map,
            ("avg_penetration_pct", "avg penetration %", "average penetration %", "penetration average"),
            anchor_col,
            -4,
        ),
        "forecast_value": resolve_column(
            header_map,
            ("estimated total sold", "forecast value", "estimated sold", "total sold"),
            anchor_col,
            -1,
        ),
        "reported_sales": resolve_column(
            header_map,
            ("reported_sales", "reported sales", "actual sales"),
            anchor_col,
            -2,
        ),
        "forecast_max": resolve_column(header_map, ("max",), anchor_col, 0),
        "forecast_min": resolve_column(header_map, ("min",), anchor_col, 1),
        "quarterly_sales": resolve_column(
            header_map,
            ("quarterly_sales", "quarterly sales"),
            anchor_col,
            -3,
        ),
        "growth_rate_pct": resolve_column(
            header_map,
            ("growth_rate_pct", "growth rate %", "growth rate"),
            anchor_col,
            -7,
        ),
        "sales_captured_in_db_pct": resolve_column(
            header_map,
            ("sales_captured_in_db_pct", "sales captured in db %", "captured in db %"),
            anchor_col,
            -8,
        ),
    }

    start_row = anchor_row + 1
    updates = apply_empirical_avg_formulas(
        sheet,
        start_row=start_row,
        count=N_QUARTERS,
        avg_col=columns["avg_penetration_pct"],
        quarterly_col=columns["quarterly_sales"],
        reported_col=columns["reported_sales"],
    )
    if updates:
        wb.app.calculate()

    raw_rows = read_rows(sheet, start_row=start_row, count=N_QUARTERS, columns=columns)
    extracted: list[dict[str, Any]] = []

    for idx, row in enumerate(raw_rows, start=1):
        forecast_max = as_float(row["forecast_max"])
        forecast_min = as_float(row["forecast_min"])
        forecast_value = as_float(row["forecast_value"])
        reported_sales = as_float(row["reported_sales"])
        avg_penetration = as_float(row["avg_penetration_pct"])

        if all(
            value is None
            for value in (forecast_max, forecast_min, forecast_value, reported_sales, avg_penetration)
        ):
            continue

        range_width = (
            forecast_max - forecast_min if forecast_max is not None and forecast_min is not None else None
        )
        num_quarters_used = as_int(row["num_quarters_used"])
        if num_quarters_used is None:
            num_quarters_used = idx

        extracted.append(
            {
                "model": labels.model,
                "ticker": labels.ticker,
                "model_period": labels.model_period,
                "model_date": labels.model_date,
                "method": "empirical",
                "parameter_name": "avg_penetration_pct",
                "parameter_value": avg_penetration,
                "num_quarters_used": num_quarters_used,
                "last_quarter_used": row["last_quarter_used"],
                "forecast_value": forecast_value,
                "actual_value": reported_sales,
                "forecast_max": forecast_max,
                "forecast_min": forecast_min,
                "range_width": range_width,
                "avg_penetration_pct": avg_penetration,
                "quarterly_sales": as_float(row["quarterly_sales"]),
                "reported_sales": reported_sales,
                "growth_rate_pct": as_float(row["growth_rate_pct"]),
                "sales_captured_in_db_pct": as_float(row["sales_captured_in_db_pct"]),
                "source_file": source_file,
            }
        )
    return extracted


def extract_regression_rows(wb: xw.Book, labels: FileLabels, source_file: str) -> list[dict[str, Any]]:
    sheet = find_sheet_case_insensitive(wb, "Regression Model")
    if sheet is None:
        print(f"Skipped {source_file}: missing 'Regression Model' sheet")
        return []

    anchor = find_anchor(sheet, "max")
    if anchor is None:
        print(f"Skipped regression extraction for {source_file}: missing 'max' anchor")
        return []

    anchor_row, anchor_col = anchor
    header_map = build_header_map(sheet, anchor_row, anchor_col)
    columns = {
        "num_quarters_used": resolve_column(
            header_map,
            ("num_quarters_used", "num quarters used", "num quarters", "quarters used"),
            anchor_col,
            -5,
        ),
        "forecast_value": resolve_column(
            header_map,
            ("tot fcst w/o sa", "total forecast w/o sa", "forecast without sa"),
            anchor_col,
            -1,
        ),
        "actual_value": resolve_column(
            header_map,
            ("actual_value", "actual value", "reported sales", "actual sales"),
            anchor_col,
            -4,
        ),
        "forecast_max": resolve_column(header_map, ("max",), anchor_col, 0),
        "forecast_min": resolve_column(header_map, ("min",), anchor_col, 1),
        "intercept": resolve_column(
            header_map,
            ("intercept",),
            anchor_col,
            -3,
        ),
        "slope": resolve_column(
            header_map,
            ("slope",),
            anchor_col,
            -2,
        ),
    }

    y_col = anchor_col - 7
    x_col = anchor_col - 11
    start_row = anchor_row + 1

    updates = apply_regression_formulas(
        sheet,
        start_row=start_row,
        count=N_QUARTERS,
        num_quarters_col=columns["num_quarters_used"],
        intercept_col=columns["intercept"],
        slope_col=columns["slope"],
        x_col=x_col,
        y_col=y_col,
    )
    if updates:
        wb.app.calculate()

    raw_rows = read_rows(sheet, start_row=start_row, count=N_QUARTERS, columns=columns)
    extracted: list[dict[str, Any]] = []
    previous_signature: tuple[Any, ...] | None = None

    for idx, row in enumerate(raw_rows, start=1):
        num_quarters_used = as_int(row["num_quarters_used"])
        if num_quarters_used is None:
            num_quarters_used = idx

        forecast_value = as_float(row["forecast_value"])
        forecast_max = as_float(row["forecast_max"])
        forecast_min = as_float(row["forecast_min"])
        intercept = as_float(row["intercept"])
        slope = as_float(row["slope"])
        actual_value = as_float(row["actual_value"])

        if all(
            value is None for value in (forecast_value, forecast_max, forecast_min, intercept, slope)
        ):
            continue

        range_width = (
            forecast_max - forecast_min if forecast_max is not None and forecast_min is not None else None
        )
        signature = (
            num_quarters_used,
            forecast_value,
            forecast_max,
            forecast_min,
            intercept,
            slope,
        )
        if previous_signature is not None and signature == previous_signature:
            continue
        previous_signature = signature

        extracted.append(
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
    return extracted


def autofit_worksheet(ws: Any, headers: list[str]) -> None:
    for idx, header in enumerate(headers, start=1):
        max_len = len(header)
        for row in range(2, ws.max_row + 1):
            value = ws.cell(row=row, column=idx).value
            if value is None:
                continue
            max_len = min(60, max(max_len, len(str(value))))
        ws.column_dimensions[get_column_letter(idx)].width = min(60, max_len + 2)


def write_output_workbook(
    destination: Path, empirical_rows: list[dict[str, Any]], regression_rows: list[dict[str, Any]]
) -> None:
    workbook = Workbook()
    default_sheet = workbook.active
    workbook.remove(default_sheet)

    empirical_sheet = workbook.create_sheet("empirical_candidates")
    regression_sheet = workbook.create_sheet("regression_candidates")

    empirical_sheet.append(EMPIRICAL_HEADERS)
    for row in empirical_rows:
        empirical_sheet.append([row.get(col) for col in EMPIRICAL_HEADERS])

    regression_sheet.append(REGRESSION_HEADERS)
    for row in regression_rows:
        regression_sheet.append([row.get(col) for col in REGRESSION_HEADERS])

    for ws, headers in (
        (empirical_sheet, EMPIRICAL_HEADERS),
        (regression_sheet, REGRESSION_HEADERS),
    ):
        for cell in ws[1]:
            cell.font = Font(bold=True)
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions
        autofit_worksheet(ws, headers)

    workbook.save(destination)
    workbook.close()


def main() -> int:
    if not input_dir.exists():
        print(f"Input folder does not exist: {input_dir}")
        return 1

    output_dir.mkdir(parents=True, exist_ok=True)
    files = list_input_files(input_dir)
    output_path = build_output_path(input_dir, output_dir)

    app = xw.App(visible=False, add_book=False)
    app.display_alerts = False
    app.screen_updating = False
    app.calculation = "manual"

    processed_files = 0
    empirical_rows: list[dict[str, Any]] = []
    regression_rows: list[dict[str, Any]] = []

    try:
        for file_path in files:
            labels = parse_file_labels(file_path.name)
            if labels is None:
                print(f"Skipped {file_path.name}: could not parse ticker/model period from filename")
                continue

            print(f"Processing {file_path.name}")
            wb: xw.Book | None = None
            try:
                wb = app.books.open(str(file_path), update_links=False)
                empirical_rows.extend(extract_empirical_rows(wb, labels, file_path.name))
                regression_rows.extend(extract_regression_rows(wb, labels, file_path.name))
                processed_files += 1
            except Exception as exc:
                print(f"Skipped {file_path.name}: processing error -> {exc}")
            finally:
                if wb is not None:
                    close_workbook_safe(wb)
    finally:
        app.quit()

    write_output_workbook(output_path, empirical_rows, regression_rows)

    print(f"Output path: {output_path}")
    print(f"Files processed: {processed_files}")
    print(f"Empirical rows: {len(empirical_rows)}")
    print(f"Regression rows: {len(regression_rows)}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
