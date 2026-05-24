from __future__ import annotations

import re
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Any, Iterable, Optional, Sequence

import xlwings as xw
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

# Configure these two paths before running.
input_dir = "./input"
output_dir = "./output"

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

PERIOD_TO_DAY = {"early": 5, "mid": 15, "late": 25}


@dataclass
class FileMetadata:
    model: str
    ticker: str
    model_period: str
    model_date: str


@dataclass
class SheetIndex:
    labels: list[tuple[str, int, int]]


def normalize_label(value: Any) -> str:
    if value is None:
        return ""
    label = str(value).strip().replace("\n", " ")
    label = re.sub(r"\s+", " ", label)
    return label.lower()


def normalize_2d(values: Any) -> list[list[Any]]:
    if isinstance(values, list):
        if not values:
            return []
        if isinstance(values[0], list):
            return values
        return [values]
    return [[values]]


def to_float(value: Any) -> Optional[float]:
    if value is None:
        return None
    if isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        return float(value)
    text = str(value).strip()
    if not text:
        return None
    is_percent = "%" in text
    text = text.replace(",", "").replace("%", "")
    try:
        parsed = float(text)
    except ValueError:
        return None
    return parsed / 100 if is_percent else parsed


def compact_signature(value: Any) -> Any:
    num = to_float(value)
    if num is not None:
        return round(num, 10)
    if value is None:
        return None
    return str(value).strip()


def subtract(a: Any, b: Any) -> Optional[float]:
    left = to_float(a)
    right = to_float(b)
    if left is None or right is None:
        return None
    return left - right


def series_to_text(values: Iterable[Any]) -> Optional[str]:
    compact = [str(v) for v in values if v not in (None, "")]
    if not compact:
        return None
    return "|".join(compact)


def read_col_values(sheet: xw.Sheet, col: int, start_row: int, end_row: int) -> list[Any]:
    if col < 1 or start_row < 1 or end_row < start_row:
        return []
    data = sheet.range((start_row, col), (end_row, col)).value
    matrix = normalize_2d(data)
    values: list[Any] = []
    for row in matrix:
        values.append(row[0] if row else None)
    return values


def build_sheet_index(sheet: xw.Sheet) -> SheetIndex:
    used = sheet.used_range
    start_row = used.row
    start_col = used.column
    matrix = normalize_2d(used.value)
    labels: list[tuple[str, int, int]] = []

    for r_offset, row_values in enumerate(matrix):
        for c_offset, cell_value in enumerate(row_values):
            if isinstance(cell_value, str):
                normalized = normalize_label(cell_value)
                if normalized:
                    labels.append((normalized, start_row + r_offset, start_col + c_offset))
    return SheetIndex(labels=labels)


def find_label(
    index: SheetIndex,
    needles: Sequence[str],
    *,
    exact: bool = False,
    same_col: Optional[int] = None,
    min_row: Optional[int] = None,
) -> Optional[tuple[int, int]]:
    normalized_needles = [n.lower() for n in needles]
    for label, row, col in index.labels:
        if same_col is not None and col != same_col:
            continue
        if min_row is not None and row < min_row:
            continue
        if exact:
            if len(normalized_needles) == 1 and label == normalized_needles[0]:
                return row, col
        else:
            if all(needle in label for needle in normalized_needles):
                return row, col
    return None


def locate_value_cell(sheet: xw.Sheet, label_row: int, label_col: int, span: int = 6) -> xw.Range:
    for offset in range(1, span + 1):
        probe = sheet.range((label_row, label_col + offset))
        if probe.value not in (None, ""):
            return probe
    for offset in range(1, span + 1):
        probe = sheet.range((label_row, label_col - offset))
        if probe.value not in (None, ""):
            return probe
    return sheet.range((label_row, label_col + 1))


def find_value_cell_for_keywords(
    sheet: xw.Sheet,
    index: SheetIndex,
    keyword_groups: Sequence[Sequence[str]],
    fallback: Optional[tuple[int, int]] = None,
) -> Optional[xw.Range]:
    for group in keyword_groups:
        found = find_label(index, group)
        if found:
            row, col = found
            return locate_value_cell(sheet, row, col)
    if fallback is None:
        return None
    row, col = fallback
    if row < 1 or col < 1:
        return None
    return sheet.range((row, col))


def find_anchor_cells(
    sheet: xw.Sheet, index: SheetIndex
) -> Optional[tuple[int, int, xw.Range, Optional[xw.Range]]]:
    max_anchor = find_label(index, ["max"], exact=True)
    if not max_anchor:
        return None
    max_row, max_col = max_anchor
    max_value_cell = locate_value_cell(sheet, max_row, max_col)

    min_row = max_row + 1
    if normalize_label(sheet.range((min_row, max_col)).value) != "min":
        min_anchor = find_label(index, ["min"], exact=True, same_col=max_col, min_row=max_row)
        if min_anchor:
            min_row = min_anchor[0]
        else:
            min_row = -1
    min_value_cell = locate_value_cell(sheet, min_row, max_col) if min_row > 0 else None
    return max_row, max_col, max_value_cell, min_value_cell


def parse_filename_metadata(file_path: Path) -> Optional[FileMetadata]:
    # Expected style:
    # MedMiner_Model - AORT - MidJan2026_Send.xlsx
    parts = [part.strip() for part in file_path.stem.split(" - ")]
    if len(parts) < 3:
        return None

    ticker = parts[1].upper()
    period_chunk = parts[2].split("_")[0]
    match = re.match(r"^(Early|Mid|Late)([A-Za-z]+)(\d{4})$", period_chunk, flags=re.IGNORECASE)
    if not match:
        return None

    period_word, month_token, year_token = match.groups()
    period_key = period_word.lower()
    day = PERIOD_TO_DAY.get(period_key)
    if day is None:
        return None

    month_num = parse_month(month_token)
    if month_num is None:
        return None

    month_abbrev = datetime(2000, month_num, 1).strftime("%b")
    model_period = f"{period_word.title()}{month_abbrev}_{year_token}"
    model_date = f"{year_token}-{month_num:02d}-{day:02d}"
    model = f"{ticker}_{model_period}"
    return FileMetadata(model=model, ticker=ticker, model_period=model_period, model_date=model_date)


def parse_month(month_token: str) -> Optional[int]:
    cleaned = month_token.strip().title()
    if not cleaned:
        return None
    for fmt in ("%b", "%B"):
        try:
            return datetime.strptime(cleaned, fmt).month
        except ValueError:
            pass
    if len(cleaned) > 3:
        for fmt in ("%b", "%B"):
            try:
                return datetime.strptime(cleaned[:3], fmt).month
            except ValueError:
                pass
    return None


def get_sheet_by_name(wb: xw.Book, target_name: str) -> Optional[xw.Sheet]:
    target = normalize_label(target_name)
    for sheet in wb.sheets:
        if normalize_label(sheet.name) == target:
            return sheet
    return None


def safe_close_workbook(wb: xw.Book) -> None:
    try:
        wb.close(save=False)
        return
    except TypeError:
        pass
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


def extract_empirical_rows(sheet: xw.Sheet, metadata: FileMetadata, source_file: str) -> list[dict[str, Any]]:
    index = build_sheet_index(sheet)
    anchors = find_anchor_cells(sheet, index)
    if anchors is None:
        print(f"Skipped empirical for {source_file}: max anchor not found")
        return []

    anchor_row, anchor_col, max_cell, min_cell = anchors

    # Anchor-based offsets (keeps scanning to a minimum).
    quarter_col = anchor_col - 11
    quarterly_sales_col = anchor_col - 7
    reported_sales_col = anchor_col - 6
    penetration_col = anchor_col - 5
    growth_rate_col = anchor_col - 4
    captured_col = anchor_col - 3

    data_end_row = anchor_row - 2
    if data_end_row < 1:
        print(f"Skipped empirical for {source_file}: invalid data region")
        return []

    avg_pen_cell = find_value_cell_for_keywords(
        sheet,
        index,
        [
            ("avg", "penetration"),
            ("average", "penetration"),
        ],
        fallback=(anchor_row - 1, anchor_col - 2),
    )
    if avg_pen_cell is None:
        print(f"Skipped empirical for {source_file}: avg penetration cell not found")
        return []

    num_quarters_cell = find_value_cell_for_keywords(
        sheet,
        index,
        [
            ("num", "quarters"),
            ("quarters", "used"),
        ],
        fallback=None,
    )
    forecast_cell = find_value_cell_for_keywords(
        sheet,
        index,
        [
            ("estimated", "total", "sold"),
            ("total", "sold", "estimate"),
            ("tot", "fcst"),
        ],
        fallback=(anchor_row, anchor_col + 1),
    )
    actual_cell = find_value_cell_for_keywords(
        sheet,
        index,
        [
            ("reported", "sales"),
            ("actual", "sales"),
        ],
        fallback=(anchor_row + 1, anchor_col + 1),
    )

    rows: list[dict[str, Any]] = []
    for n_quarters in range(1, N_QUARTERS + 1):
        start_row = data_end_row - n_quarters + 1
        if start_row < 1:
            break

        avg_formula = (
            f"=AVERAGE(R{start_row}C{penetration_col}:R{data_end_row}C{penetration_col})"
        )
        try:
            avg_pen_cell.formula2 = avg_formula
        except Exception:
            avg_pen_cell.formula = avg_formula

        if num_quarters_cell is not None:
            num_quarters_cell.value = n_quarters

        # Recalculate only after formula changes.
        sheet.book.app.calculate()

        quarter_labels = read_col_values(sheet, quarter_col, start_row, data_end_row)
        quarterly_sales = read_col_values(sheet, quarterly_sales_col, start_row, data_end_row)
        reported_sales = read_col_values(sheet, reported_sales_col, start_row, data_end_row)
        growth_rates = read_col_values(sheet, growth_rate_col, start_row, data_end_row)
        captured_pcts = read_col_values(sheet, captured_col, start_row, data_end_row)

        forecast_max = max_cell.value if max_cell else None
        forecast_min = min_cell.value if min_cell else None
        row = {
            "model": metadata.model,
            "ticker": metadata.ticker,
            "model_period": metadata.model_period,
            "model_date": metadata.model_date,
            "method": "empirical",
            "parameter_name": "avg_penetration_pct",
            "parameter_value": avg_pen_cell.value,
            "num_quarters_used": n_quarters,
            "last_quarter_used": quarter_labels[-1] if quarter_labels else None,
            "forecast_value": forecast_cell.value if forecast_cell else None,
            "actual_value": actual_cell.value if actual_cell else None,
            "forecast_max": forecast_max,
            "forecast_min": forecast_min,
            "range_width": subtract(forecast_max, forecast_min),
            "avg_penetration_pct": avg_pen_cell.value,
            "quarterly_sales": series_to_text(quarterly_sales),
            "reported_sales": series_to_text(reported_sales),
            "growth_rate_pct": series_to_text(growth_rates),
            "sales_captured_in_db_pct": series_to_text(captured_pcts),
            "source_file": source_file,
        }
        rows.append(row)

    return rows


def extract_regression_rows(sheet: xw.Sheet, metadata: FileMetadata, source_file: str) -> list[dict[str, Any]]:
    index = build_sheet_index(sheet)
    anchors = find_anchor_cells(sheet, index)
    if anchors is None:
        print(f"Skipped regression for {source_file}: max anchor not found")
        return []

    anchor_row, anchor_col, max_cell, min_cell = anchors
    y_col = anchor_col - 7
    x_col = anchor_col - 11
    data_end_row = anchor_row - 2
    if data_end_row < 1:
        print(f"Skipped regression for {source_file}: invalid data region")
        return []

    num_quarters_cell = find_value_cell_for_keywords(
        sheet,
        index,
        [
            ("num", "quarters"),
            ("quarters", "used"),
        ],
        fallback=None,
    )
    forecast_cell = find_value_cell_for_keywords(
        sheet,
        index,
        [
            ("tot", "fcst", "w/o", "sa"),
            ("total", "forecast", "without", "sa"),
        ],
        fallback=(anchor_row, anchor_col + 1),
    )
    actual_cell = find_value_cell_for_keywords(
        sheet,
        index,
        [
            ("actual", "sales"),
            ("reported", "sales"),
        ],
        fallback=None,
    )

    # Keep temporary formulas close to anchor and in R1C1 mode.
    intercept_cell = sheet.range((anchor_row - 1, anchor_col + 1))
    slope_cell = sheet.range((anchor_row - 1, anchor_col + 2))

    rows: list[dict[str, Any]] = []
    prev_signature: Optional[tuple[Any, ...]] = None
    for n_quarters in range(1, N_QUARTERS + 1):
        start_row = data_end_row - n_quarters + 1
        if start_row < 1:
            break

        if num_quarters_cell is not None:
            num_quarters_cell.value = n_quarters

        intercept_formula = (
            f"=INTERCEPT(R{start_row}C{y_col}:R{data_end_row}C{y_col},"
            f"R{start_row}C{x_col}:R{data_end_row}C{x_col})"
        )
        slope_formula = (
            f"=SLOPE(R{start_row}C{y_col}:R{data_end_row}C{y_col},"
            f"R{start_row}C{x_col}:R{data_end_row}C{x_col})"
        )
        try:
            intercept_cell.formula2 = intercept_formula
            slope_cell.formula2 = slope_formula
        except Exception:
            intercept_cell.formula = intercept_formula
            slope_cell.formula = slope_formula

        # Recalculate only after formula changes.
        sheet.book.app.calculate()

        forecast_max = max_cell.value if max_cell else None
        forecast_min = min_cell.value if min_cell else None
        signature = (
            compact_signature(intercept_cell.value),
            compact_signature(slope_cell.value),
            compact_signature(forecast_cell.value if forecast_cell else None),
            compact_signature(forecast_max),
            compact_signature(forecast_min),
        )

        # Some models duplicate the final row when n_quarters reaches the cap.
        if n_quarters == N_QUARTERS and prev_signature is not None and signature == prev_signature:
            continue

        row = {
            "model": metadata.model,
            "ticker": metadata.ticker,
            "model_period": metadata.model_period,
            "model_date": metadata.model_date,
            "method": "regression",
            "parameter_name": "num_quarters_used",
            "parameter_value": n_quarters,
            "num_quarters_used": n_quarters,
            "forecast_value": forecast_cell.value if forecast_cell else None,
            "actual_value": actual_cell.value if actual_cell else None,
            "forecast_max": forecast_max,
            "forecast_min": forecast_min,
            "range_width": subtract(forecast_max, forecast_min),
            "intercept": intercept_cell.value,
            "slope": slope_cell.value,
            "source_file": source_file,
        }
        rows.append(row)
        prev_signature = signature

    return rows


def write_sheet(ws: Any, columns: Sequence[str], rows: Sequence[dict[str, Any]]) -> None:
    ws.append(list(columns))
    for col_idx in range(1, len(columns) + 1):
        ws.cell(row=1, column=col_idx).font = Font(bold=True)

    for row_data in rows:
        ws.append([row_data.get(column) for column in columns])

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions

    for col_idx, column_name in enumerate(columns, start=1):
        max_len = len(column_name)
        for row_idx in range(2, ws.max_row + 1):
            value = ws.cell(row=row_idx, column=col_idx).value
            if value is None:
                continue
            max_len = max(max_len, len(str(value)))
        ws.column_dimensions[get_column_letter(col_idx)].width = min(max(max_len + 2, 12), 72)


def write_output_workbook(
    output_path: Path,
    empirical_rows: Sequence[dict[str, Any]],
    regression_rows: Sequence[dict[str, Any]],
) -> None:
    wb = Workbook()
    wb.remove(wb.active)

    empirical_ws = wb.create_sheet("empirical_candidates")
    regression_ws = wb.create_sheet("regression_candidates")

    write_sheet(empirical_ws, EMPIRICAL_COLUMNS, empirical_rows)
    write_sheet(regression_ws, REGRESSION_COLUMNS, regression_rows)
    wb.save(output_path)


def choose_output_path(input_path: Path, output_path: Path) -> Path:
    folder_name = input_path.name
    base = output_path / f"{folder_name}_PARAM.xlsx"
    if not base.exists():
        return base

    suffix = 1
    while True:
        candidate = output_path / f"{folder_name}_PARAM.{suffix}.xlsx"
        if not candidate.exists():
            return candidate
        suffix += 1


def process_workbook(
    wb: xw.Book,
    metadata: FileMetadata,
    source_file: str,
) -> tuple[list[dict[str, Any]], list[dict[str, Any]]]:
    empirical_rows: list[dict[str, Any]] = []
    regression_rows: list[dict[str, Any]] = []

    empirical_sheet = get_sheet_by_name(wb, "Empirical Model")
    if empirical_sheet is None:
        print(f"Skipped empirical for {source_file}: missing sheet 'Empirical Model'")
    else:
        empirical_rows = extract_empirical_rows(empirical_sheet, metadata, source_file)

    regression_sheet = get_sheet_by_name(wb, "Regression Model")
    if regression_sheet is None:
        print(f"Skipped regression for {source_file}: missing sheet 'Regression Model'")
    else:
        regression_rows = extract_regression_rows(regression_sheet, metadata, source_file)

    return empirical_rows, regression_rows


def main() -> None:
    input_path = Path(input_dir).expanduser().resolve()
    output_path = Path(output_dir).expanduser().resolve()

    if not input_path.exists() or not input_path.is_dir():
        raise FileNotFoundError(f"Input directory does not exist: {input_path}")
    output_path.mkdir(parents=True, exist_ok=True)

    output_file = choose_output_path(input_path, output_path)
    empirical_rows: list[dict[str, Any]] = []
    regression_rows: list[dict[str, Any]] = []
    processed_files = 0

    app: Optional[xw.App] = None
    try:
        app = xw.App(visible=False, add_book=False)
        app.display_alerts = False
        app.screen_updating = False

        for file_path in sorted(input_path.iterdir()):
            if not file_path.is_file():
                print(f"Skipped: {file_path.name} (not a file)")
                continue
            if file_path.name.startswith("~"):
                print(f"Skipped: {file_path.name} (temporary file)")
                continue
            if file_path.suffix.lower() != ".xlsx":
                print(f"Skipped: {file_path.name} (not an .xlsx file)")
                continue

            metadata = parse_filename_metadata(file_path)
            if metadata is None:
                print(
                    f"Skipped: {file_path.name} "
                    "(filename does not match '... - TICKER - MidJan2026_...')"
                )
                continue

            wb: Optional[xw.Book] = None
            try:
                wb = app.books.open(str(file_path), update_links=False)
                file_empirical, file_regression = process_workbook(wb, metadata, file_path.name)
                empirical_rows.extend(file_empirical)
                regression_rows.extend(file_regression)
                processed_files += 1
                print(f"Processed: {file_path.name}")
            except Exception as exc:
                print(f"Skipped: {file_path.name} (workbook error: {exc})")
            finally:
                if wb is not None:
                    safe_close_workbook(wb)
    finally:
        if app is not None:
            app.quit()

    write_output_workbook(output_file, empirical_rows, regression_rows)
    print(f"Output path: {output_file}")
    print(f"Number of files processed: {processed_files}")
    print(f"Number of empirical rows: {len(empirical_rows)}")
    print(f"Number of regression rows: {len(regression_rows)}")


if __name__ == "__main__":
    main()
