#!/usr/bin/env python3
"""
Extract empirical and regression candidates from model workbooks.

Workflow summary:
1) Open one hidden Excel app for the full run.
2) Open each source workbook once and process both model sheets while open.
3) Never save source workbooks.
4) Write one output workbook with:
   - empirical_candidates
   - regression_candidates
"""

from __future__ import annotations

import re
from dataclasses import dataclass
from datetime import date
from pathlib import Path
from typing import Any, Iterable

try:
    import xlwings as xw
except ImportError as exc:
    raise SystemExit(
        "Missing dependency 'xlwings'. Install it with: pip install xlwings"
    ) from exc

try:
    from openpyxl import Workbook
    from openpyxl.styles import Font
    from openpyxl.utils import get_column_letter
except ImportError as exc:
    raise SystemExit(
        "Missing dependency 'openpyxl'. Install it with: pip install openpyxl"
    ) from exc


# ---------------------------------------------------------------------------
# User-configurable inputs
# ---------------------------------------------------------------------------
input_dir = "./input"
output_dir = "./output"


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

PHASE_DAY_MAP = {"Early": 5, "Mid": 15, "Late": 25}
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
MONTH_ABBR = {
    1: "Jan",
    2: "Feb",
    3: "Mar",
    4: "Apr",
    5: "May",
    6: "Jun",
    7: "Jul",
    8: "Aug",
    9: "Sep",
    10: "Oct",
    11: "Nov",
    12: "Dec",
}

# Fallback offsets (relative to the located "max" anchor column) used when
# label-based column discovery is unavailable.
EMPIRICAL_FALLBACK_OFFSETS = {
    "quarter_col": -12,
    "quarterly_sales_col": -11,
    "reported_sales_col": -7,
    "growth_rate_col": -6,
    "captured_pct_col": -5,
}


@dataclass(frozen=True)
class FileMeta:
    model: str
    ticker: str
    model_period: str
    model_date: str


@dataclass(frozen=True)
class SheetSnapshot:
    top_row: int
    left_col: int
    values: list[list[Any]]
    text_cells: list[tuple[int, int, str]]


def normalize_text(value: str) -> str:
    return re.sub(r"\s+", " ", value.strip().lower())


def coerce_2d(values: Any) -> list[list[Any]]:
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
    if isinstance(value, (int, float)):
        return float(value)
    if isinstance(value, str):
        text = value.strip().replace(",", "")
        if not text:
            return None
        try:
            if text.endswith("%"):
                return float(text[:-1]) / 100.0
            return float(text)
        except ValueError:
            return None
    return None


def safe_col(anchor_col: int, offset: int) -> int:
    return max(1, anchor_col + offset)


def snapshot_sheet(sheet: xw.Sheet) -> SheetSnapshot:
    used = sheet.used_range
    values_2d = coerce_2d(used.value)
    text_cells: list[tuple[int, int, str]] = []
    for r_off, row_values in enumerate(values_2d):
        for c_off, value in enumerate(row_values):
            if isinstance(value, str):
                norm = normalize_text(value)
                if norm:
                    text_cells.append((used.row + r_off, used.column + c_off, norm))
    return SheetSnapshot(
        top_row=used.row, left_col=used.column, values=values_2d, text_cells=text_cells
    )


def find_anchor(snapshot: SheetSnapshot, anchor_text: str = "max") -> tuple[int, int]:
    target = normalize_text(anchor_text)
    for row, col, text in snapshot.text_cells:
        if text == target:
            return row, col
    raise ValueError(f'Anchor "{anchor_text}" not found.')


def label_matches(cell_text: str, alias: str) -> bool:
    if alias in {"max", "min"}:
        return cell_text == alias
    return cell_text == alias or alias in cell_text


def find_label_cell(
    snapshot: SheetSnapshot,
    aliases: Iterable[str],
    anchor: tuple[int, int],
) -> tuple[int, int] | None:
    normalized_aliases = [normalize_text(alias) for alias in aliases]
    matches: list[tuple[int, int, int]] = []
    anchor_row, anchor_col = anchor
    for row, col, text in snapshot.text_cells:
        if any(label_matches(text, alias) for alias in normalized_aliases):
            distance = abs(anchor_row - row) + abs(anchor_col - col)
            matches.append((distance, row, col))
    if not matches:
        return None
    matches.sort(key=lambda item: item[0])
    _, row, col = matches[0]
    return row, col


def find_label_column(
    snapshot: SheetSnapshot,
    aliases: Iterable[str],
    anchor: tuple[int, int],
) -> int | None:
    cell = find_label_cell(snapshot, aliases, anchor)
    return None if cell is None else cell[1]


def read_numeric_neighbor(sheet: xw.Sheet, row: int, col: int) -> float | None:
    for d_row, d_col in ((0, 1), (1, 0), (0, 2), (1, 1), (0, -1), (-1, 1)):
        value = to_float(sheet.cells(row + d_row, col + d_col).value)
        if value is not None:
            return value
    return None


def collect_history_rows(
    sheet: xw.Sheet,
    start_row: int,
    required_cols: list[int],
    max_scan_rows: int = 160,
) -> list[int]:
    rows: list[int] = []
    row = start_row
    scanned = 0
    started = False
    while row >= 1 and scanned < max_scan_rows:
        scanned += 1
        values = [to_float(sheet.cells(row, col).value) for col in required_cols]
        is_valid = all(value is not None for value in values)
        if is_valid:
            rows.append(row)
            started = True
        elif started:
            break
        row -= 1
    rows.reverse()
    return rows


def month_number(month_token: str) -> int | None:
    cleaned = re.sub(r"[^A-Za-z]", "", month_token).lower()
    if len(cleaned) < 3:
        return None
    return MONTH_MAP.get(cleaned[:3])


def parse_file_meta(file_name: str) -> FileMeta:
    stem = Path(file_name).stem
    parts = [part.strip() for part in stem.split(" - ")]

    ticker = parts[1] if len(parts) >= 2 else "UNKNOWN"
    period_token = parts[2] if len(parts) >= 3 else (parts[-1] if parts else stem)
    period_token = re.sub(r"[_\s-]*send$", "", period_token, flags=re.IGNORECASE).strip()

    match = re.search(r"(Early|Mid|Late)([A-Za-z]+)(\d{4})", period_token, re.IGNORECASE)
    model_period = period_token.replace(" ", "_") if period_token else "unknown_period"
    model_date = ""

    if match:
        phase_raw, month_raw, year_raw = match.groups()
        phase = phase_raw.title()
        year = int(year_raw)
        month_num = month_number(month_raw)
        if month_num is not None:
            model_period = f"{phase}{MONTH_ABBR[month_num]}_{year}"
            model_date = date(year, month_num, PHASE_DAY_MAP[phase]).isoformat()

    model = f"{ticker}_{model_period}"
    return FileMeta(model=model, ticker=ticker, model_period=model_period, model_date=model_date)


def get_empirical_columns(snapshot: SheetSnapshot, anchor_col: int, anchor: tuple[int, int]) -> dict[str, int]:
    quarter_col = find_label_column(snapshot, ["quarter", "qtr"], anchor)
    quarterly_sales_col = find_label_column(
        snapshot, ["quarterly sales", "qtr sales", "db sales"], anchor
    )
    reported_sales_col = find_label_column(
        snapshot, ["reported sales", "actual sales"], anchor
    )
    growth_rate_col = find_label_column(snapshot, ["growth rate", "growth %"], anchor)
    captured_pct_col = find_label_column(
        snapshot, ["sales captured in db", "captured in db", "captured %"], anchor
    )

    return {
        "quarter_col": quarter_col or safe_col(anchor_col, EMPIRICAL_FALLBACK_OFFSETS["quarter_col"]),
        "quarterly_sales_col": quarterly_sales_col
        or safe_col(anchor_col, EMPIRICAL_FALLBACK_OFFSETS["quarterly_sales_col"]),
        "reported_sales_col": reported_sales_col
        or safe_col(anchor_col, EMPIRICAL_FALLBACK_OFFSETS["reported_sales_col"]),
        "growth_rate_col": growth_rate_col
        or safe_col(anchor_col, EMPIRICAL_FALLBACK_OFFSETS["growth_rate_col"]),
        "captured_pct_col": captured_pct_col
        or safe_col(anchor_col, EMPIRICAL_FALLBACK_OFFSETS["captured_pct_col"]),
    }


def extract_empirical_rows(
    wb: xw.Book,
    meta: FileMeta,
    source_file: str,
) -> list[dict[str, Any]]:
    try:
        sheet = wb.sheets["Empirical Model"]
    except Exception:
        print(f"Skipped empirical extraction for {source_file}: missing 'Empirical Model' sheet")
        return []

    snapshot = snapshot_sheet(sheet)
    try:
        anchor_row, anchor_col = find_anchor(snapshot, "max")
    except ValueError as exc:
        print(f"Skipped empirical extraction for {source_file}: {exc}")
        return []
    anchor = (anchor_row, anchor_col)

    columns = get_empirical_columns(snapshot, anchor_col, anchor)
    history_rows = collect_history_rows(
        sheet,
        start_row=anchor_row - 1,
        required_cols=[columns["quarterly_sales_col"], columns["reported_sales_col"]],
    )
    if not history_rows:
        print(f"Skipped empirical extraction for {source_file}: no usable historical rows found")
        return []

    estimated_total_cell = find_label_cell(
        snapshot, ["estimated total sold", "estimated total", "total sold"], anchor
    )
    actual_sales_cell = find_label_cell(snapshot, ["reported sales", "actual sales"], anchor)
    min_label_cell = find_label_cell(snapshot, ["min"], anchor)

    forecast_max = read_numeric_neighbor(sheet, anchor_row, anchor_col)
    forecast_min = (
        read_numeric_neighbor(sheet, min_label_cell[0], min_label_cell[1])
        if min_label_cell
        else None
    )
    range_width = (
        (forecast_max - forecast_min)
        if forecast_max is not None and forecast_min is not None
        else None
    )

    helper_avg_cell = sheet.cells(anchor_row + 2, anchor_col + 2)
    rows: list[dict[str, Any]] = []
    max_quarters = min(10, len(history_rows))

    for num_quarters in range(1, max_quarters + 1):
        selected_rows = history_rows[-num_quarters:]
        start_row = selected_rows[0]
        end_row = selected_rows[-1]

        helper_avg_cell.formula2 = (
            f"=AVERAGE("
            f"R{start_row}C{columns['quarterly_sales_col']}:R{end_row}C{columns['quarterly_sales_col']}/"
            f"R{start_row}C{columns['reported_sales_col']}:R{end_row}C{columns['reported_sales_col']}"
            f")"
        )
        wb.app.calculate()

        avg_penetration_pct = to_float(helper_avg_cell.value)
        quarterly_sales = to_float(sheet.cells(end_row, columns["quarterly_sales_col"]).value)
        reported_sales = to_float(sheet.cells(end_row, columns["reported_sales_col"]).value)
        growth_rate_pct = to_float(sheet.cells(end_row, columns["growth_rate_col"]).value)
        sales_captured_in_db_pct = to_float(sheet.cells(end_row, columns["captured_pct_col"]).value)
        last_quarter_used = sheet.cells(end_row, columns["quarter_col"]).value

        forecast_value = (
            read_numeric_neighbor(sheet, estimated_total_cell[0], estimated_total_cell[1])
            if estimated_total_cell
            else None
        )
        if forecast_value is None and reported_sales is not None and avg_penetration_pct:
            forecast_value = reported_sales / avg_penetration_pct

        actual_value = (
            read_numeric_neighbor(sheet, actual_sales_cell[0], actual_sales_cell[1])
            if actual_sales_cell
            else None
        )
        if actual_value is None:
            actual_value = reported_sales

        rows.append(
            {
                "model": meta.model,
                "ticker": meta.ticker,
                "model_period": meta.model_period,
                "model_date": meta.model_date,
                "method": "empirical",
                "parameter_name": "avg_penetration_pct",
                "parameter_value": avg_penetration_pct,
                "num_quarters_used": num_quarters,
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
                "sales_captured_in_db_pct": sales_captured_in_db_pct,
                "source_file": source_file,
            }
        )

    helper_avg_cell.value = None
    return rows


def regression_rows_equal(left: dict[str, Any], right: dict[str, Any], tol: float = 1e-9) -> bool:
    for field in ("intercept", "slope", "forecast_value", "forecast_max", "forecast_min"):
        left_value = left.get(field)
        right_value = right.get(field)
        if isinstance(left_value, (int, float)) and isinstance(right_value, (int, float)):
            if abs(float(left_value) - float(right_value)) > tol:
                return False
        elif left_value != right_value:
            return False
    return True


def extract_regression_rows(
    wb: xw.Book,
    meta: FileMeta,
    source_file: str,
) -> list[dict[str, Any]]:
    try:
        sheet = wb.sheets["Regression Model"]
    except Exception:
        print(f"Skipped regression extraction for {source_file}: missing 'Regression Model' sheet")
        return []

    snapshot = snapshot_sheet(sheet)
    try:
        anchor_row, anchor_col = find_anchor(snapshot, "max")
    except ValueError as exc:
        print(f"Skipped regression extraction for {source_file}: {exc}")
        return []
    anchor = (anchor_row, anchor_col)

    x_col = safe_col(anchor_col, -11)
    y_col = safe_col(anchor_col, -7)
    history_rows = collect_history_rows(sheet, start_row=anchor_row - 1, required_cols=[x_col, y_col])
    if len(history_rows) < 2:
        print(
            f"Skipped regression extraction for {source_file}: "
            "need at least 2 rows of x/y history"
        )
        return []

    min_label_cell = find_label_cell(snapshot, ["min"], anchor)
    forecast_total_cell = find_label_cell(
        snapshot,
        ["tot fcst w/o sa", "tot fcst wo sa", "total fcst w/o sa", "total forecast without sa"],
        anchor,
    )

    forecast_max = read_numeric_neighbor(sheet, anchor_row, anchor_col)
    forecast_min = (
        read_numeric_neighbor(sheet, min_label_cell[0], min_label_cell[1])
        if min_label_cell
        else None
    )
    range_width = (
        (forecast_max - forecast_min)
        if forecast_max is not None and forecast_min is not None
        else None
    )

    helper_intercept = sheet.cells(anchor_row + 2, anchor_col + 2)
    helper_slope = sheet.cells(anchor_row + 3, anchor_col + 2)

    rows: list[dict[str, Any]] = []
    max_quarters = min(10, len(history_rows))
    for num_quarters in range(2, max_quarters + 1):
        selected_rows = history_rows[-num_quarters:]
        start_row = selected_rows[0]
        end_row = selected_rows[-1]

        helper_intercept.formula2 = (
            f"=INTERCEPT(R{start_row}C{y_col}:R{end_row}C{y_col},"
            f"R{start_row}C{x_col}:R{end_row}C{x_col})"
        )
        helper_slope.formula2 = (
            f"=SLOPE(R{start_row}C{y_col}:R{end_row}C{y_col},"
            f"R{start_row}C{x_col}:R{end_row}C{x_col})"
        )
        wb.app.calculate()

        intercept = to_float(helper_intercept.value)
        slope = to_float(helper_slope.value)
        if intercept is None or slope is None:
            continue

        forecast_total_without_sa = (
            read_numeric_neighbor(sheet, forecast_total_cell[0], forecast_total_cell[1])
            if forecast_total_cell
            else None
        )
        if forecast_total_without_sa is None:
            latest_x = to_float(sheet.cells(end_row, x_col).value)
            if latest_x is not None:
                forecast_total_without_sa = intercept + (slope * latest_x)

        candidate = {
            "model": meta.model,
            "ticker": meta.ticker,
            "model_period": meta.model_period,
            "model_date": meta.model_date,
            "method": "regression",
            "parameter_name": "num_quarters_used",
            "parameter_value": num_quarters,
            "num_quarters_used": num_quarters,
            "forecast_value": forecast_total_without_sa,
            "actual_value": "",
            "forecast_max": forecast_max,
            "forecast_min": forecast_min,
            "range_width": range_width,
            "intercept": intercept,
            "slope": slope,
            "source_file": source_file,
        }
        if rows and regression_rows_equal(rows[-1], candidate):
            continue
        rows.append(candidate)

    helper_intercept.value = None
    helper_slope.value = None
    return rows


def close_workbook_safely(wb: xw.Book) -> None:
    try:
        wb.close(save=False)
        return
    except TypeError:
        pass
    except Exception:
        pass

    for closer in (
        lambda: wb.close(False),
        lambda: wb.api.Close(SaveChanges=False),
        lambda: wb.close(),
    ):
        try:
            closer()
            return
        except Exception:
            continue


def build_output_path(input_path: Path, output_path: Path) -> Path:
    folder_name = input_path.resolve().name
    base_candidate = output_path / f"{folder_name}_PARAM.xlsx"
    if not base_candidate.exists():
        return base_candidate
    suffix = 1
    while True:
        candidate = output_path / f"{folder_name}_PARAM.{suffix}.xlsx"
        if not candidate.exists():
            return candidate
        suffix += 1


def write_sheet(
    workbook: Workbook,
    sheet_name: str,
    columns: list[str],
    rows: list[dict[str, Any]],
) -> None:
    sheet = workbook.create_sheet(title=sheet_name)
    sheet.append(columns)
    for col_idx in range(1, len(columns) + 1):
        sheet.cell(row=1, column=col_idx).font = Font(bold=True)

    for row_dict in rows:
        sheet.append([row_dict.get(col, "") if row_dict.get(col) is not None else "" for col in columns])

    sheet.freeze_panes = "A2"
    last_col_letter = get_column_letter(len(columns))
    sheet.auto_filter.ref = f"A1:{last_col_letter}{max(1, sheet.max_row)}"

    for col_idx, col_name in enumerate(columns, start=1):
        col_letter = get_column_letter(col_idx)
        max_len = len(col_name)
        for cell in sheet[col_letter]:
            if cell.value is None:
                continue
            max_len = max(max_len, len(str(cell.value)))
        sheet.column_dimensions[col_letter].width = min(52, max(12, max_len + 2))


def write_output_workbook(
    output_file: Path,
    empirical_rows: list[dict[str, Any]],
    regression_rows: list[dict[str, Any]],
) -> None:
    workbook = Workbook()
    workbook.remove(workbook.active)
    write_sheet(workbook, "empirical_candidates", EMPIRICAL_COLUMNS, empirical_rows)
    write_sheet(workbook, "regression_candidates", REGRESSION_COLUMNS, regression_rows)
    workbook.save(output_file)


def list_source_files(input_path: Path) -> list[Path]:
    candidates: list[Path] = []
    for file_path in sorted(input_path.iterdir()):
        if not file_path.is_file():
            print(f"Skipped {file_path.name}: not a file")
            continue
        if file_path.name.startswith("~"):
            print(f"Skipped {file_path.name}: temporary Excel file")
            continue
        if file_path.suffix.lower() != ".xlsx":
            print(f"Skipped {file_path.name}: not an .xlsx file")
            continue
        if re.search(r"_PARAM(\.\d+)?\.xlsx$", file_path.name, re.IGNORECASE):
            print(f"Skipped {file_path.name}: output artifact")
            continue
        candidates.append(file_path)
    return candidates


def main() -> None:
    input_path = Path(input_dir).expanduser().resolve()
    output_path = Path(output_dir).expanduser().resolve()

    if not input_path.exists() or not input_path.is_dir():
        raise SystemExit(f"Input directory does not exist or is not a directory: {input_path}")
    output_path.mkdir(parents=True, exist_ok=True)

    source_files = list_source_files(input_path)
    empirical_rows: list[dict[str, Any]] = []
    regression_rows: list[dict[str, Any]] = []
    processed_files = 0

    if source_files:
        app = xw.App(visible=False, add_book=False)
        app.display_alerts = False
        app.screen_updating = False
        try:
            for file_path in source_files:
                print(f"Processing file: {file_path.name}")
                workbook = None
                try:
                    workbook = app.books.open(str(file_path), update_links=False)
                    meta = parse_file_meta(file_path.name)
                    try:
                        empirical_rows.extend(
                            extract_empirical_rows(workbook, meta=meta, source_file=file_path.name)
                        )
                    except Exception as exc:
                        print(f"Skipped empirical extraction for {file_path.name}: {exc}")
                    try:
                        regression_rows.extend(
                            extract_regression_rows(workbook, meta=meta, source_file=file_path.name)
                        )
                    except Exception as exc:
                        print(f"Skipped regression extraction for {file_path.name}: {exc}")
                    processed_files += 1
                except Exception as exc:
                    print(f"Skipped {file_path.name}: {exc}")
                finally:
                    if workbook is not None:
                        close_workbook_safely(workbook)
        finally:
            app.quit()
    else:
        print(f"No eligible .xlsx files found in {input_path}")

    output_file = build_output_path(input_path=input_path, output_path=output_path)
    write_output_workbook(output_file, empirical_rows, regression_rows)

    print(f"Output path: {output_file}")
    print(f"Number of files processed: {processed_files}")
    print(f"Number of empirical rows: {len(empirical_rows)}")
    print(f"Number of regression rows: {len(regression_rows)}")


if __name__ == "__main__":
    main()
