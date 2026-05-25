from __future__ import annotations

import re
from datetime import date, datetime
from pathlib import Path
from typing import Any, Iterable

import xlwings as xw
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

# User-configurable paths.
input_dir = "/workspace/input"
output_dir = "/workspace/output"


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

# Anchor-relative offsets used to read candidate values quickly from each table row.
EMPIRICAL_OFFSETS = {
    "num_quarters_used": -8,
    "last_quarter_used": -7,
    "quarterly_sales": -6,
    "reported_sales": -5,
    "growth_rate_pct": -4,
    "sales_captured_in_db_pct": -3,
    "forecast_value": -2,  # estimated total sold
    "actual_value": -1,  # reported sales
    "forecast_max": 0,
    "forecast_min": 1,
}

REGRESSION_OFFSETS = {
    "forecast_value": -2,  # TOT FCST w/o SA
    "actual_value": -1,  # optional
    "forecast_max": 0,
    "forecast_min": 1,
}

PERIOD_DAY = {"Early": 5, "Mid": 15, "Late": 25}
MODEL_PERIOD_RE = re.compile(r"(Early|Mid|Late)([A-Za-z]{3})(\d{4})", re.IGNORECASE)


def to_matrix(values: Any) -> list[list[Any]]:
    if values is None:
        return []
    if isinstance(values, list):
        if not values:
            return []
        if isinstance(values[0], list):
            return values
        return [values]
    return [[values]]


def is_blank(value: Any) -> bool:
    if value is None:
        return True
    if isinstance(value, str) and value.strip() == "":
        return True
    return False


def to_float(value: Any) -> float | None:
    if isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        return float(value)
    if isinstance(value, str):
        text = value.strip().replace(",", "").replace("%", "")
        if text == "":
            return None
        try:
            return float(text)
        except ValueError:
            return None
    return None


def clean_output_value(value: Any) -> Any:
    if isinstance(value, datetime):
        return value.strftime("%Y-%m-%d")
    if isinstance(value, date):
        return value.isoformat()
    return value


def parse_file_label(file_path: Path) -> dict[str, str] | None:
    stem = file_path.stem
    pieces = [part.strip() for part in stem.split(" - ")]
    if len(pieces) < 3:
        return None

    ticker = pieces[1].strip().upper()
    model_piece = pieces[2]
    match = MODEL_PERIOD_RE.search(model_piece)
    if not match:
        return None

    period_prefix = match.group(1).title()
    month_abbr = match.group(2).title()
    year_text = match.group(3)
    day = PERIOD_DAY.get(period_prefix)
    if day is None:
        return None

    try:
        month_num = datetime.strptime(month_abbr, "%b").month
        model_date = date(int(year_text), month_num, day).isoformat()
    except ValueError:
        return None

    model_period = f"{period_prefix}{month_abbr}_{year_text}"
    model = f"{ticker}_{model_period}"
    return {
        "ticker": ticker,
        "model_period": model_period,
        "model_date": model_date,
        "model": model,
    }


def output_workbook_path(input_path: Path, output_path: Path) -> Path:
    output_path.mkdir(parents=True, exist_ok=True)
    folder_name = input_path.name
    base = f"{folder_name}_PARAM"
    candidate = output_path / f"{base}.xlsx"
    idx = 1
    while candidate.exists():
        candidate = output_path / f"{base}.{idx}.xlsx"
        idx += 1
    return candidate


def find_anchor_max(sheet: xw.Sheet) -> tuple[int, int] | None:
    try:
        found = sheet.api.Cells.Find(What="max", LookAt=1, MatchCase=False)
        if found is not None:
            return int(found.Row), int(found.Column)
    except Exception:
        pass

    used = sheet.used_range
    values = to_matrix(used.value)
    if not values:
        return None

    top_row = used.row
    left_col = used.column
    for r_idx, row in enumerate(values):
        for c_idx, cell_value in enumerate(row):
            if isinstance(cell_value, str) and cell_value.strip().lower() == "max":
                return top_row + r_idx, left_col + c_idx
    return None


def close_source_workbook(wb: xw.Book) -> None:
    try:
        wb.close(save=False)
        return
    except TypeError:
        # Older xlwings versions may not support close(save=False).
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


def regression_duplicate_key(row: dict[str, Any]) -> tuple[Any, ...]:
    def norm(num: Any) -> Any:
        f = to_float(num)
        if f is None:
            return None
        return round(f, 10)

    return (
        row.get("num_quarters_used"),
        norm(row.get("intercept")),
        norm(row.get("slope")),
        norm(row.get("forecast_value")),
        norm(row.get("forecast_max")),
        norm(row.get("forecast_min")),
    )


def detect_numeric_history_start(
    sheet: xw.Sheet,
    x_col: int,
    y_col: int,
    history_last_row: int,
    max_lookback: int = 80,
) -> int:
    start_row = history_last_row
    lookback = 0
    while start_row >= 1 and lookback < max_lookback:
        x_val = to_float(sheet.range((start_row, x_col)).value)
        y_val = to_float(sheet.range((start_row, y_col)).value)
        if x_val is None or y_val is None:
            break
        start_row -= 1
        lookback += 1
    return start_row + 1


def process_empirical_sheet(
    wb: xw.Book,
    file_meta: dict[str, str],
    source_file_name: str,
    n_quarters: int = 10,
) -> list[dict[str, Any]]:
    try:
        sheet = wb.sheets["Empirical Model"]
    except Exception:
        print(f"  skipped empirical: sheet 'Empirical Model' not found in {source_file_name}")
        return []

    anchor = find_anchor_max(sheet)
    if anchor is None:
        print(f"  skipped empirical: 'max' anchor not found in {source_file_name}")
        return []
    anchor_row, anchor_col = anchor

    row_start = anchor_row + 1
    row_end = anchor_row + n_quarters
    min_off = min(EMPIRICAL_OFFSETS.values())
    max_off = max(EMPIRICAL_OFFSETS.values())
    col_start = anchor_col + min_off
    col_end = anchor_col + max_off

    # Write all average penetration formulas at once and calculate once.
    history_last_row = anchor_row - 1
    penetration_col = anchor_col + EMPIRICAL_OFFSETS["sales_captured_in_db_pct"]
    temp_col = max(sheet.used_range.last_cell.column + 1, anchor_col + 20)
    formula_rows: list[list[str]] = []
    for i in range(1, n_quarters + 1):
        hist_start = max(1, history_last_row - i + 1)
        formula_rows.append(
            [
                f'=IFERROR(AVERAGE(R{hist_start}C{penetration_col}:R{history_last_row}C{penetration_col}),"")'
            ]
        )

    temp_rng = sheet.range((row_start, temp_col), (row_end, temp_col))
    temp_rng.formula2 = formula_rows
    wb.app.calculate()
    avg_pen_values = [row[0] if row else None for row in to_matrix(temp_rng.value)]
    temp_rng.clear_contents()

    table_values = to_matrix(sheet.range((row_start, col_start), (row_end, col_end)).value)
    rows: list[dict[str, Any]] = []
    for idx, table_row in enumerate(table_values):
        row_num = row_start + idx

        def get_from_offset(offset_key: str) -> Any:
            src_col = anchor_col + EMPIRICAL_OFFSETS[offset_key]
            value_idx = src_col - col_start
            if value_idx < 0 or value_idx >= len(table_row):
                return None
            return table_row[value_idx]

        num_quarters_used_raw = get_from_offset("num_quarters_used")
        num_quarters_used = int(to_float(num_quarters_used_raw) or (idx + 1))
        forecast_max = to_float(get_from_offset("forecast_max"))
        forecast_min = to_float(get_from_offset("forecast_min"))
        forecast_value = to_float(get_from_offset("forecast_value"))
        actual_value = to_float(get_from_offset("actual_value"))
        avg_penetration = to_float(avg_pen_values[idx] if idx < len(avg_pen_values) else None)

        if all(v is None for v in (forecast_max, forecast_min, forecast_value, avg_penetration)):
            continue

        range_width = (
            forecast_max - forecast_min
            if forecast_max is not None and forecast_min is not None
            else None
        )

        rows.append(
            {
                "model": file_meta["model"],
                "ticker": file_meta["ticker"],
                "model_period": file_meta["model_period"],
                "model_date": file_meta["model_date"],
                "method": "empirical",
                "parameter_name": "avg_penetration_pct",
                "parameter_value": avg_penetration,
                "num_quarters_used": num_quarters_used,
                "last_quarter_used": clean_output_value(get_from_offset("last_quarter_used")),
                "forecast_value": forecast_value,  # estimated total sold
                "actual_value": actual_value,  # reported sales
                "forecast_max": forecast_max,
                "forecast_min": forecast_min,
                "range_width": range_width,
                "avg_penetration_pct": avg_penetration,
                "quarterly_sales": to_float(get_from_offset("quarterly_sales")),
                "reported_sales": to_float(get_from_offset("reported_sales")),
                "growth_rate_pct": to_float(get_from_offset("growth_rate_pct")),
                "sales_captured_in_db_pct": to_float(get_from_offset("sales_captured_in_db_pct")),
                "source_file": source_file_name,
            }
        )

    return rows


def process_regression_sheet(
    wb: xw.Book,
    file_meta: dict[str, str],
    source_file_name: str,
    max_quarters: int = 10,
) -> list[dict[str, Any]]:
    try:
        sheet = wb.sheets["Regression Model"]
    except Exception:
        print(f"  skipped regression: sheet 'Regression Model' not found in {source_file_name}")
        return []

    anchor = find_anchor_max(sheet)
    if anchor is None:
        print(f"  skipped regression: 'max' anchor not found in {source_file_name}")
        return []
    anchor_row, anchor_col = anchor

    y_col = anchor_col - 7
    x_col = anchor_col - 11
    history_last_row = anchor_row - 1
    history_start_row = detect_numeric_history_start(sheet, x_col, y_col, history_last_row)
    available = max(0, history_last_row - history_start_row + 1)
    n_list = list(range(2, min(max_quarters, available) + 1))
    if not n_list:
        return []

    out_row_start = anchor_row + 1
    out_row_end = out_row_start + len(n_list) - 1
    temp_base_col = max(sheet.used_range.last_cell.column + 1, anchor_col + 20)
    intercept_col = temp_base_col
    slope_col = temp_base_col + 1

    intercept_formulas: list[list[str]] = []
    slope_formulas: list[list[str]] = []
    for n in n_list:
        hist_start = history_last_row - n + 1
        intercept_formulas.append(
            [
                f'=IFERROR(INTERCEPT(R{hist_start}C{y_col}:R{history_last_row}C{y_col},R{hist_start}C{x_col}:R{history_last_row}C{x_col}),"")'
            ]
        )
        slope_formulas.append(
            [
                f'=IFERROR(SLOPE(R{hist_start}C{y_col}:R{history_last_row}C{y_col},R{hist_start}C{x_col}:R{history_last_row}C{x_col}),"")'
            ]
        )

    intercept_rng = sheet.range((out_row_start, intercept_col), (out_row_end, intercept_col))
    slope_rng = sheet.range((out_row_start, slope_col), (out_row_end, slope_col))
    intercept_rng.formula2 = intercept_formulas
    slope_rng.formula2 = slope_formulas
    wb.app.calculate()
    intercept_values = [row[0] if row else None for row in to_matrix(intercept_rng.value)]
    slope_values = [row[0] if row else None for row in to_matrix(slope_rng.value)]
    intercept_rng.clear_contents()
    slope_rng.clear_contents()

    support_col_start = anchor_col + min(REGRESSION_OFFSETS.values())
    support_col_end = anchor_col + max(REGRESSION_OFFSETS.values())
    support_values = to_matrix(
        sheet.range((out_row_start, support_col_start), (out_row_end, support_col_end)).value
    )
    next_x = to_float(sheet.range((history_last_row + 1, x_col)).value)

    rows: list[dict[str, Any]] = []
    prev_key: tuple[Any, ...] | None = None
    for idx, n in enumerate(n_list):
        table_row = support_values[idx] if idx < len(support_values) else []

        def support_value(offset_key: str) -> Any:
            src_col = anchor_col + REGRESSION_OFFSETS[offset_key]
            value_idx = src_col - support_col_start
            if value_idx < 0 or value_idx >= len(table_row):
                return None
            return table_row[value_idx]

        intercept = to_float(intercept_values[idx] if idx < len(intercept_values) else None)
        slope = to_float(slope_values[idx] if idx < len(slope_values) else None)
        forecast_value = to_float(support_value("forecast_value"))
        if forecast_value is None and intercept is not None and slope is not None and next_x is not None:
            forecast_value = intercept + (slope * next_x)

        forecast_max = to_float(support_value("forecast_max"))
        forecast_min = to_float(support_value("forecast_min"))
        actual_value = to_float(support_value("actual_value"))
        range_width = (
            forecast_max - forecast_min
            if forecast_max is not None and forecast_min is not None
            else None
        )

        row = {
            "model": file_meta["model"],
            "ticker": file_meta["ticker"],
            "model_period": file_meta["model_period"],
            "model_date": file_meta["model_date"],
            "method": "regression",
            "parameter_name": "num_quarters_used",
            "parameter_value": n,
            "num_quarters_used": n,
            "forecast_value": forecast_value,  # TOT FCST w/o SA
            "actual_value": actual_value if actual_value is not None else None,
            "forecast_max": forecast_max,
            "forecast_min": forecast_min,
            "range_width": range_width,
            "intercept": intercept,
            "slope": slope,
            "source_file": source_file_name,
        }

        if all(
            value is None
            for value in (
                row["intercept"],
                row["slope"],
                row["forecast_value"],
                row["forecast_max"],
                row["forecast_min"],
            )
        ):
            continue

        current_key = regression_duplicate_key(row)
        if prev_key is not None and current_key == prev_key:
            continue
        prev_key = current_key
        rows.append(row)

    return rows


def write_sheet(
    workbook: Workbook,
    sheet_name: str,
    columns: list[str],
    rows: Iterable[dict[str, Any]],
) -> None:
    ws = workbook.create_sheet(title=sheet_name)
    ws.append(columns)
    for col_idx in range(1, len(columns) + 1):
        ws.cell(row=1, column=col_idx).font = Font(bold=True)

    for row in rows:
        ws.append([clean_output_value(row.get(col)) for col in columns])

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(ws.max_column)}{ws.max_row}"

    for col_idx, col_name in enumerate(columns, start=1):
        max_len = len(col_name)
        for cell in ws.iter_cols(min_col=col_idx, max_col=col_idx, min_row=2, max_row=ws.max_row):
            for item in cell:
                if item.value is None:
                    continue
                max_len = max(max_len, len(str(item.value)))
        ws.column_dimensions[get_column_letter(col_idx)].width = min(max(12, max_len + 2), 42)


def create_output_workbook(
    output_path: Path,
    empirical_rows: list[dict[str, Any]],
    regression_rows: list[dict[str, Any]],
) -> None:
    out_wb = Workbook()
    out_wb.remove(out_wb.active)
    write_sheet(out_wb, "empirical_candidates", EMPIRICAL_COLUMNS, empirical_rows)
    write_sheet(out_wb, "regression_candidates", REGRESSION_COLUMNS, regression_rows)
    out_wb.save(output_path)


def iter_source_files(folder: Path) -> Iterable[Path]:
    for file_path in sorted(folder.iterdir()):
        if not file_path.is_file():
            continue
        if file_path.name.startswith("~"):
            print(f"SKIPPED {file_path.name}: temp file")
            continue
        if file_path.suffix.lower() != ".xlsx":
            print(f"SKIPPED {file_path.name}: not an .xlsx file")
            continue
        if re.search(r"_PARAM(\.\d+)?\.xlsx$", file_path.name, re.IGNORECASE):
            print(f"SKIPPED {file_path.name}: output workbook pattern")
            continue
        yield file_path


def main() -> None:
    input_path = Path(input_dir).expanduser().resolve()
    output_path = Path(output_dir).expanduser().resolve()

    if not input_path.exists() or not input_path.is_dir():
        raise FileNotFoundError(f"input_dir does not exist or is not a directory: {input_path}")

    sources = list(iter_source_files(input_path))
    output_file = output_workbook_path(input_path, output_path)

    empirical_rows: list[dict[str, Any]] = []
    regression_rows: list[dict[str, Any]] = []
    processed_files = 0

    app = xw.App(visible=False, add_book=False)
    try:
        app.display_alerts = False
        app.screen_updating = False
        try:
            app.calculation = "manual"
        except Exception:
            pass

        for file_path in sources:
            file_meta = parse_file_label(file_path)
            if file_meta is None:
                print(f"SKIPPED {file_path.name}: could not parse ticker/model period/date")
                continue

            print(f"PROCESSING {file_path.name}")
            wb = None
            try:
                wb = app.books.open(str(file_path), update_links=False)
                empirical_rows.extend(
                    process_empirical_sheet(
                        wb=wb,
                        file_meta=file_meta,
                        source_file_name=file_path.name,
                        n_quarters=10,
                    )
                )
                regression_rows.extend(
                    process_regression_sheet(
                        wb=wb,
                        file_meta=file_meta,
                        source_file_name=file_path.name,
                        max_quarters=10,
                    )
                )
                processed_files += 1
            except Exception as exc:
                print(f"SKIPPED {file_path.name}: processing error ({exc})")
            finally:
                if wb is not None:
                    close_source_workbook(wb)
    finally:
        try:
            app.quit()
        except Exception:
            pass

    create_output_workbook(output_file, empirical_rows, regression_rows)

    print(f"OUTPUT {output_file}")
    print(f"FILES_PROCESSED {processed_files}")
    print(f"EMPIRICAL_ROWS {len(empirical_rows)}")
    print(f"REGRESSION_ROWS {len(regression_rows)}")


if __name__ == "__main__":
    main()
