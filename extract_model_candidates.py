from __future__ import annotations

from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path
from typing import Any
import re

import pandas as pd
import xlwings as xw
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter


# -----------------------------
# User-editable IO configuration
# -----------------------------
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

N_QUARTERS = 10
PHASE_DAY = {"early": 5, "mid": 15, "late": 25}


@dataclass
class FileMeta:
    model: str
    ticker: str
    model_period: str
    model_date: str


def normalize_text(value: Any) -> str:
    if value is None:
        return ""
    text = str(value).strip().lower()
    text = text.replace("\n", " ").replace("_", " ")
    return re.sub(r"\s+", " ", text)


def is_blank(value: Any) -> bool:
    if value is None:
        return True
    if isinstance(value, str) and value.strip() == "":
        return True
    return False


def to_2d(values: Any) -> list[list[Any]]:
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
    if isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        return float(value)
    text = str(value).strip().replace(",", "")
    if text == "":
        return None
    try:
        return float(text)
    except ValueError:
        return None


def safe_subtract(left: Any, right: Any) -> float | None:
    l_num = to_float(left)
    r_num = to_float(right)
    if l_num is None or r_num is None:
        return None
    return l_num - r_num


def parse_file_meta(file_name: str) -> FileMeta:
    stem = Path(file_name).stem
    parts = [p.strip() for p in stem.split(" - ") if p.strip()]

    ticker = ""
    if len(parts) >= 2:
        ticker = parts[1].upper()
    else:
        ticker_match = re.search(r"\b([A-Z]{2,8})\b", stem)
        ticker = ticker_match.group(1) if ticker_match else "UNKNOWN"

    period_match = re.search(r"\b(Early|Mid|Late)([A-Za-z]{3,9})(\d{4})\b", stem, re.IGNORECASE)
    if period_match:
        phase = period_match.group(1).title()
        month_token = period_match.group(2)
        month_abbr = month_token[:3].title()
        year_int = int(period_match.group(3))
        model_period = f"{phase}{month_abbr}_{year_int}"
        day_int = PHASE_DAY[phase.lower()]
        month_int = datetime.strptime(month_abbr, "%b").month
        model_date = date(year_int, month_int, day_int).isoformat()
    else:
        model_period = "UNKNOWN_PERIOD"
        model_date = ""

    model = f"{ticker}_{model_period}"
    return FileMeta(model=model, ticker=ticker, model_period=model_period, model_date=model_date)


def resolve_output_path(in_dir: Path, out_dir: Path) -> Path:
    base = f"{in_dir.name}_PARAM"
    candidate = out_dir / f"{base}.xlsx"
    suffix_counter = 1
    while candidate.exists():
        candidate = out_dir / f"{base}.{suffix_counter}.xlsx"
        suffix_counter += 1
    return candidate


def close_workbook_without_save(wb: xw.Book) -> None:
    try:
        wb.close(save=False)
        return
    except TypeError:
        pass
    except Exception:
        pass

    try:
        wb.api.Close(False)  # COM/appscript fallback
        return
    except Exception:
        pass

    try:
        wb.close()
    except Exception:
        pass


def find_anchor(sheet: xw.Sheet, target: str = "max") -> tuple[int, int] | None:
    target_norm = normalize_text(target)

    # Fast path: native Excel Find.
    try:
        found = sheet.api.UsedRange.Find(What=target, MatchCase=False)
        if found is not None:
            return int(found.Row), int(found.Column)
    except Exception:
        pass

    # Safe fallback: single scan of used range values.
    used = sheet.used_range
    values_2d = to_2d(used.value)
    if not values_2d:
        return None

    for r_idx, row_values in enumerate(values_2d):
        for c_idx, cell_value in enumerate(row_values):
            if normalize_text(cell_value) == target_norm:
                return used.row + r_idx, used.column + c_idx
    return None


def read_header_offsets(sheet: xw.Sheet, header_row: int, anchor_col: int) -> dict[str, int]:
    used = sheet.used_range
    first_col = used.column
    last_col = used.column + used.columns.count - 1
    header_values = sheet.range((header_row, first_col), (header_row, last_col)).value
    header_row_values = header_values if isinstance(header_values, list) else [header_values]

    offsets: dict[str, int] = {}
    for idx, value in enumerate(header_row_values):
        norm = normalize_text(value)
        if norm:
            col = first_col + idx
            offsets[norm] = col - anchor_col
    return offsets


def pick_offset(header_offsets: dict[str, int], candidates: list[str], default: int | None = None) -> int | None:
    for header, offset in header_offsets.items():
        for candidate in candidates:
            if candidate in header:
                return offset
    return default


def detect_data_rows(sheet: xw.Sheet, start_row: int, check_cols: list[int], max_rows: int = 400) -> list[int]:
    valid_cols = sorted({col for col in check_cols if col >= 1})
    if not valid_cols:
        return []

    first_col = valid_cols[0]
    last_col = valid_cols[-1]
    end_row = start_row + max_rows - 1
    block = sheet.range((start_row, first_col), (end_row, last_col)).value
    block_2d = to_2d(block)

    rows: list[int] = []
    blank_streak = 0
    for idx, row_values in enumerate(block_2d):
        selected_values = [row_values[col - first_col] for col in valid_cols]
        if all(is_blank(v) for v in selected_values):
            blank_streak += 1
            if blank_streak >= 2:
                break
            continue
        blank_streak = 0
        rows.append(start_row + idx)
    return rows


def read_cell_by_offset(sheet: xw.Sheet, row: int, anchor_col: int, offset: int | None) -> Any:
    if offset is None:
        return None
    col = anchor_col + offset
    if col < 1:
        return None
    return sheet.cells(row, col).value


def create_empirical_rows(wb: xw.Book, meta: FileMeta, source_file: str) -> list[dict[str, Any]]:
    try:
        sheet = wb.sheets["Empirical Model"]
    except Exception:
        print(f"skipped {source_file}: missing sheet 'Empirical Model'")
        return []

    anchor = find_anchor(sheet, "max")
    if anchor is None:
        print(f"skipped {source_file}: could not find 'max' anchor in Empirical Model")
        return []
    anchor_row, anchor_col = anchor

    offsets = read_header_offsets(sheet, anchor_row, anchor_col)
    forecast_max_offset = 0
    forecast_min_offset = pick_offset(offsets, ["min"], default=1)
    forecast_value_offset = pick_offset(
        offsets,
        ["estimated total sold", "est total sold", "estimated total", "forecast", "tot fcst"],
        default=-1,
    )
    reported_sales_offset = pick_offset(offsets, ["reported sales", "actual"], default=2)
    quarterly_sales_offset = pick_offset(offsets, ["quarterly sales"], default=-6)
    growth_rate_offset = pick_offset(offsets, ["growth rate"], default=-4)
    sales_captured_offset = pick_offset(offsets, ["sales captured in db", "captured in db"], default=-3)
    penetration_offset = pick_offset(offsets, ["avg penetration", "penetration"], default=sales_captured_offset)
    last_quarter_offset = pick_offset(offsets, ["last quarter"], default=-8)

    row_check_cols = [
        anchor_col + forecast_max_offset,
        anchor_col + (forecast_min_offset or 0),
        anchor_col + (penetration_offset or 0),
    ]
    data_rows = detect_data_rows(sheet, anchor_row + 1, row_check_cols, max_rows=400)
    if not data_rows:
        print(f"skipped {source_file}: no empirical data rows detected")
        return []

    used = sheet.used_range
    scratch_col = used.column + used.columns.count + 3
    scratch_start_row = anchor_row + 1
    pen_col = anchor_col + penetration_offset if penetration_offset is not None else None

    formulas_written = False
    for n_quarters in range(1, N_QUARTERS + 1):
        scratch_row = scratch_start_row + n_quarters - 1
        scratch_cell = sheet.cells(scratch_row, scratch_col)
        scratch_cell.clear_contents()
        if pen_col is None or pen_col < 1 or len(data_rows) < n_quarters:
            continue
        start_row = data_rows[-n_quarters]
        end_row = data_rows[-1]
        scratch_cell.formula2 = f"=AVERAGE(R{start_row}C{pen_col}:R{end_row}C{pen_col})"
        formulas_written = True

    if formulas_written:
        wb.app.calculate()

    avg_pen_values = sheet.range(
        (scratch_start_row, scratch_col),
        (scratch_start_row + N_QUARTERS - 1, scratch_col),
    ).value
    avg_pen_list = avg_pen_values if isinstance(avg_pen_values, list) else [avg_pen_values]

    rows: list[dict[str, Any]] = []
    for n_quarters in range(1, N_QUARTERS + 1):
        if n_quarters > len(data_rows):
            continue
        source_row = data_rows[n_quarters - 1]

        forecast_max = read_cell_by_offset(sheet, source_row, anchor_col, forecast_max_offset)
        forecast_min = read_cell_by_offset(sheet, source_row, anchor_col, forecast_min_offset)
        forecast_value = read_cell_by_offset(sheet, source_row, anchor_col, forecast_value_offset)
        reported_sales = read_cell_by_offset(sheet, source_row, anchor_col, reported_sales_offset)
        quarterly_sales = read_cell_by_offset(sheet, source_row, anchor_col, quarterly_sales_offset)
        growth_rate_pct = read_cell_by_offset(sheet, source_row, anchor_col, growth_rate_offset)
        sales_captured_in_db_pct = read_cell_by_offset(sheet, source_row, anchor_col, sales_captured_offset)
        last_quarter_used = read_cell_by_offset(sheet, source_row, anchor_col, last_quarter_offset)

        avg_penetration_pct = avg_pen_list[n_quarters - 1] if n_quarters - 1 < len(avg_pen_list) else None

        if all(
            is_blank(value)
            for value in (
                forecast_value,
                forecast_max,
                forecast_min,
                reported_sales,
                avg_penetration_pct,
            )
        ):
            continue

        rows.append(
            {
                "model": meta.model,
                "ticker": meta.ticker,
                "model_period": meta.model_period,
                "model_date": meta.model_date,
                "method": "empirical",
                "parameter_name": "avg_penetration_pct",
                "parameter_value": avg_penetration_pct,
                "num_quarters_used": n_quarters,
                "last_quarter_used": last_quarter_used,
                "forecast_value": forecast_value,
                "actual_value": reported_sales,
                "forecast_max": forecast_max,
                "forecast_min": forecast_min,
                "range_width": safe_subtract(forecast_max, forecast_min),
                "avg_penetration_pct": avg_penetration_pct,
                "quarterly_sales": quarterly_sales,
                "reported_sales": reported_sales,
                "growth_rate_pct": growth_rate_pct,
                "sales_captured_in_db_pct": sales_captured_in_db_pct,
                "source_file": source_file,
            }
        )

    # Keep temporary formulas temporary.
    sheet.range(
        (scratch_start_row, scratch_col),
        (scratch_start_row + N_QUARTERS - 1, scratch_col),
    ).clear_contents()

    return rows


def signature_value(value: Any) -> Any:
    number = to_float(value)
    if number is not None:
        return round(number, 10)
    if isinstance(value, str):
        return value.strip()
    return value


def create_regression_rows(wb: xw.Book, meta: FileMeta, source_file: str) -> list[dict[str, Any]]:
    try:
        sheet = wb.sheets["Regression Model"]
    except Exception:
        print(f"skipped {source_file}: missing sheet 'Regression Model'")
        return []

    anchor = find_anchor(sheet, "max")
    if anchor is None:
        print(f"skipped {source_file}: could not find 'max' anchor in Regression Model")
        return []
    anchor_row, anchor_col = anchor

    # Required fixed offsets from the existing model logic.
    y_col = anchor_col - 7
    x_col = anchor_col - 11

    offsets = read_header_offsets(sheet, anchor_row, anchor_col)
    forecast_max_offset = 0
    forecast_min_offset = pick_offset(offsets, ["min"], default=1)
    forecast_value_offset = pick_offset(
        offsets,
        ["tot fcst w/o sa", "total fcst w/o sa", "tot fcst", "forecast"],
        default=-1,
    )
    num_quarters_offset = pick_offset(offsets, ["num quarters"], default=None)
    actual_offset = pick_offset(offsets, ["actual", "reported sales"], default=None)

    data_rows = detect_data_rows(sheet, anchor_row + 1, [x_col, y_col], max_rows=400)
    if not data_rows:
        print(f"skipped {source_file}: no regression data rows detected")
        return []

    used = sheet.used_range
    scratch_intercept_col = used.column + used.columns.count + 3
    scratch_slope_col = scratch_intercept_col + 1
    scratch_start_row = anchor_row + 1

    formulas_written = False
    for n_quarters in range(1, N_QUARTERS + 1):
        intercept_cell = sheet.cells(scratch_start_row + n_quarters - 1, scratch_intercept_col)
        slope_cell = sheet.cells(scratch_start_row + n_quarters - 1, scratch_slope_col)
        intercept_cell.clear_contents()
        slope_cell.clear_contents()

        if len(data_rows) < n_quarters:
            continue
        start_row = data_rows[-n_quarters]
        end_row = data_rows[-1]

        intercept_cell.formula2 = (
            f"=INTERCEPT(R{start_row}C{y_col}:R{end_row}C{y_col},R{start_row}C{x_col}:R{end_row}C{x_col})"
        )
        slope_cell.formula2 = f"=SLOPE(R{start_row}C{y_col}:R{end_row}C{y_col},R{start_row}C{x_col}:R{end_row}C{x_col})"
        formulas_written = True

    if formulas_written:
        wb.app.calculate()

    intercept_values = sheet.range(
        (scratch_start_row, scratch_intercept_col),
        (scratch_start_row + N_QUARTERS - 1, scratch_intercept_col),
    ).value
    slope_values = sheet.range(
        (scratch_start_row, scratch_slope_col),
        (scratch_start_row + N_QUARTERS - 1, scratch_slope_col),
    ).value

    intercept_list = intercept_values if isinstance(intercept_values, list) else [intercept_values]
    slope_list = slope_values if isinstance(slope_values, list) else [slope_values]

    rows: list[dict[str, Any]] = []
    previous_signature: tuple[Any, ...] | None = None
    for n_quarters in range(1, N_QUARTERS + 1):
        if n_quarters > len(data_rows):
            continue
        source_row = data_rows[n_quarters - 1]
        num_quarters_used = (
            read_cell_by_offset(sheet, source_row, anchor_col, num_quarters_offset)
            if num_quarters_offset is not None
            else n_quarters
        )

        forecast_value = read_cell_by_offset(sheet, source_row, anchor_col, forecast_value_offset)
        actual_value = read_cell_by_offset(sheet, source_row, anchor_col, actual_offset)
        forecast_max = read_cell_by_offset(sheet, source_row, anchor_col, forecast_max_offset)
        forecast_min = read_cell_by_offset(sheet, source_row, anchor_col, forecast_min_offset)
        intercept = intercept_list[n_quarters - 1] if n_quarters - 1 < len(intercept_list) else None
        slope = slope_list[n_quarters - 1] if n_quarters - 1 < len(slope_list) else None

        if all(is_blank(v) for v in (forecast_value, forecast_max, forecast_min, intercept, slope)):
            continue

        row_signature = (
            signature_value(forecast_value),
            signature_value(forecast_max),
            signature_value(forecast_min),
            signature_value(intercept),
            signature_value(slope),
        )
        if row_signature == previous_signature:
            continue
        previous_signature = row_signature

        rows.append(
            {
                "model": meta.model,
                "ticker": meta.ticker,
                "model_period": meta.model_period,
                "model_date": meta.model_date,
                "method": "regression",
                "parameter_name": "num_quarters_used",
                "parameter_value": num_quarters_used,
                "num_quarters_used": num_quarters_used,
                "forecast_value": forecast_value,
                "actual_value": actual_value,
                "forecast_max": forecast_max,
                "forecast_min": forecast_min,
                "range_width": safe_subtract(forecast_max, forecast_min),
                "intercept": intercept,
                "slope": slope,
                "source_file": source_file,
            }
        )

    # Keep temporary formulas temporary.
    sheet.range(
        (scratch_start_row, scratch_intercept_col),
        (scratch_start_row + N_QUARTERS - 1, scratch_slope_col),
    ).clear_contents()

    return rows


def format_output_sheet(worksheet: Any) -> None:
    worksheet.freeze_panes = "A2"
    worksheet.auto_filter.ref = worksheet.dimensions
    for cell in worksheet[1]:
        cell.font = Font(bold=True)

    for col_idx in range(1, worksheet.max_column + 1):
        max_len = 0
        for row_idx in range(1, worksheet.max_row + 1):
            value = worksheet.cell(row=row_idx, column=col_idx).value
            text = "" if value is None else str(value)
            if len(text) > max_len:
                max_len = len(text)
        worksheet.column_dimensions[get_column_letter(col_idx)].width = min(max(max_len + 2, 12), 48)


def write_output_workbook(output_path: Path, empirical_rows: list[dict[str, Any]], regression_rows: list[dict[str, Any]]) -> None:
    empirical_df = pd.DataFrame(empirical_rows).reindex(columns=EMPIRICAL_COLUMNS)
    regression_df = pd.DataFrame(regression_rows).reindex(columns=REGRESSION_COLUMNS)

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        empirical_df.to_excel(writer, sheet_name="empirical_candidates", index=False)
        regression_df.to_excel(writer, sheet_name="regression_candidates", index=False)
        format_output_sheet(writer.book["empirical_candidates"])
        format_output_sheet(writer.book["regression_candidates"])


def collect_xlsx_candidates(in_dir: Path) -> tuple[list[Path], list[str]]:
    candidates: list[Path] = []
    skip_messages: list[str] = []
    for item in sorted(in_dir.iterdir()):
        if not item.is_file():
            skip_messages.append(f"skipped {item.name}: not a file")
            continue
        if item.name.startswith("~"):
            skip_messages.append(f"skipped {item.name}: temporary file")
            continue
        if item.suffix.lower() != ".xlsx":
            skip_messages.append(f"skipped {item.name}: not .xlsx")
            continue
        candidates.append(item)
    return candidates, skip_messages


def main() -> None:
    in_dir = Path(input_dir).expanduser().resolve()
    out_dir = Path(output_dir).expanduser().resolve()
    out_dir.mkdir(parents=True, exist_ok=True)

    if not in_dir.exists():
        print(f"input directory not found: {in_dir}")
        return
    if not in_dir.is_dir():
        print(f"input path is not a directory: {in_dir}")
        return

    output_path = resolve_output_path(in_dir, out_dir)

    empirical_rows: list[dict[str, Any]] = []
    regression_rows: list[dict[str, Any]] = []
    processed_file_count = 0

    xlsx_files, skip_messages = collect_xlsx_candidates(in_dir)
    for message in skip_messages:
        print(message)

    app = xw.App(visible=False, add_book=False)
    app.display_alerts = False
    app.screen_updating = False
    try:
        try:
            app.enable_events = False
        except Exception:
            pass
        try:
            app.calculation = "manual"
        except Exception:
            pass

        for file_path in xlsx_files:
            wb: xw.Book | None = None
            try:
                wb = app.books.open(str(file_path), update_links=False)
                meta = parse_file_meta(file_path.name)
                empirical_rows.extend(create_empirical_rows(wb, meta, file_path.name))
                regression_rows.extend(create_regression_rows(wb, meta, file_path.name))
                processed_file_count += 1
                print(f"processed {file_path.name}")
            except Exception as exc:
                print(f"skipped {file_path.name}: {exc}")
            finally:
                if wb is not None:
                    close_workbook_without_save(wb)
    finally:
        app.quit()

    write_output_workbook(output_path, empirical_rows, regression_rows)

    print(f"output path: {output_path}")
    print(f"number of files processed: {processed_file_count}")
    print(f"number of empirical rows: {len(empirical_rows)}")
    print(f"number of regression rows: {len(regression_rows)}")


if __name__ == "__main__":
    main()
