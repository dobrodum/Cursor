from __future__ import annotations

import re
from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path
from typing import Any

import xlwings as xw
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

# ---- Configure paths here ---------------------------------------------------
input_dir = Path("./input")
output_dir = Path("./output")
# -----------------------------------------------------------------------------

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


@dataclass
class FileLabel:
    model: str
    ticker: str
    model_period: str
    model_date: str


def normalize_text(value: Any) -> str:
    if value is None:
        return ""
    text = str(value).strip().lower()
    text = re.sub(r"[^a-z0-9]+", " ", text)
    return " ".join(text.split())


def flatten_single_column(values: Any) -> list[Any]:
    if isinstance(values, list):
        if values and isinstance(values[0], list):
            return [row[0] if row else None for row in values]
        return values
    return [values]


def as_2d(values: Any) -> list[list[Any]]:
    if isinstance(values, list):
        if values and isinstance(values[0], list):
            return values
        return [values]
    return [[values]]


def to_float(value: Any) -> float | None:
    if value is None or value == "":
        return None
    if isinstance(value, bool):
        return float(value)
    if isinstance(value, (int, float)):
        return float(value)
    if isinstance(value, str):
        cleaned = value.strip().replace(",", "")
        if cleaned.endswith("%"):
            try:
                return float(cleaned[:-1]) / 100.0
            except ValueError:
                return None
        try:
            return float(cleaned)
        except ValueError:
            return None
    return None


def to_int(value: Any) -> int | None:
    number = to_float(value)
    if number is None:
        return None
    rounded = int(round(number))
    if abs(number - rounded) > 1e-6:
        return None
    return rounded


def safe_subtract(left: float | None, right: float | None) -> float | None:
    if left is None or right is None:
        return None
    return left - right


def round_signature(value: float | None, digits: int = 8) -> float | None:
    if value is None:
        return None
    return round(value, digits)


def parse_file_label(file_name: str) -> FileLabel:
    stem = Path(file_name).stem
    parts = [part.strip() for part in stem.split(" - ")]
    if len(parts) < 3:
        raise ValueError("filename does not match expected '... - TICKER - Period_Send.xlsx' format")

    ticker = parts[1]
    period_token = re.sub(r"_send.*$", "", parts[2], flags=re.IGNORECASE)
    period_match = re.search(r"(Early|Mid|Late)([A-Za-z]{3,9})(\d{4})", period_token, flags=re.IGNORECASE)
    if not period_match:
        raise ValueError("could not parse model period token")

    period_prefix_raw, month_raw, year_raw = period_match.groups()
    period_prefix = period_prefix_raw.capitalize()
    year_int = int(year_raw)

    month_candidate = month_raw[:3].title()
    try:
        month_int = datetime.strptime(month_candidate, "%b").month
    except ValueError:
        month_int = datetime.strptime(month_raw.title(), "%B").month
    month_abbrev = datetime(year_int, month_int, 1).strftime("%b")

    day_map = {"Early": 5, "Mid": 15, "Late": 25}
    day = day_map[period_prefix]
    model_period = f"{period_prefix}{month_abbrev}_{year_int}"
    model_date = date(year_int, month_int, day).isoformat()
    model = f"{ticker}_{model_period}"

    return FileLabel(
        model=model,
        ticker=ticker,
        model_period=model_period,
        model_date=model_date,
    )


def unique_output_path(in_dir: Path, out_dir: Path) -> Path:
    base_name = f"{in_dir.name}_PARAM.xlsx"
    candidate = out_dir / base_name
    if not candidate.exists():
        return candidate

    index = 1
    while True:
        candidate = out_dir / f"{in_dir.name}_PARAM.{index}.xlsx"
        if not candidate.exists():
            return candidate
        index += 1


def find_sheet_case_insensitive(wb: xw.Book, name: str) -> xw.Sheet | None:
    target = name.strip().lower()
    for sheet in wb.sheets:
        if sheet.name.strip().lower() == target:
            return sheet
    return None


def find_max_anchor(sheet: xw.Sheet) -> tuple[int, int, xw.Range]:
    used = sheet.used_range
    values = as_2d(used.value)
    start_row = used.row
    start_col = used.column

    for r_idx, row_values in enumerate(values):
        for c_idx, cell_value in enumerate(row_values):
            if normalize_text(cell_value) == "max":
                return start_row + r_idx, start_col + c_idx, used

    raise ValueError("could not find 'max' anchor cell")


def read_header_map(sheet: xw.Sheet, row: int, first_col: int, last_col: int) -> dict[int, str]:
    values = sheet.range((row, first_col), (row, last_col)).value
    row_values = flatten_single_column(values)
    headers: dict[int, str] = {}
    for offset, value in enumerate(row_values):
        headers[first_col + offset] = normalize_text(value)
    return headers


def pick_col(headers: dict[int, str], include_phrases: list[str], exclude_phrases: list[str] | None = None) -> int | None:
    exclude_phrases = exclude_phrases or []
    for col, text in headers.items():
        if not text:
            continue
        if any(phrase in text for phrase in include_phrases) and not any(
            phrase in text for phrase in exclude_phrases
        ):
            return col
    return None


def collect_numeric_rows(sheet: xw.Sheet, col: int, first_row: int, last_row: int) -> list[int]:
    if col < 1 or last_row < first_row:
        return []
    values = sheet.range((first_row, col), (last_row, col)).value
    data = flatten_single_column(values)
    rows: list[int] = []
    for idx, value in enumerate(data):
        if to_float(value) is not None:
            rows.append(first_row + idx)
    return rows


def collect_numeric_pair_rows(
    sheet: xw.Sheet,
    col_a: int,
    col_b: int,
    first_row: int,
    last_row: int,
) -> list[int]:
    if col_a < 1 or col_b < 1 or last_row < first_row:
        return []
    values_a = flatten_single_column(sheet.range((first_row, col_a), (last_row, col_a)).value)
    values_b = flatten_single_column(sheet.range((first_row, col_b), (last_row, col_b)).value)
    rows: list[int] = []
    for idx, (value_a, value_b) in enumerate(zip(values_a, values_b)):
        if to_float(value_a) is not None and to_float(value_b) is not None:
            rows.append(first_row + idx)
    return rows


def iter_candidate_rows(
    sheet: xw.Sheet,
    start_row: int,
    num_quarters_col: int | None,
    max_scan: int = 200,
) -> list[tuple[int, int]]:
    rows: list[tuple[int, int]] = []
    if num_quarters_col:
        values = flatten_single_column(
            sheet.range((start_row, num_quarters_col), (start_row + max_scan - 1, num_quarters_col)).value
        )
        blank_streak = 0
        for idx, value in enumerate(values):
            row = start_row + idx
            quarters = to_int(value)
            if quarters is None:
                blank_streak += 1
                if blank_streak >= 5 and rows:
                    break
                continue
            blank_streak = 0
            rows.append((row, quarters))

    if not rows:
        rows = [(start_row + offset, offset + 1) for offset in range(N_QUARTERS)]
    return rows


def set_formula2_r1c1(cell: xw.Range, formula_r1c1: str) -> None:
    try:
        cell.formula2 = formula_r1c1
    except Exception:
        cell.api.Formula2R1C1 = formula_r1c1


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


def process_empirical_sheet(wb: xw.Book, label: FileLabel, source_file: str) -> list[dict[str, Any]]:
    sheet = find_sheet_case_insensitive(wb, "Empirical Model")
    if sheet is None:
        print(f"  skipped Empirical Model for {source_file}: sheet not found")
        return []

    anchor_row, anchor_col, used = find_max_anchor(sheet)
    first_col = used.column
    last_col = used.column + used.columns.count - 1
    headers = read_header_map(sheet, anchor_row, first_col, last_col)

    num_quarters_col = pick_col(headers, ["num quarters", "quarters used", "n quarters", "num quarter"])
    last_quarter_col = pick_col(headers, ["last quarter used", "last quarter", "quarter used"])
    forecast_col = pick_col(
        headers,
        ["estimated total sold", "est total sold", "forecast value", "total sold", "tot fcst"],
        ["max", "min"],
    )
    actual_col = pick_col(headers, ["reported sales", "actual value", "actual sales", "actual"])
    forecast_min_col = pick_col(headers, ["min"])
    quarterly_sales_col = pick_col(headers, ["quarterly sales", "quarter sales"])
    reported_sales_col = pick_col(headers, ["reported sales"])
    growth_col = pick_col(headers, ["growth rate", "growth"])
    captured_col = pick_col(headers, ["sales captured in db", "captured in db", "penetration", "capture"])

    forecast_max_col = anchor_col
    if forecast_min_col is None:
        forecast_min_col = anchor_col + 1

    history_rows_capture = collect_numeric_rows(sheet, captured_col or -1, 1, anchor_row - 1)
    history_rows_ratio = collect_numeric_pair_rows(
        sheet,
        quarterly_sales_col or -1,
        reported_sales_col or -1,
        1,
        anchor_row - 1,
    )

    helper_col = last_col + 2
    helper_cell = sheet.cells(anchor_row, helper_col)
    candidate_rows = iter_candidate_rows(sheet, anchor_row + 1, num_quarters_col)

    records: list[dict[str, Any]] = []
    for row, n_quarters in candidate_rows:
        if n_quarters < 1 or n_quarters > N_QUARTERS:
            continue

        forecast_max = to_float(sheet.cells(row, forecast_max_col).value)
        forecast_min = to_float(sheet.cells(row, forecast_min_col).value)
        forecast_value = to_float(sheet.cells(row, forecast_col).value) if forecast_col else None
        actual_value = to_float(sheet.cells(row, actual_col).value) if actual_col else None
        quarterly_sales = to_float(sheet.cells(row, quarterly_sales_col).value) if quarterly_sales_col else None
        reported_sales = to_float(sheet.cells(row, reported_sales_col).value) if reported_sales_col else None
        growth_rate_pct = to_float(sheet.cells(row, growth_col).value) if growth_col else None
        sales_captured_in_db_pct = to_float(sheet.cells(row, captured_col).value) if captured_col else None
        last_quarter_used = sheet.cells(row, last_quarter_col).value if last_quarter_col else None

        row_has_data = any(
            value is not None
            for value in (
                forecast_max,
                forecast_min,
                forecast_value,
                actual_value,
                quarterly_sales,
                reported_sales,
                growth_rate_pct,
                sales_captured_in_db_pct,
            )
        )
        if not row_has_data:
            continue

        avg_penetration_pct: float | None = None
        if captured_col and len(history_rows_capture) >= n_quarters:
            start_row = history_rows_capture[-n_quarters]
            end_row = history_rows_capture[-1]
            set_formula2_r1c1(helper_cell, f"=AVERAGE(R{start_row}C{captured_col}:R{end_row}C{captured_col})")
            wb.app.calculate()
            avg_penetration_pct = to_float(helper_cell.value)
        elif quarterly_sales_col and reported_sales_col and len(history_rows_ratio) >= n_quarters:
            start_row = history_rows_ratio[-n_quarters]
            end_row = history_rows_ratio[-1]
            formula = (
                f"=IFERROR(SUM(R{start_row}C{quarterly_sales_col}:R{end_row}C{quarterly_sales_col})/"
                f"SUM(R{start_row}C{reported_sales_col}:R{end_row}C{reported_sales_col}),\"\")"
            )
            set_formula2_r1c1(helper_cell, formula)
            wb.app.calculate()
            avg_penetration_pct = to_float(helper_cell.value)

        if avg_penetration_pct is None:
            avg_penetration_pct = sales_captured_in_db_pct

        record = {
            "model": label.model,
            "ticker": label.ticker,
            "model_period": label.model_period,
            "model_date": label.model_date,
            "method": "empirical",
            "parameter_name": "avg_penetration_pct",
            "parameter_value": avg_penetration_pct,
            "num_quarters_used": n_quarters,
            "last_quarter_used": last_quarter_used,
            "forecast_value": forecast_value,
            "actual_value": actual_value,
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
        records.append(record)

    helper_cell.value = None
    return records


def process_regression_sheet(wb: xw.Book, label: FileLabel, source_file: str) -> list[dict[str, Any]]:
    sheet = find_sheet_case_insensitive(wb, "Regression Model")
    if sheet is None:
        print(f"  skipped Regression Model for {source_file}: sheet not found")
        return []

    anchor_row, anchor_col, used = find_max_anchor(sheet)
    first_col = used.column
    last_col = used.column + used.columns.count - 1
    headers = read_header_map(sheet, anchor_row, first_col, last_col)

    num_quarters_col = pick_col(headers, ["num quarters", "quarters used", "n quarters", "num quarter"])
    forecast_col = pick_col(
        headers,
        ["tot fcst w o sa", "tot fcst wo sa", "tot fcst w sa", "tot fcst", "forecast"],
        ["max", "min"],
    )
    actual_col = pick_col(headers, ["actual value", "actual sales", "reported sales", "actual"])
    forecast_min_col = pick_col(headers, ["min"])

    forecast_max_col = anchor_col
    if forecast_min_col is None:
        forecast_min_col = anchor_col + 1

    y_col = anchor_col - 7
    x_col = anchor_col - 11
    history_rows = collect_numeric_pair_rows(sheet, x_col, y_col, 1, anchor_row - 1)

    helper_col = last_col + 2
    intercept_cell = sheet.cells(anchor_row, helper_col)
    slope_cell = sheet.cells(anchor_row, helper_col + 1)

    candidate_rows = iter_candidate_rows(sheet, anchor_row + 1, num_quarters_col)
    records: list[dict[str, Any]] = []
    previous_signature: tuple[Any, ...] | None = None

    for row, n_quarters in candidate_rows:
        if n_quarters < 1 or n_quarters > N_QUARTERS:
            continue
        if len(history_rows) < n_quarters:
            continue

        start_row = history_rows[-n_quarters]
        end_row = history_rows[-1]

        set_formula2_r1c1(
            intercept_cell,
            f"=INTERCEPT(R{start_row}C{y_col}:R{end_row}C{y_col},R{start_row}C{x_col}:R{end_row}C{x_col})",
        )
        set_formula2_r1c1(
            slope_cell,
            f"=SLOPE(R{start_row}C{y_col}:R{end_row}C{y_col},R{start_row}C{x_col}:R{end_row}C{x_col})",
        )
        wb.app.calculate()

        intercept = to_float(intercept_cell.value)
        slope = to_float(slope_cell.value)
        forecast_value = to_float(sheet.cells(row, forecast_col).value) if forecast_col else None
        actual_value = to_float(sheet.cells(row, actual_col).value) if actual_col else None
        forecast_max = to_float(sheet.cells(row, forecast_max_col).value)
        forecast_min = to_float(sheet.cells(row, forecast_min_col).value)

        row_has_data = any(value is not None for value in (forecast_value, forecast_max, forecast_min, intercept, slope))
        if not row_has_data:
            continue

        signature = (
            n_quarters,
            round_signature(intercept),
            round_signature(slope),
            round_signature(forecast_value),
            round_signature(forecast_max),
            round_signature(forecast_min),
        )
        if signature == previous_signature:
            continue
        previous_signature = signature

        record = {
            "model": label.model,
            "ticker": label.ticker,
            "model_period": label.model_period,
            "model_date": label.model_date,
            "method": "regression",
            "parameter_name": "num_quarters_used",
            "parameter_value": n_quarters,
            "num_quarters_used": n_quarters,
            "forecast_value": forecast_value,
            "actual_value": actual_value,
            "forecast_max": forecast_max,
            "forecast_min": forecast_min,
            "range_width": safe_subtract(forecast_max, forecast_min),
            "intercept": intercept,
            "slope": slope,
            "source_file": source_file,
        }
        records.append(record)

    intercept_cell.value = None
    slope_cell.value = None
    return records


def autosize_columns(worksheet) -> None:
    for col_idx in range(1, worksheet.max_column + 1):
        max_len = 0
        for row_idx in range(1, worksheet.max_row + 1):
            value = worksheet.cell(row=row_idx, column=col_idx).value
            if value is None:
                continue
            text = str(value)
            if len(text) > max_len:
                max_len = len(text)
        worksheet.column_dimensions[get_column_letter(col_idx)].width = min(max(max_len + 2, 10), 42)


def write_output_sheet(workbook: Workbook, sheet_name: str, columns: list[str], rows: list[dict[str, Any]]) -> None:
    ws = workbook.create_sheet(title=sheet_name)
    ws.append(columns)
    for row in rows:
        ws.append([row.get(col) for col in columns])

    for cell in ws[1]:
        cell.font = Font(bold=True)
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    autosize_columns(ws)


def build_output_workbook(
    output_path: Path,
    empirical_rows: list[dict[str, Any]],
    regression_rows: list[dict[str, Any]],
) -> None:
    out_wb = Workbook()
    default_sheet = out_wb.active
    out_wb.remove(default_sheet)

    write_output_sheet(out_wb, "empirical_candidates", EMPIRICAL_COLUMNS, empirical_rows)
    write_output_sheet(out_wb, "regression_candidates", REGRESSION_COLUMNS, regression_rows)
    out_wb.save(output_path)


def main() -> None:
    if not input_dir.exists():
        print(f"Input directory not found: {input_dir.resolve()}")
        return

    output_dir.mkdir(parents=True, exist_ok=True)
    output_path = unique_output_path(input_dir, output_dir)

    files_processed = 0
    empirical_rows: list[dict[str, Any]] = []
    regression_rows: list[dict[str, Any]] = []

    app = xw.App(visible=False, add_book=False)
    app.display_alerts = False
    app.screen_updating = False
    app.calculation = "manual"

    try:
        for file_path in sorted(input_dir.iterdir()):
            if not file_path.is_file():
                continue

            if file_path.name.startswith("~"):
                print(f"Skipped {file_path.name}: temporary workbook")
                continue

            if file_path.suffix.lower() != ".xlsx":
                print(f"Skipped {file_path.name}: not an .xlsx file")
                continue

            try:
                label = parse_file_label(file_path.name)
            except Exception as exc:
                print(f"Skipped {file_path.name}: {exc}")
                continue

            workbook = None
            try:
                workbook = app.books.open(str(file_path), update_links=False)
                try:
                    file_empirical_rows = process_empirical_sheet(workbook, label, file_path.name)
                except Exception as exc:
                    file_empirical_rows = []
                    print(f"  skipped Empirical Model for {file_path.name}: {exc}")

                try:
                    file_regression_rows = process_regression_sheet(workbook, label, file_path.name)
                except Exception as exc:
                    file_regression_rows = []
                    print(f"  skipped Regression Model for {file_path.name}: {exc}")

                empirical_rows.extend(file_empirical_rows)
                regression_rows.extend(file_regression_rows)
                files_processed += 1
                print(
                    f"Processed {file_path.name}: "
                    f"empirical_rows={len(file_empirical_rows)}, regression_rows={len(file_regression_rows)}"
                )
            except Exception as exc:
                print(f"Skipped {file_path.name}: {exc}")
            finally:
                if workbook is not None:
                    safe_close_workbook(workbook)
    finally:
        app.quit()

    build_output_workbook(output_path, empirical_rows, regression_rows)

    print(f"Output workbook: {output_path.resolve()}")
    print(f"Files processed: {files_processed}")
    print(f"Empirical rows: {len(empirical_rows)}")
    print(f"Regression rows: {len(regression_rows)}")


if __name__ == "__main__":
    main()
