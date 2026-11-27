from __future__ import annotations

from dataclasses import dataclass
from datetime import date
from pathlib import Path
import re
from typing import Any

import openpyxl
from openpyxl.styles import Font
import xlwings as xw


# --------- User-configurable paths ---------
input_dir = Path("/path/to/input")
output_dir = Path("/path/to/output")


EMPIRICAL_SHEET_NAME = "Empirical Model"
REGRESSION_SHEET_NAME = "Regression Model"
N_QUARTERS = 10

# Anchor-based column offsets for empirical inputs.
# These defaults are based on common model layouts and can be adjusted if needed.
EMPIRICAL_COLUMN_OFFSET_CANDIDATES = [
    (-12, -11, -10),  # (quarter_label_col, quarterly_sales_col, reported_sales_col)
    (-11, -10, -9),
    (-10, -9, -8),
]

EARLY_MID_LATE_DAY = {"Early": 5, "Mid": 15, "Late": 25}

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

FILENAME_RE = re.compile(
    r".*-\s*(?P<ticker>[A-Za-z0-9]+)\s*-\s*"
    r"(?P<phase>Early|Mid|Late)(?P<month>[A-Za-z]+)(?P<year>\d{4}).*\.xlsx$",
    re.IGNORECASE,
)


@dataclass
class FileLabel:
    model: str
    ticker: str
    model_period: str
    model_date: str


def to_float(value: Any) -> float | None:
    if value is None or value == "":
        return None
    try:
        return float(value)
    except (TypeError, ValueError):
        return None


def fmt_number(value: float | None) -> float | str:
    return "" if value is None else value


def round_sig(value: float | None, digits: int = 10) -> float | None:
    if value is None:
        return None
    return round(value, digits)


def parse_file_label(file_name: str) -> FileLabel:
    match = FILENAME_RE.match(file_name)
    if not match:
        ticker_guess = file_name.replace(".xlsx", "")
        return FileLabel(
            model=f"{ticker_guess}_unknown",
            ticker=ticker_guess,
            model_period="unknown",
            model_date="",
        )

    ticker = match.group("ticker").upper()
    phase = match.group("phase").title()
    month_token = match.group("month")[:3].title()
    year = int(match.group("year"))
    day = EARLY_MID_LATE_DAY[phase]

    month_number = date.fromisoformat(f"{year}-01-01").replace(month=1).month
    month_number = {
        "Jan": 1,
        "Feb": 2,
        "Mar": 3,
        "Apr": 4,
        "May": 5,
        "Jun": 6,
        "Jul": 7,
        "Aug": 8,
        "Sep": 9,
        "Oct": 10,
        "Nov": 11,
        "Dec": 12,
    }.get(month_token, month_number)

    model_period = f"{phase}{month_token}_{year}"
    model_date = date(year, month_number, day).isoformat()
    model = f"{ticker}_{model_period}"

    return FileLabel(model=model, ticker=ticker, model_period=model_period, model_date=model_date)


def next_output_path(src_input_dir: Path, dst_output_dir: Path) -> Path:
    dst_output_dir.mkdir(parents=True, exist_ok=True)
    base_name = f"{src_input_dir.name}_PARAM.xlsx"
    output_path = dst_output_dir / base_name
    if not output_path.exists():
        return output_path

    counter = 1
    while True:
        candidate = dst_output_dir / f"{src_input_dir.name}_PARAM.{counter}.xlsx"
        if not candidate.exists():
            return candidate
        counter += 1


def get_sheet(wb: xw.Book, name: str) -> xw.Sheet | None:
    for sheet in wb.sheets:
        if sheet.name.strip().lower() == name.lower():
            return sheet
    return None


def find_max_anchor(sheet: xw.Sheet) -> tuple[int, int] | None:
    xl_values = -4163
    xl_whole = 1
    xl_by_rows = 1
    xl_next = 1

    found = sheet.api.Cells.Find(
        What="max",
        After=sheet.api.Cells(1, 1),
        LookIn=xl_values,
        LookAt=xl_whole,
        SearchOrder=xl_by_rows,
        SearchDirection=xl_next,
        MatchCase=False,
    )
    if found is None:
        return None
    return int(found.Row), int(found.Column)


def close_workbook_no_save(wb: xw.Book) -> None:
    try:
        wb.close(save=False)
        return
    except TypeError:
        pass
    except Exception:
        # Fall through to safe-close fallbacks.
        pass

    try:
        wb.api.Close(SaveChanges=False)
        return
    except TypeError:
        pass
    except Exception:
        pass

    try:
        wb.api.Close(False)
    except Exception:
        # Last resort: close without any save call.
        try:
            wb.close()
        except Exception:
            pass


def set_formula2_r1c1(cell: xw.Range, formula: str) -> None:
    try:
        cell.formula2 = formula
    except Exception:
        cell.api.Formula2R1C1 = formula


def collect_tail_numeric_rows(sheet: xw.Sheet, col: int, end_row: int) -> list[tuple[int, float]]:
    if col < 1 or end_row < 1:
        return []

    values = sheet.range((1, col), (end_row, col)).value
    if end_row == 1:
        values = [values]

    rows: list[tuple[int, float]] = []
    started = False
    for row in range(end_row, 0, -1):
        num = to_float(values[row - 1])
        if num is None:
            if started:
                break
            continue
        started = True
        rows.append((row, num))
    rows.reverse()
    return rows


def contiguous_tail(rows: list[int]) -> list[int]:
    if not rows:
        return []
    tail = [rows[-1]]
    for row in reversed(rows[:-1]):
        if tail[-1] - row == 1:
            tail.append(row)
        else:
            break
    return list(reversed(tail))


def make_empirical_rows(
    wb: xw.Book,
    sheet: xw.Sheet,
    anchor_row: int,
    anchor_col: int,
    label: FileLabel,
    source_file: str,
) -> list[dict[str, Any]]:
    best_records: list[dict[str, Any]] = []

    for quarter_off, q_sales_off, reported_off in EMPIRICAL_COLUMN_OFFSET_CANDIDATES:
        quarter_col = anchor_col + quarter_off
        q_sales_col = anchor_col + q_sales_off
        reported_col = anchor_col + reported_off
        if min(quarter_col, q_sales_col, reported_col) < 1:
            continue

        q_sales_tail = collect_tail_numeric_rows(sheet, q_sales_col, anchor_row - 1)
        reported_tail = collect_tail_numeric_rows(sheet, reported_col, anchor_row - 1)
        if not q_sales_tail or not reported_tail:
            continue

        row_to_q_sales = {r: v for r, v in q_sales_tail}
        row_to_reported = {r: v for r, v in reported_tail}
        common_rows = sorted(set(row_to_q_sales).intersection(row_to_reported))
        common_rows = contiguous_tail(common_rows)
        if len(common_rows) < 2:
            continue

        quarter_vals = sheet.range((common_rows[0], quarter_col), (common_rows[-1], quarter_col)).value
        if common_rows[0] == common_rows[-1]:
            quarter_vals = [quarter_vals]

        records: list[dict[str, Any]] = []
        for idx, row in enumerate(common_rows):
            reported = row_to_reported.get(row)
            q_sales = row_to_q_sales.get(row)
            if reported in (None, 0) or q_sales is None:
                continue

            penetration = q_sales / reported
            quarter_label = quarter_vals[idx] if idx < len(quarter_vals) else ""
            records.append(
                {
                    "row": row,
                    "quarter_label": quarter_label,
                    "quarterly_sales": q_sales,
                    "reported_sales": reported,
                    "sales_captured_in_db_pct": penetration,
                }
            )

        if len(records) > len(best_records):
            best_records = records

    if len(best_records) < 2:
        return []

    used_last_col = sheet.used_range.last_cell.column
    used_last_row = sheet.used_range.last_cell.row
    scratch_pen_col = max(used_last_col + 2, anchor_col + 2)
    scratch_calc_row = max(used_last_row + 2, anchor_row + 2)
    avg_cell = sheet.range((scratch_calc_row, scratch_pen_col))

    first_row = best_records[0]["row"]
    last_row = best_records[-1]["row"]
    penetration_col_values: list[list[Any]] = []
    row_to_pen = {rec["row"]: rec["sales_captured_in_db_pct"] for rec in best_records}
    for row in range(first_row, last_row + 1):
        penetration_col_values.append([row_to_pen.get(row, None)])
    sheet.range((first_row, scratch_pen_col), (last_row, scratch_pen_col)).value = penetration_col_values

    rows: list[dict[str, Any]] = []
    total_records = len(best_records)
    max_quarters = min(N_QUARTERS, total_records)
    for n_quarters_used in range(1, max_quarters + 1):
        subset = best_records[-n_quarters_used:]
        subset_start_row = subset[0]["row"]
        subset_end_row = subset[-1]["row"]

        set_formula2_r1c1(
            avg_cell,
            f"=AVERAGE(R{subset_start_row}C{scratch_pen_col}:R{subset_end_row}C{scratch_pen_col})",
        )
        wb.app.calculate()
        avg_penetration_pct = to_float(avg_cell.value)

        latest = subset[-1]
        previous = best_records[-2] if total_records >= 2 else None
        quarterly_sales = to_float(latest["quarterly_sales"])
        reported_sales = to_float(latest["reported_sales"])

        growth_rate_pct = None
        if previous:
            prev_reported = to_float(previous["reported_sales"])
            if prev_reported not in (None, 0) and reported_sales is not None:
                growth_rate_pct = (reported_sales / prev_reported) - 1

        subset_penetrations = [to_float(item["sales_captured_in_db_pct"]) for item in subset]
        subset_penetrations = [x for x in subset_penetrations if x not in (None, 0)]
        pen_min = min(subset_penetrations) if subset_penetrations else None
        pen_max = max(subset_penetrations) if subset_penetrations else None

        forecast_value = None
        if quarterly_sales is not None and avg_penetration_pct not in (None, 0):
            forecast_value = quarterly_sales / avg_penetration_pct

        forecast_max = None
        if quarterly_sales is not None and pen_min not in (None, 0):
            forecast_max = quarterly_sales / pen_min

        forecast_min = None
        if quarterly_sales is not None and pen_max not in (None, 0):
            forecast_min = quarterly_sales / pen_max

        range_width = None
        if forecast_max is not None and forecast_min is not None:
            range_width = forecast_max - forecast_min

        rows.append(
            {
                "model": label.model,
                "ticker": label.ticker,
                "model_period": label.model_period,
                "model_date": label.model_date,
                "method": "empirical",
                "parameter_name": "avg_penetration_pct",
                "parameter_value": fmt_number(avg_penetration_pct),
                "num_quarters_used": n_quarters_used,
                "last_quarter_used": str(latest["quarter_label"]) if latest["quarter_label"] is not None else "",
                "forecast_value": fmt_number(forecast_value),
                "actual_value": fmt_number(reported_sales),
                "forecast_max": fmt_number(forecast_max),
                "forecast_min": fmt_number(forecast_min),
                "range_width": fmt_number(range_width),
                "avg_penetration_pct": fmt_number(avg_penetration_pct),
                "quarterly_sales": fmt_number(quarterly_sales),
                "reported_sales": fmt_number(reported_sales),
                "growth_rate_pct": fmt_number(growth_rate_pct),
                "sales_captured_in_db_pct": fmt_number(
                    to_float(latest.get("sales_captured_in_db_pct"))
                ),
                "source_file": source_file,
            }
        )

    return rows


def make_regression_rows(
    wb: xw.Book,
    sheet: xw.Sheet,
    anchor_row: int,
    anchor_col: int,
    label: FileLabel,
    source_file: str,
) -> list[dict[str, Any]]:
    y_col = anchor_col - 7
    x_col = anchor_col - 11
    if min(x_col, y_col) < 1:
        return []

    x_tail = collect_tail_numeric_rows(sheet, x_col, anchor_row - 1)
    y_tail = collect_tail_numeric_rows(sheet, y_col, anchor_row - 1)
    if not x_tail or not y_tail:
        return []

    row_to_x = {r: v for r, v in x_tail}
    row_to_y = {r: v for r, v in y_tail}
    common_rows = sorted(set(row_to_x).intersection(row_to_y))
    common_rows = contiguous_tail(common_rows)
    if len(common_rows) < 2:
        return []

    records = [(row, row_to_x[row], row_to_y[row]) for row in common_rows]
    max_quarters = min(N_QUARTERS, len(records))

    x_future = to_float(sheet.range((anchor_row, x_col)).value)
    if x_future is None:
        x_future = records[-1][1] + 1

    used_last_col = sheet.used_range.last_cell.column
    used_last_row = sheet.used_range.last_cell.row
    scratch_col = max(used_last_col + 2, anchor_col + 2)
    scratch_row = max(used_last_row + 2, anchor_row + 2)
    intercept_cell = sheet.range((scratch_row, scratch_col))
    slope_cell = sheet.range((scratch_row, scratch_col + 1))
    forecast_cell = sheet.range((scratch_row, scratch_col + 2))

    intercept_ref = intercept_cell.get_address(row_absolute=True, column_absolute=True, include_sheetname=False)
    slope_ref = slope_cell.get_address(row_absolute=True, column_absolute=True, include_sheetname=False)

    rows: list[dict[str, Any]] = []
    previous_signature: tuple[Any, ...] | None = None

    for n_quarters_used in range(2, max_quarters + 1):
        subset = records[-n_quarters_used:]
        start_row = subset[0][0]
        end_row = subset[-1][0]

        set_formula2_r1c1(
            intercept_cell,
            f"=INTERCEPT(R{start_row}C{y_col}:R{end_row}C{y_col},R{start_row}C{x_col}:R{end_row}C{x_col})",
        )
        set_formula2_r1c1(
            slope_cell,
            f"=SLOPE(R{start_row}C{y_col}:R{end_row}C{y_col},R{start_row}C{x_col}:R{end_row}C{x_col})",
        )
        set_formula2_r1c1(forecast_cell, f"={intercept_ref}+{slope_ref}*{x_future}")
        wb.app.calculate()

        intercept = to_float(intercept_cell.value)
        slope = to_float(slope_cell.value)
        forecast_total_without_sa = to_float(forecast_cell.value)
        y_values = [row[2] for row in subset]
        forecast_max = max(y_values) if y_values else None
        forecast_min = min(y_values) if y_values else None
        range_width = None
        if forecast_max is not None and forecast_min is not None:
            range_width = forecast_max - forecast_min

        signature = (
            round_sig(intercept),
            round_sig(slope),
            round_sig(forecast_total_without_sa),
            round_sig(forecast_max),
            round_sig(forecast_min),
        )
        if previous_signature == signature:
            continue
        previous_signature = signature

        rows.append(
            {
                "model": label.model,
                "ticker": label.ticker,
                "model_period": label.model_period,
                "model_date": label.model_date,
                "method": "regression",
                "parameter_name": "num_quarters_used",
                "parameter_value": n_quarters_used,
                "num_quarters_used": n_quarters_used,
                "forecast_value": fmt_number(forecast_total_without_sa),
                "actual_value": "",
                "forecast_max": fmt_number(forecast_max),
                "forecast_min": fmt_number(forecast_min),
                "range_width": fmt_number(range_width),
                "intercept": fmt_number(intercept),
                "slope": fmt_number(slope),
                "source_file": source_file,
            }
        )

    return rows


def write_sheet(
    workbook: openpyxl.Workbook,
    sheet_name: str,
    headers: list[str],
    rows: list[dict[str, Any]],
) -> None:
    ws = workbook.create_sheet(title=sheet_name)
    ws.append(headers)
    for row in rows:
        ws.append([row.get(col, "") for col in headers])

    for cell in ws[1]:
        cell.font = Font(bold=True)

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions

    for col_idx, header in enumerate(headers, start=1):
        max_len = len(header)
        for row_idx in range(2, ws.max_row + 1):
            value = ws.cell(row=row_idx, column=col_idx).value
            if value is None:
                continue
            max_len = max(max_len, len(str(value)))
        ws.column_dimensions[openpyxl.utils.get_column_letter(col_idx)].width = min(max_len + 2, 60)


def main() -> None:
    if not input_dir.exists():
        raise FileNotFoundError(f"input_dir does not exist: {input_dir}")
    output_dir.mkdir(parents=True, exist_ok=True)

    output_path = next_output_path(input_dir, output_dir)

    candidates = sorted(input_dir.iterdir(), key=lambda p: p.name.lower())
    excel_files: list[Path] = []
    for path in candidates:
        if not path.is_file():
            print(f"Skipped: {path.name} (not a file)")
            continue
        if path.name.startswith("~"):
            print(f"Skipped: {path.name} (temporary file)")
            continue
        if path.suffix.lower() != ".xlsx":
            print(f"Skipped: {path.name} (not .xlsx)")
            continue
        excel_files.append(path)

    empirical_rows: list[dict[str, Any]] = []
    regression_rows: list[dict[str, Any]] = []
    processed_files = 0

    app = xw.App(visible=False, add_book=False)
    app.display_alerts = False
    app.screen_updating = False
    try:
        try:
            app.calculation = "manual"
        except Exception:
            pass

        for file_path in excel_files:
            wb: xw.Book | None = None
            try:
                wb = app.books.open(str(file_path), update_links=False)
                label = parse_file_label(file_path.name)

                empirical_sheet = get_sheet(wb, EMPIRICAL_SHEET_NAME)
                if empirical_sheet is None:
                    print(f"Skipped empirical: {file_path.name} (missing '{EMPIRICAL_SHEET_NAME}')")
                else:
                    empirical_anchor = find_max_anchor(empirical_sheet)
                    if empirical_anchor is None:
                        print(f"Skipped empirical: {file_path.name} (missing 'max' anchor)")
                    else:
                        rows = make_empirical_rows(
                            wb=wb,
                            sheet=empirical_sheet,
                            anchor_row=empirical_anchor[0],
                            anchor_col=empirical_anchor[1],
                            label=label,
                            source_file=file_path.name,
                        )
                        empirical_rows.extend(rows)

                regression_sheet = get_sheet(wb, REGRESSION_SHEET_NAME)
                if regression_sheet is None:
                    print(f"Skipped regression: {file_path.name} (missing '{REGRESSION_SHEET_NAME}')")
                else:
                    regression_anchor = find_max_anchor(regression_sheet)
                    if regression_anchor is None:
                        print(f"Skipped regression: {file_path.name} (missing 'max' anchor)")
                    else:
                        rows = make_regression_rows(
                            wb=wb,
                            sheet=regression_sheet,
                            anchor_row=regression_anchor[0],
                            anchor_col=regression_anchor[1],
                            label=label,
                            source_file=file_path.name,
                        )
                        regression_rows.extend(rows)

                processed_files += 1
                print(f"Processed: {file_path.name}")
            except Exception as exc:
                print(f"Skipped: {file_path.name} (error: {exc})")
            finally:
                if wb is not None:
                    close_workbook_no_save(wb)
    finally:
        app.quit()

    out_wb = openpyxl.Workbook()
    default_sheet = out_wb.active
    out_wb.remove(default_sheet)
    write_sheet(out_wb, "empirical_candidates", EMPIRICAL_HEADERS, empirical_rows)
    write_sheet(out_wb, "regression_candidates", REGRESSION_HEADERS, regression_rows)
    out_wb.save(output_path)

    print(f"Output path: {output_path}")
    print(f"Files processed: {processed_files}")
    print(f"Empirical rows: {len(empirical_rows)}")
    print(f"Regression rows: {len(regression_rows)}")


if __name__ == "__main__":
    main()
