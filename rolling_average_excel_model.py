#!/usr/bin/env python3
"""Simple rolling-average Excel model runner.

Run this on a Mac or Windows machine that has Microsoft Excel installed.

Example:
    python3 rolling_average_excel_model.py "/path/to/workbook.xlsx" --visible
"""

import argparse
import sys
from datetime import date, datetime, time, timedelta
from pathlib import Path


EXCEL_EPOCH = datetime(1899, 12, 30)

DEFAULT_SHEET = "Empirical Model"
DEFAULT_PARAM_CELL = "J271"
DEFAULT_TOTAL_CELL = "F265"
DEFAULT_MAX_CELL = "F266"
DEFAULT_MIN_CELL = "F267"
DEFAULT_DATE_COLUMN = "A"
DEFAULT_COLUMN_B = "B"
DEFAULT_COLUMN_C = "C"
DEFAULT_DATA_COLUMN = "I"
DEFAULT_ROWS_PER_QUARTER = 3
DEFAULT_START_ROW = 7
DEFAULT_LAST_ROW = 209

RESULT_HEADERS = [
    "Iteration",
    "Date",
    "Quarter",
    "Column B",
    "Column C",
    "Avg Penetration",
    "Total Sold",
    "Max Value",
    "Min Value",
    "Actual (if available)",
    "Error",
    "Error %",
]


def parse_args():
    parser = argparse.ArgumentParser(description="Run the rolling-average Excel model.")
    parser.add_argument("file_path", help="Path to the Excel workbook")
    parser.add_argument("--sheet-name", default=DEFAULT_SHEET)
    parser.add_argument("--param-cell", default=DEFAULT_PARAM_CELL)
    parser.add_argument("--total-cell", default=DEFAULT_TOTAL_CELL)
    parser.add_argument("--max-cell", default=DEFAULT_MAX_CELL)
    parser.add_argument("--min-cell", default=DEFAULT_MIN_CELL)
    parser.add_argument("--date-column", default=DEFAULT_DATE_COLUMN)
    parser.add_argument("--column-b", default=DEFAULT_COLUMN_B)
    parser.add_argument("--column-c", default=DEFAULT_COLUMN_C)
    parser.add_argument("--data-column", default=DEFAULT_DATA_COLUMN)
    parser.add_argument("--actual-column", default=None)
    parser.add_argument("--rows-per-quarter", type=int, default=DEFAULT_ROWS_PER_QUARTER)
    parser.add_argument("--start-row", type=int, default=DEFAULT_START_ROW)
    parser.add_argument("--last-row", type=int, default=DEFAULT_LAST_ROW)
    parser.add_argument("--output-dir", default=None)
    parser.add_argument("--visible", action="store_true")
    return parser.parse_args()


def build_output_path(source_path, output_dir, generated_at):
    if output_dir:
        target_dir = Path(output_dir).expanduser().resolve()
    else:
        target_dir = source_path.parent

    target_dir.mkdir(parents=True, exist_ok=True)
    timestamp = generated_at.strftime("%Y%m%d_%H%M%S")
    return target_dir / f"{source_path.stem}_ENHANCED_{timestamp}.xlsx"


def parse_excel_date(value):
    if value in (None, ""):
        return None

    if isinstance(value, datetime):
        return value

    if isinstance(value, date):
        return datetime.combine(value, time.min)

    if isinstance(value, (int, float)) and not isinstance(value, bool):
        return EXCEL_EPOCH + timedelta(days=float(value))

    if isinstance(value, str):
        text = value.strip()
        if not text:
            return None

        try:
            return datetime.fromisoformat(text)
        except ValueError:
            pass

        for fmt in ("%m/%d/%Y", "%m/%d/%y", "%d/%m/%Y", "%d/%m/%y", "%Y/%m/%d", "%Y-%m-%d"):
            try:
                return datetime.strptime(text, fmt)
            except ValueError:
                continue

    return None


def get_quarter(value):
    parsed = parse_excel_date(value)
    if parsed is None:
        return None if value in (None, "") else "Unknown"
    return f"Q{((parsed.month - 1) // 3) + 1}"


def get_window_starts(start_row, last_row, rows_per_quarter):
    if rows_per_quarter <= 0:
        raise ValueError("rows_per_quarter must be greater than 0")
    if start_row <= 0 or last_row <= 0:
        raise ValueError("start_row and last_row must be positive integers")
    if start_row > last_row:
        raise ValueError("start_row cannot be greater than last_row")

    starts = []
    current_start = max(start_row, last_row - rows_per_quarter + 1)

    while current_start > start_row:
        starts.append(current_start)
        current_start -= rows_per_quarter

    starts.append(start_row)
    return starts


def restore_parameter_cell(cell, original_formula, original_value):
    if isinstance(original_formula, str) and original_formula.startswith("="):
        cell.formula = original_formula
    else:
        cell.value = original_value


def calculate_error(total_sold, actual):
    if actual in (None, "") or total_sold is None:
        return None, None

    try:
        error = total_sold - actual
    except TypeError:
        return None, None

    if actual == 0:
        return error, None

    return error, error / actual


def collect_results(app, sheet, args):
    results = []
    parameter_cell = sheet.range(args.param_cell)
    original_formula = parameter_cell.formula
    original_value = parameter_cell.value

    try:
        for iteration, first_row in enumerate(
            get_window_starts(args.start_row, args.last_row, args.rows_per_quarter),
            start=1,
        ):
            range_used = f"{args.data_column}{first_row}:{args.data_column}{args.last_row}"
            run_date = sheet.range(f"{args.date_column}{first_row}").value

            parameter_cell.formula = f"=AVERAGE({range_used})"
            app.calculate()

            total_sold = sheet.range(args.total_cell).value
            max_value = sheet.range(args.max_cell).value
            min_value = sheet.range(args.min_cell).value
            actual = None

            if args.actual_column:
                actual = sheet.range(f"{args.actual_column}{first_row}").value

            error, error_pct = calculate_error(total_sold, actual)

            results.append(
                [
                    iteration,
                    run_date,
                    get_quarter(run_date),
                    sheet.range(f"{args.column_b}{first_row}").value,
                    sheet.range(f"{args.column_c}{first_row}").value,
                    parameter_cell.value,
                    total_sold,
                    max_value,
                    min_value,
                    actual,
                    error,
                    error_pct,
                ]
            )
    finally:
        restore_parameter_cell(parameter_cell, original_formula, original_value)

    return results


def build_summary_rows(source_path, output_path, args, results_count, generated_at):
    return [
        ["RUN SUMMARY", ""],
        ["File", source_path.name],
        ["Source Path", str(source_path)],
        ["Output Path", str(output_path)],
        ["Sheet", args.sheet_name],
        ["Method", "Rolling average penetration model"],
        ["Rows per quarter", args.rows_per_quarter],
        ["Start row", args.start_row],
        ["Last row", args.last_row],
        ["Iterations", results_count],
        ["Parameter cell", args.param_cell],
        ["Data column", args.data_column],
        ["Actual column", args.actual_column or "Not provided"],
        ["Generated at", generated_at.strftime("%Y-%m-%d %H:%M:%S")],
    ]


def write_output_workbook(app, results, source_path, output_path, args, generated_at):
    output_book = app.books.add()

    try:
        results_sheet = output_book.sheets[0]
        results_sheet.name = "Results"
        results_sheet.range("A1").value = RESULT_HEADERS
        results_sheet.range("A2").value = results

        summary_sheet = output_book.sheets.add("Run_Summary", after=results_sheet)
        summary_sheet.range("A1").value = build_summary_rows(
            source_path,
            output_path,
            args,
            len(results),
            generated_at,
        )

        try:
            results_sheet.autofit()
            summary_sheet.autofit()
        except Exception:
            pass

        output_book.save(str(output_path))
    finally:
        output_book.close()


def run_model(args):
    source_path = Path(args.file_path).expanduser().resolve()
    if not source_path.exists():
        raise FileNotFoundError(f"Excel file not found: {source_path}")

    if sys.platform.startswith("linux"):
        raise RuntimeError(
            "xlwings needs desktop Excel. Run this script on the Mac or Windows "
            "machine that has Microsoft Excel and the workbook."
        )

    try:
        import xlwings as xw
    except ImportError as exc:
        raise RuntimeError(
            "xlwings is not installed. Install it with 'pip install xlwings'."
        ) from exc

    generated_at = datetime.now()
    output_path = build_output_path(source_path, args.output_dir, generated_at)

    with xw.App(visible=args.visible, add_book=False) as app:
        app.display_alerts = False

        try:
            app.screen_updating = False
        except Exception:
            pass

        workbook = app.books.open(str(source_path), update_links=False, read_only=False)

        try:
            sheet = workbook.sheets[args.sheet_name]
            results = collect_results(app, sheet, args)
            write_output_workbook(app, results, source_path, output_path, args, generated_at)
        finally:
            workbook.close()

    return output_path


def main():
    args = parse_args()
    output_path = run_model(args)
    print(f"Done - saved as {output_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
