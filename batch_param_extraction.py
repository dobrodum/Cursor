#!/usr/bin/env python3
"""Batch parameter extraction from MedMiner Excel workbooks.

Processes a folder of MedMiner model workbooks and extracts parameter
sensitivity data from both "Empirical Model" and "Regression Model" sheets.
Writes a combined output workbook with formatted results.

Requires macOS or Windows with Microsoft Excel installed (uses xlwings).

Example:
    python3 batch_param_extraction.py \\
        --input-dir "/path/to/model/workbooks" \\
        --output-dir "/path/to/outputs"
"""

import argparse
import sys
from pathlib import Path

import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font


DEFAULT_N_QUARTERS = 10
DEFAULT_ROWS_PER_QUARTER = 3


def parse_args():
    parser = argparse.ArgumentParser(
        description="Batch-extract parameters from MedMiner Excel workbooks."
    )
    parser.add_argument(
        "--input-dir",
        required=True,
        help="Folder containing .xlsx workbooks to process.",
    )
    parser.add_argument(
        "--output-dir",
        required=True,
        help="Folder where the output workbook will be saved.",
    )
    parser.add_argument(
        "--n-quarters",
        type=int,
        default=DEFAULT_N_QUARTERS,
        help=f"Number of quarter windows to evaluate (default: {DEFAULT_N_QUARTERS}).",
    )
    parser.add_argument(
        "--rows-per-quarter",
        type=int,
        default=DEFAULT_ROWS_PER_QUARTER,
        help=f"Rows per quarter in the data (default: {DEFAULT_ROWS_PER_QUARTER}).",
    )
    parser.add_argument(
        "--visible",
        action="store_true",
        help="Keep Excel visible while processing.",
    )
    return parser.parse_args()


def normalize(val):
    if val is None:
        return ""
    return str(val).strip().lower()


def find_anchor(sheet, target="max"):
    max_row = sheet.used_range.last_cell.row
    max_col = sheet.used_range.last_cell.column

    for col in range(1, max_col + 1):
        vals = sheet.range(
            sheet.cells(1, col),
            sheet.cells(max_row, col),
        ).value

        if not isinstance(vals, list):
            vals = [vals]

        for row_idx, v in enumerate(vals, start=1):
            if normalize(v) == target:
                return sheet.cells(row_idx, col)

    return None


def make_source_file_label(file_path):
    parts = file_path.stem.split(" - ")

    if len(parts) < 3:
        raise ValueError("expected 'MedMiner_Model - TICKER - DATE_Send'")

    ticker = parts[1]
    raw_label = parts[-1].split("_")[0].replace("-", "_")
    file_label = f"{raw_label[:-4]}_{raw_label[-4:]}"

    return f"{ticker}_{file_label}"


def process_empirical(wb, file_path, n_quarters, rows_per_quarter):
    calc_pulls = []
    raw_pulls = []

    sheet = wb.sheets["Empirical Model"]

    anchor_e = find_anchor(sheet, "max")

    if anchor_e is None:
        print(f"Empirical MAX anchor not found — skipping {file_path.name}")
        return [], []

    max_row = anchor_e.row
    min_row = anchor_e.row + 1
    est_row = anchor_e.row - 1
    pen_row = anchor_e.row + 5

    est_col = anchor_e.column + 1
    sales_cap_col = anchor_e.column + 4
    param_col = anchor_e.column + 5
    b_col = anchor_e.column - 3
    e_col = anchor_e.column
    g_col = anchor_e.column + 2
    h_col = anchor_e.column + 3

    last_b_row = sheet.cells(sheet.cells.last_cell.row, b_col).end("up").row
    last_row = last_b_row - 3

    for i in range(1, n_quarters + 1):
        first_row = last_row - (i * rows_per_quarter) + 1

        if first_row < 1:
            break

        sheet.cells(pen_row, param_col).formula2 = (
            f"=AVERAGE(R{first_row}C{sales_cap_col}:R{last_row}C{sales_cap_col})"
        )

        wb.app.calculate()

        avg_pen = sheet.cells(pen_row, param_col).value
        est_total_sold = sheet.cells(est_row, est_col).value
        max_val = sheet.cells(max_row, est_col).value
        min_val = sheet.cells(min_row, est_col).value

        if isinstance(avg_pen, (int, float)):
            avg_pen = round(avg_pen * 100, 5)

        calc_pulls.append([
            i,
            avg_pen,
            est_total_sold,
            max_val,
            min_val,
        ])

    for k in range(1, 11):
        r = last_row - ((k - 1) * 3)

        if r < 1:
            break

        b_val = sheet.cells(r, b_col).value

        if isinstance(b_val, str) and len(b_val) == 4 and b_val[0].upper() == "Q":
            b_val = f"Q{b_val[1]} 20{b_val[2:]}"

        e_val = sheet.cells(r, e_col).value
        g_val = sheet.cells(r, g_col).value
        h_val = sheet.cells(r, h_col).value
        i_val = sheet.cells(r, sales_cap_col).value

        if isinstance(h_val, (int, float)):
            h_val = h_val * 100

        if isinstance(i_val, (int, float)):
            i_val = i_val * 100

        raw_pulls.append([
            b_val,
            e_val,
            g_val,
            h_val,
            i_val,
        ])

    return calc_pulls, raw_pulls


def process_regression(wb, file_path):
    calc_pulls = []

    sheet = wb.sheets["Regression Model"]

    anchor = find_anchor(sheet, "max")

    if anchor is None:
        print(f"Regression MAX anchor not found — skipping {file_path.name}")
        return []

    anchor_row = anchor.row
    anchor_col = anchor.column

    last_data_row = anchor_row - 1
    fixed_end = last_data_row - 1
    start_row = fixed_end - 3

    slope_col = anchor_col - 3
    intercept_col = anchor_col - 2
    tot_fcst_col = anchor_col + 1

    y_col = anchor_col - 7
    x_col = anchor_col - 11

    i = 0

    while True:
        current_start = start_row - i

        if (
            current_start < 1
            or sheet.cells(current_start, y_col).value is None
            or sheet.cells(current_start, x_col).value is None
        ):
            break

        sheet.cells(last_data_row, intercept_col).formula2 = (
            f"=INTERCEPT("
            f"R{current_start}C{y_col}:R{fixed_end}C{y_col},"
            f"R{current_start}C{x_col}:R{fixed_end}C{x_col})"
        )

        sheet.cells(last_data_row, slope_col).formula2 = (
            f"=SLOPE("
            f"R{current_start}C{y_col}:R{fixed_end}C{y_col},"
            f"R{current_start}C{x_col}:R{fixed_end}C{x_col})"
        )

        wb.app.calculate()

        row_values = [
            i + 4,
            sheet.cells(last_data_row, intercept_col).value,
            sheet.cells(last_data_row, slope_col).value,
            sheet.cells(last_data_row, tot_fcst_col).value,
            sheet.cells(anchor_row, tot_fcst_col).value,
            sheet.cells(anchor_row + 1, tot_fcst_col).value,
        ]

        if calc_pulls and row_values[1:] == calc_pulls[-1][1:]:
            break

        calc_pulls.append(row_values)

        i += 1

    return calc_pulls


def format_sheet(ws):
    for cell in ws[1]:
        cell.font = Font(size=12, bold=True)
        cell.alignment = Alignment(horizontal="center")

    ws.column_dimensions["A"].width = 32
    for col_idx in range(2, ws.max_column + 1):
        ws.column_dimensions[ws.cell(1, col_idx).column_letter].width = 16

    for row_idx in range(2, ws.max_row + 1):
        cell = ws.cell(row_idx, 1)
        cell.font = Font(size=12, bold=True)
        cell.alignment = Alignment(
            horizontal="center",
            vertical="center",
            wrap_text=True,
        )


def build_output_path(in_path, out_dir):
    output_file = out_dir / f"{in_path.name}_PARAM.xlsx"

    if output_file.exists():
        n = 1
        while True:
            candidate = out_dir / f"{in_path.name}_PARAM.{n}.xlsx"
            if not candidate.exists():
                output_file = candidate
                break
            n += 1

    return output_file


def main():
    args = parse_args()

    if sys.platform.startswith("linux"):
        raise RuntimeError(
            "xlwings needs desktop Excel. Run this script on a Mac or Windows "
            "machine that has Microsoft Excel installed."
        )

    try:
        import xlwings as xw
    except ImportError as exc:
        raise RuntimeError(
            "xlwings is not installed. Install it with 'pip install xlwings'."
        ) from exc

    in_path = Path(args.input_dir).expanduser().resolve()
    out_dir = Path(args.output_dir).expanduser().resolve()
    out_dir.mkdir(parents=True, exist_ok=True)

    output_file = build_output_path(in_path, out_dir)

    excel_files = [
        f for f in in_path.glob("*.xlsx")
        if not f.name.startswith("~")
        and f.resolve() != output_file.resolve()
    ]

    if not excel_files:
        raise FileNotFoundError(f"No .xlsx files found in {in_path}")

    empirical_results_per_file = []
    regression_results_per_file = []

    with xw.App(visible=args.visible) as app:
        app.display_alerts = False
        app.screen_updating = False

        for file_path in excel_files:
            try:
                source_file = make_source_file_label(file_path)

                wb = app.books.open(
                    str(file_path), read_only=True, update_links=False
                )

                try:
                    emp_calc, emp_raw = process_empirical(
                        wb, file_path, args.n_quarters, args.rows_per_quarter
                    )
                    reg_calc = process_regression(wb, file_path)

                    if emp_calc:
                        empirical_results_per_file.append(
                            (source_file, emp_calc, emp_raw)
                        )

                    if reg_calc:
                        regression_results_per_file.append((source_file, reg_calc))

                    print(f"Processed: {file_path.name}")

                finally:
                    wb.close()

            except Exception as e:
                print(f"Skipped {file_path.name}: {e}")

    if not empirical_results_per_file and not regression_results_per_file:
        raise ValueError("No files were processed successfully.")

    emp_col_calc = [
        "num_quarters_used",
        "avg_penetration_pct",
        "estimated_total_sold",
        "max",
        "min",
    ]

    emp_col_raw = [
        "last_quarter_used",
        "quarterly_sales",
        "reported_sales",
        "growth_rate_pct",
        "sales_captured_in_db_pct",
    ]

    reg_col_calc = [
        "num_quarters_used",
        "intercept",
        "slope",
        "forecast_total_wo_sa",
        "max",
        "min",
    ]

    emp_rows = []
    current_row = 2

    for file_label, calc, raw in empirical_results_per_file:
        df_c = pd.DataFrame(calc, columns=emp_col_calc)
        df_r = pd.DataFrame(raw, columns=emp_col_raw)
        df_file = pd.concat([df_c, df_r], axis=1)

        df_file.insert(0, "model", file_label)

        current_row += len(df_file)
        emp_rows.append(df_file)

    if emp_rows:
        df_empirical = pd.concat(emp_rows, ignore_index=True)
    else:
        df_empirical = pd.DataFrame(columns=["model"] + emp_col_calc + emp_col_raw)

    reg_rows = []
    current_row = 2

    for file_label, calc in regression_results_per_file:
        df_file = pd.DataFrame(calc, columns=reg_col_calc)

        df_file.insert(0, "model", file_label)

        current_row += len(df_file)
        reg_rows.append(df_file)

    if reg_rows:
        df_regression = pd.concat(reg_rows, ignore_index=True)
    else:
        df_regression = pd.DataFrame(columns=["model"] + reg_col_calc)

    with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
        df_empirical.to_excel(writer, index=False, sheet_name="Empirical Param")
        df_regression.to_excel(writer, index=False, sheet_name="Regression Param")

    wb2 = load_workbook(output_file)

    format_sheet(wb2["Empirical Param"])
    format_sheet(wb2["Regression Param"])

    wb2.save(output_file)

    print(f"\nSaved: {output_file}")
    print(f"Files processed: {len(excel_files)}")
    print(f"Empirical rows: {len(df_empirical)}")
    print(f"Regression rows: {len(df_regression)}")


if __name__ == "__main__":
    main()
