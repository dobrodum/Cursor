from __future__ import annotations

import re
from datetime import date, datetime
from pathlib import Path
from typing import Any

import openpyxl
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter
import xlwings as xw


# Configure these paths before running the script.
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

MODEL_DAY_MAP = {"early": 5, "mid": 15, "late": 25}


def to_2d(values: Any) -> list[list[Any]]:
    if values is None:
        return []
    if isinstance(values, (list, tuple)):
        if not values:
            return []
        first = values[0]
        if isinstance(first, (list, tuple)):
            return [list(row) for row in values]
        return [list(values)]
    return [[values]]


def normalize_text(value: Any) -> str:
    if value is None:
        return ""
    return re.sub(r"\s+", " ", str(value).strip().lower())


def to_float(value: Any) -> float | None:
    try:
        if value is None or value == "":
            return None
        return float(value)
    except (TypeError, ValueError):
        return None


def parse_file_label(file_path: Path) -> dict[str, str]:
    stem = file_path.stem
    parts = [part.strip() for part in stem.split(" - ")]

    ticker = parts[1] if len(parts) >= 2 else ""
    period_segment = parts[2] if len(parts) >= 3 else parts[-1]
    period_token = period_segment.split("_")[0]

    match = re.search(
        r"(Early|Mid|Late)([A-Za-z]{3,9})(\d{4})", period_token, flags=re.IGNORECASE
    )

    model_period = ""
    model_date = ""
    if match:
        part = match.group(1).title()
        month_token = match.group(2).title()[:3]
        year = int(match.group(3))
        try:
            month_num = datetime.strptime(month_token, "%b").month
            day = MODEL_DAY_MAP[part.lower()]
            model_period = f"{part}{month_token}_{year}"
            model_date = date(year, month_num, day).isoformat()
        except ValueError:
            model_period = f"{part}{month_token}_{year}"

    if not ticker:
        ticker = re.sub(r"[^A-Za-z0-9]+", "_", stem).strip("_")
    if not model_period:
        model_period = "unknown_period"

    return {
        "ticker": ticker,
        "model_period": model_period,
        "model_date": model_date,
        "model": f"{ticker}_{model_period}",
    }


def make_output_path(input_folder: Path, output_folder: Path) -> Path:
    input_name = input_folder.name or "input"
    output_folder.mkdir(parents=True, exist_ok=True)

    candidate = output_folder / f"{input_name}_PARAM.xlsx"
    if not candidate.exists():
        return candidate

    index = 1
    while True:
        candidate = output_folder / f"{input_name}_PARAM.{index}.xlsx"
        if not candidate.exists():
            return candidate
        index += 1


def safe_close_workbook(book: xw.Book) -> None:
    try:
        book.close(save=False)  # Newer xlwings versions may support this.
        return
    except TypeError:
        pass
    except Exception:
        pass

    try:
        book.api.Close(SaveChanges=False)
        return
    except Exception:
        pass

    try:
        book.close()
    except Exception:
        pass


def find_anchor(sheet: xw.Sheet, label: str = "max") -> tuple[int, int]:
    used = sheet.used_range
    data = to_2d(used.value)
    if not data:
        raise ValueError(f'No used range on sheet "{sheet.name}"')

    matches: list[tuple[int, int]] = []
    label_normalized = normalize_text(label)
    for r_offset, row in enumerate(data):
        for c_offset, value in enumerate(row):
            if normalize_text(value) == label_normalized:
                matches.append((used.row + r_offset, used.column + c_offset))

    if not matches:
        raise ValueError(f'Anchor "{label}" not found on sheet "{sheet.name}"')

    # Prefer the "max" label that sits next to a "min" label.
    for row, col in matches:
        right_text = normalize_text(sheet.range((row, col + 1)).value)
        left_text = normalize_text(sheet.range((row, col - 1)).value) if col > 1 else ""
        if right_text == "min" or left_text == "min":
            return row, col

    return matches[0]


def read_row_headers(
    sheet: xw.Sheet, row: int, col_start: int, col_end: int
) -> dict[str, int]:
    if col_end < col_start:
        return {}
    values = sheet.range((row, col_start), (row, col_end)).value
    if not isinstance(values, list):
        values = [values]
    mapping: dict[str, int] = {}
    for idx, value in enumerate(values):
        key = normalize_text(value)
        if key and key not in mapping:
            mapping[key] = col_start + idx
    return mapping


def pick_col(
    header_map: dict[str, int], fallback: int, keyword_options: list[tuple[str, ...]]
) -> int:
    for key, col in header_map.items():
        for keywords in keyword_options:
            if all(word in key for word in keywords):
                return col
    return max(1, fallback)


def get_last_data_row(sheet: xw.Sheet, col: int, min_row: int) -> int:
    last_row = sheet.range((sheet.cells.last_cell.row, col)).end("up").row
    return max(last_row, min_row)


def set_formula2_r1c1(cell: xw.Range, formula_r1c1: str) -> None:
    try:
        cell.formula2 = formula_r1c1
        return
    except Exception:
        pass
    try:
        cell.api.Formula2R1C1 = formula_r1c1
        return
    except Exception:
        pass
    cell.api.FormulaR1C1 = formula_r1c1


def compute_rolling_avg_penetration(
    sheet: xw.Sheet,
    source_col: int,
    source_start_row: int,
    source_end_row: int,
    n_quarters: int = 10,
) -> list[float | None]:
    if source_end_row < source_start_row:
        return [None] * n_quarters

    values = sheet.range((source_start_row, source_col), (source_end_row, source_col)).value
    if not isinstance(values, list):
        values = [values]
    numeric_rows: list[int] = []
    for idx, value in enumerate(values):
        if to_float(value) is not None:
            numeric_rows.append(source_start_row + idx)

    if not numeric_rows:
        return [None] * n_quarters

    history_start = numeric_rows[0]
    history_end = numeric_rows[-1]

    scratch_col = 16379  # XFA
    scratch_start_row = 2
    for idx in range(n_quarters):
        n = idx + 1
        start_row = max(history_start, history_end - n + 1)
        target = sheet.range((scratch_start_row + idx, scratch_col))
        formula = f"=AVERAGE(R{start_row}C{source_col}:R{history_end}C{source_col})"
        set_formula2_r1c1(target, formula)

    sheet.book.app.calculate()
    output_vals = sheet.range(
        (scratch_start_row, scratch_col), (scratch_start_row + n_quarters - 1, scratch_col)
    ).value
    if not isinstance(output_vals, list):
        output_vals = [output_vals]

    # Remove temporary formulas immediately.
    sheet.range(
        (scratch_start_row, scratch_col), (scratch_start_row + n_quarters - 1, scratch_col)
    ).clear_contents()

    return [to_float(v) for v in output_vals]


def process_empirical_sheet(
    sheet: xw.Sheet, meta: dict[str, str], source_file: str
) -> list[dict[str, Any]]:
    anchor_row, anchor_col = find_anchor(sheet, "max")
    n_quarters = 10

    header_map = read_row_headers(
        sheet, anchor_row, max(1, anchor_col - 20), anchor_col + 4
    )

    cols = {
        "num_quarters_used": pick_col(
            header_map,
            anchor_col - 11,
            [("num", "quarter"), ("quarters", "used"), ("n", "quarters")],
        ),
        "last_quarter_used": pick_col(
            header_map,
            anchor_col - 10,
            [("last", "quarter"), ("latest", "quarter")],
        ),
        "forecast_value": pick_col(
            header_map,
            anchor_col - 2,
            [("estimated", "total", "sold"), ("forecast", "total"), ("tot", "fcst")],
        ),
        "reported_sales": pick_col(
            header_map,
            anchor_col - 1,
            [("reported", "sales"), ("actual", "sales")],
        ),
        "quarterly_sales": pick_col(
            header_map,
            anchor_col - 6,
            [("quarterly", "sales"), ("quarter", "sales")],
        ),
        "growth_rate_pct": pick_col(
            header_map,
            anchor_col - 5,
            [("growth", "rate"), ("growth", "%")],
        ),
        "sales_captured_in_db_pct": pick_col(
            header_map,
            anchor_col - 4,
            [("captured", "db"), ("sales", "captured"), ("db", "%")],
        ),
        "avg_penetration_pct_cell": pick_col(
            header_map,
            anchor_col - 3,
            [("avg", "penetration"), ("average", "penetration")],
        ),
        "penetration_source_col": pick_col(
            header_map,
            anchor_col - 8,
            [("penetration",), ("sales", "captured", "db")],
        ),
        "forecast_max": anchor_col,
        "forecast_min": anchor_col + 1,
    }

    source_end_row = get_last_data_row(sheet, cols["penetration_source_col"], anchor_row + 1)
    avg_penetrations = compute_rolling_avg_penetration(
        sheet,
        source_col=cols["penetration_source_col"],
        source_start_row=anchor_row + 1,
        source_end_row=source_end_row,
        n_quarters=n_quarters,
    )

    row_start = anchor_row + 1
    row_end = row_start + n_quarters - 1
    col_start = max(
        1,
        min(
            cols["num_quarters_used"],
            cols["last_quarter_used"],
            cols["forecast_value"],
            cols["reported_sales"],
            cols["forecast_min"],
            cols["quarterly_sales"],
            cols["growth_rate_pct"],
            cols["sales_captured_in_db_pct"],
            cols["avg_penetration_pct_cell"],
        ),
    )
    col_end = max(
        cols["num_quarters_used"],
        cols["last_quarter_used"],
        cols["forecast_value"],
        cols["reported_sales"],
        cols["forecast_max"],
        cols["forecast_min"],
        cols["quarterly_sales"],
        cols["growth_rate_pct"],
        cols["sales_captured_in_db_pct"],
        cols["avg_penetration_pct_cell"],
    )

    block = to_2d(sheet.range((row_start, col_start), (row_end, col_end)).value)
    if len(block) < n_quarters:
        block.extend([[] for _ in range(n_quarters - len(block))])

    def block_value(row_idx: int, col: int) -> Any:
        row = block[row_idx]
        offset = col - col_start
        if offset < 0 or offset >= len(row):
            return None
        return row[offset]

    rows: list[dict[str, Any]] = []
    for idx in range(n_quarters):
        num_quarters = to_float(block_value(idx, cols["num_quarters_used"])) or float(idx + 1)
        last_quarter = block_value(idx, cols["last_quarter_used"])
        forecast_value = to_float(block_value(idx, cols["forecast_value"]))
        reported_sales = to_float(block_value(idx, cols["reported_sales"]))
        forecast_max = to_float(block_value(idx, cols["forecast_max"]))
        forecast_min = to_float(block_value(idx, cols["forecast_min"]))
        range_width = (
            forecast_max - forecast_min
            if forecast_max is not None and forecast_min is not None
            else None
        )

        fallback_avg = to_float(block_value(idx, cols["avg_penetration_pct_cell"]))
        avg_penetration = avg_penetrations[idx] if avg_penetrations[idx] is not None else fallback_avg

        row = {
            "model": meta["model"],
            "ticker": meta["ticker"],
            "model_period": meta["model_period"],
            "model_date": meta["model_date"],
            "method": "empirical",
            "parameter_name": "avg_penetration_pct",
            "parameter_value": avg_penetration,
            "num_quarters_used": int(num_quarters),
            "last_quarter_used": last_quarter,
            "forecast_value": forecast_value,
            "actual_value": reported_sales,
            "forecast_max": forecast_max,
            "forecast_min": forecast_min,
            "range_width": range_width,
            "avg_penetration_pct": avg_penetration,
            "quarterly_sales": to_float(block_value(idx, cols["quarterly_sales"])),
            "reported_sales": reported_sales,
            "growth_rate_pct": to_float(block_value(idx, cols["growth_rate_pct"])),
            "sales_captured_in_db_pct": to_float(block_value(idx, cols["sales_captured_in_db_pct"])),
            "source_file": source_file,
        }

        # Keep the row even when some optional fields are empty;
        # drop only completely empty candidates.
        has_signal = any(
            row.get(field) is not None
            for field in ("forecast_value", "forecast_max", "forecast_min", "avg_penetration_pct")
        )
        if has_signal:
            rows.append(row)

    return rows


def compute_regression_coefficients(
    sheet: xw.Sheet,
    x_col: int,
    y_col: int,
    num_quarters_list: list[int],
    data_start_row: int,
    data_end_row: int,
) -> list[tuple[float | None, float | None]]:
    if data_end_row < data_start_row:
        return [(None, None) for _ in num_quarters_list]

    col_left = min(x_col, y_col)
    col_right = max(x_col, y_col)
    x_offset = x_col - col_left
    y_offset = y_col - col_left

    values = to_2d(sheet.range((data_start_row, col_left), (data_end_row, col_right)).value)
    xy_rows: list[tuple[int, float, float]] = []
    for idx, row in enumerate(values):
        if x_offset >= len(row) or y_offset >= len(row):
            continue
        x_val = to_float(row[x_offset])
        y_val = to_float(row[y_offset])
        if x_val is None or y_val is None:
            continue
        xy_rows.append((data_start_row + idx, x_val, y_val))

    if len(xy_rows) < 2:
        return [(None, None) for _ in num_quarters_list]

    scratch_start_row = 2
    scratch_intercept_col = 16380  # XFB
    scratch_slope_col = 16381  # XFC

    formulas_written = 0
    for idx, n_quarters in enumerate(num_quarters_list):
        if n_quarters < 2 or n_quarters > len(xy_rows):
            continue

        start_row = xy_rows[-n_quarters][0]
        end_row = xy_rows[-1][0]
        intercept_cell = sheet.range((scratch_start_row + idx, scratch_intercept_col))
        slope_cell = sheet.range((scratch_start_row + idx, scratch_slope_col))
        intercept_formula = (
            f"=INTERCEPT(R{start_row}C{y_col}:R{end_row}C{y_col},"
            f"R{start_row}C{x_col}:R{end_row}C{x_col})"
        )
        slope_formula = (
            f"=SLOPE(R{start_row}C{y_col}:R{end_row}C{y_col},"
            f"R{start_row}C{x_col}:R{end_row}C{x_col})"
        )
        set_formula2_r1c1(intercept_cell, intercept_formula)
        set_formula2_r1c1(slope_cell, slope_formula)
        formulas_written += 1

    if formulas_written:
        sheet.book.app.calculate()

    results: list[tuple[float | None, float | None]] = []
    for idx, n_quarters in enumerate(num_quarters_list):
        if n_quarters < 2 or n_quarters > len(xy_rows):
            results.append((None, None))
            continue

        intercept = to_float(sheet.range((scratch_start_row + idx, scratch_intercept_col)).value)
        slope = to_float(sheet.range((scratch_start_row + idx, scratch_slope_col)).value)
        results.append((intercept, slope))

    # Clear temporary formulas once values are captured.
    if num_quarters_list:
        end_scratch_row = scratch_start_row + len(num_quarters_list) - 1
        sheet.range(
            (scratch_start_row, scratch_intercept_col), (end_scratch_row, scratch_slope_col)
        ).clear_contents()

    return results


def process_regression_sheet(
    sheet: xw.Sheet, meta: dict[str, str], source_file: str
) -> list[dict[str, Any]]:
    anchor_row, anchor_col = find_anchor(sheet, "max")

    header_map = read_row_headers(
        sheet, anchor_row, max(1, anchor_col - 20), anchor_col + 4
    )

    cols = {
        "num_quarters_used": pick_col(
            header_map,
            anchor_col - 11,
            [("num", "quarter"), ("quarters", "used"), ("n", "quarters")],
        ),
        "forecast_value": pick_col(
            header_map,
            anchor_col - 2,
            [("tot", "fcst", "w/o", "sa"), ("forecast", "without", "sa"), ("forecast", "total")],
        ),
        "forecast_max": anchor_col,
        "forecast_min": anchor_col + 1,
    }

    y_col = anchor_col - 7
    x_col = anchor_col - 11

    n_rows_to_scan = 40
    row_start = anchor_row + 1
    row_end = row_start + n_rows_to_scan - 1
    col_start = max(1, min(cols["num_quarters_used"], cols["forecast_value"], cols["forecast_max"]))
    col_end = max(cols["forecast_min"], cols["num_quarters_used"], cols["forecast_value"])

    grid = to_2d(sheet.range((row_start, col_start), (row_end, col_end)).value)
    row_indices: list[int] = []
    num_quarters: list[int] = []

    for idx, row in enumerate(grid):
        offset_num_q = cols["num_quarters_used"] - col_start
        offset_max = cols["forecast_max"] - col_start
        offset_min = cols["forecast_min"] - col_start

        num_q_value = to_float(row[offset_num_q]) if 0 <= offset_num_q < len(row) else None
        max_value = to_float(row[offset_max]) if 0 <= offset_max < len(row) else None
        min_value = to_float(row[offset_min]) if 0 <= offset_min < len(row) else None

        if num_q_value is None and max_value is None and min_value is None:
            continue
        row_indices.append(row_start + idx)
        num_quarters.append(int(num_q_value) if num_q_value is not None else idx + 1)

    if not row_indices:
        return []

    data_end_row = get_last_data_row(sheet, y_col, anchor_row + 1)
    coeffs = compute_regression_coefficients(
        sheet,
        x_col=x_col,
        y_col=y_col,
        num_quarters_list=num_quarters,
        data_start_row=anchor_row + 1,
        data_end_row=data_end_row,
    )

    rows: list[dict[str, Any]] = []
    previous_signature: tuple[Any, ...] | None = None
    for idx, row_num in enumerate(row_indices):
        n_quarters = num_quarters[idx]
        intercept, slope = coeffs[idx]

        forecast_value = to_float(sheet.range((row_num, cols["forecast_value"])).value)
        forecast_max = to_float(sheet.range((row_num, cols["forecast_max"])).value)
        forecast_min = to_float(sheet.range((row_num, cols["forecast_min"])).value)
        range_width = (
            forecast_max - forecast_min
            if forecast_max is not None and forecast_min is not None
            else None
        )

        signature = (
            n_quarters,
            round(forecast_value, 8) if forecast_value is not None else None,
            round(forecast_max, 8) if forecast_max is not None else None,
            round(forecast_min, 8) if forecast_min is not None else None,
            round(intercept, 8) if intercept is not None else None,
            round(slope, 8) if slope is not None else None,
        )
        if previous_signature is not None and signature == previous_signature:
            continue
        previous_signature = signature

        rows.append(
            {
                "model": meta["model"],
                "ticker": meta["ticker"],
                "model_period": meta["model_period"],
                "model_date": meta["model_date"],
                "method": "regression",
                "parameter_name": "num_quarters_used",
                "parameter_value": n_quarters,
                "num_quarters_used": n_quarters,
                "forecast_value": forecast_value,
                "actual_value": "",
                "forecast_max": forecast_max,
                "forecast_min": forecast_min,
                "range_width": range_width,
                "intercept": intercept,
                "slope": slope,
                "source_file": source_file,
            }
        )

    return rows


def write_output_workbook(
    output_path: Path,
    empirical_rows: list[dict[str, Any]],
    regression_rows: list[dict[str, Any]],
) -> None:
    workbook = openpyxl.Workbook()
    workbook.remove(workbook.active)

    def write_sheet(name: str, headers: list[str], rows: list[dict[str, Any]]) -> None:
        ws = workbook.create_sheet(title=name)
        ws.append(headers)
        for cell in ws[1]:
            cell.font = Font(bold=True)

        for row in rows:
            ws.append([row.get(col, "") for col in headers])

        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions

        for col_idx, header in enumerate(headers, start=1):
            max_len = len(header)
            for row_idx in range(2, ws.max_row + 1):
                value = ws.cell(row=row_idx, column=col_idx).value
                if value is None:
                    continue
                max_len = max(max_len, len(str(value)))
            ws.column_dimensions[get_column_letter(col_idx)].width = min(max_len + 2, 42)

    write_sheet("empirical_candidates", EMPIRICAL_HEADERS, empirical_rows)
    write_sheet("regression_candidates", REGRESSION_HEADERS, regression_rows)
    workbook.save(output_path)


def main() -> None:
    if not input_dir.exists():
        print(f"Skipped: input directory does not exist -> {input_dir.resolve()}")
        return

    output_path = make_output_path(input_dir, output_dir)
    empirical_rows: list[dict[str, Any]] = []
    regression_rows: list[dict[str, Any]] = []
    processed_files = 0

    app = xw.App(visible=False, add_book=False)
    app.display_alerts = False
    app.screen_updating = False

    try:
        for file_path in sorted(input_dir.iterdir()):
            if not file_path.is_file():
                print(f"Skipped: {file_path.name} (not a file)")
                continue
            if file_path.name.startswith("~"):
                print(f"Skipped: {file_path.name} (temporary file)")
                continue
            if file_path.suffix.lower() != ".xlsx":
                print(f"Skipped: {file_path.name} (not .xlsx)")
                continue

            print(f"Processed file: {file_path.name}")
            try:
                wb = app.books.open(str(file_path), update_links=False)
            except Exception as exc:
                print(f"Skipped: {file_path.name} (open failed: {exc})")
                continue

            try:
                meta = parse_file_label(file_path)
                if "Empirical Model" in [s.name for s in wb.sheets]:
                    empirical_rows.extend(
                        process_empirical_sheet(
                            wb.sheets["Empirical Model"], meta=meta, source_file=file_path.name
                        )
                    )
                else:
                    print(f'Skipped empirical: {file_path.name} ("Empirical Model" not found)')

                if "Regression Model" in [s.name for s in wb.sheets]:
                    regression_rows.extend(
                        process_regression_sheet(
                            wb.sheets["Regression Model"], meta=meta, source_file=file_path.name
                        )
                    )
                else:
                    print(f'Skipped regression: {file_path.name} ("Regression Model" not found)')

                processed_files += 1
            except Exception as exc:
                print(f"Skipped: {file_path.name} (processing failed: {exc})")
            finally:
                safe_close_workbook(wb)
    finally:
        app.quit()

    write_output_workbook(output_path, empirical_rows, regression_rows)
    print(f"Output path: {output_path.resolve()}")
    print(f"Number of files processed: {processed_files}")
    print(f"Number of empirical rows: {len(empirical_rows)}")
    print(f"Number of regression rows: {len(regression_rows)}")


if __name__ == "__main__":
    main()
