from __future__ import annotations

from dataclasses import dataclass
from datetime import date
from pathlib import Path
import re
from typing import Any, Iterable

import xlwings as xw
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

# User-configurable paths.
input_dir = Path("./input")
output_dir = Path("./output")

N_QUARTERS = 10
EMPIRICAL_SHEET_NAME = "Empirical Model"
REGRESSION_SHEET_NAME = "Regression Model"

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

MONTH_TO_NUM = {
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

PHASE_TO_DAY = {"Early": 5, "Mid": 15, "Late": 25}
PERIOD_RE = re.compile(r"(Early|Mid|Late)\s*([A-Za-z]{3,9})\s*[_-]?\s*(\d{4})", re.IGNORECASE)


@dataclass
class SheetSnapshot:
    ws: Any
    start_row: int
    start_col: int
    values: list[list[Any]]
    text_cells: list[tuple[str, int, int]]


def normalize_text(value: Any) -> str:
    if value is None:
        return ""
    text = str(value).strip().lower()
    text = text.replace("\n", " ")
    text = re.sub(r"\s+", " ", text)
    return text


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


def build_snapshot(ws: Any) -> SheetSnapshot:
    used = ws.used_range
    matrix = to_2d(used.value)
    start_row = int(used.row)
    start_col = int(used.column)
    text_cells: list[tuple[str, int, int]] = []

    for row_idx, row_values in enumerate(matrix):
        abs_row = start_row + row_idx
        for col_idx, raw_value in enumerate(row_values):
            if isinstance(raw_value, str):
                norm = normalize_text(raw_value)
                if norm:
                    abs_col = start_col + col_idx
                    text_cells.append((norm, abs_row, abs_col))

    return SheetSnapshot(
        ws=ws,
        start_row=start_row,
        start_col=start_col,
        values=matrix,
        text_cells=text_cells,
    )


def value_from_snapshot(snapshot: SheetSnapshot, row: int, col: int) -> Any:
    row_idx = row - snapshot.start_row
    col_idx = col - snapshot.start_col
    if row_idx < 0 or col_idx < 0:
        return None
    if row_idx >= len(snapshot.values):
        return None
    row_values = snapshot.values[row_idx]
    if col_idx >= len(row_values):
        return None
    return row_values[col_idx]


def to_number(value: Any) -> float | None:
    if value is None:
        return None
    if isinstance(value, bool):
        return float(value)
    if isinstance(value, (int, float)):
        return float(value)
    if isinstance(value, str):
        text = value.strip().replace(",", "")
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
    return None


def find_anchor_max(snapshot: SheetSnapshot) -> tuple[int, int] | None:
    exact: tuple[int, int] | None = None
    partial: tuple[int, int] | None = None
    for norm, row, col in snapshot.text_cells:
        if norm == "max" and exact is None:
            exact = (row, col)
        elif "max" in norm and partial is None:
            partial = (row, col)
    return exact or partial


def best_match_cell(
    snapshot: SheetSnapshot,
    aliases: Iterable[str],
    anchor_row: int,
    anchor_col: int,
) -> tuple[int, int] | None:
    alias_norm = [normalize_text(alias) for alias in aliases]
    best: tuple[int, int, int] | None = None

    for norm, row, col in snapshot.text_cells:
        if not any(norm == alias or alias in norm for alias in alias_norm):
            continue
        distance = abs(row - anchor_row) + abs(col - anchor_col)
        if best is None or distance < best[0]:
            best = (distance, row, col)

    if best is None:
        return None
    return best[1], best[2]


def locate_value_cell(
    snapshot: SheetSnapshot,
    aliases: Iterable[str],
    anchor_row: int,
    anchor_col: int,
) -> tuple[int, int] | None:
    label_cell = best_match_cell(snapshot, aliases, anchor_row, anchor_col)
    if label_cell is None:
        return None
    row, col = label_cell

    candidates = [
        (row, col + 1),
        (row, col + 2),
        (row, col + 3),
        (row + 1, col),
        (row + 1, col + 1),
        (row - 1, col + 1),
    ]

    for cand_row, cand_col in candidates:
        if cand_row < 1 or cand_col < 1:
            continue
        try:
            value = snapshot.ws.range((cand_row, cand_col)).value
        except Exception:
            continue
        if value not in (None, ""):
            return cand_row, cand_col

    return row, col + 1


def read_cell(ws: Any, cell: tuple[int, int] | None) -> Any:
    if cell is None:
        return None
    return ws.range(cell).value


def safe_close_workbook(wb: Any) -> None:
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
        return
    except Exception:
        pass

    wb.api.Close(False)


def set_formula_r1c1(cell: Any, formula_r1c1: str) -> None:
    try:
        cell.api.Formula2R1C1 = formula_r1c1
        return
    except Exception:
        pass

    try:
        cell.api.FormulaR1C1 = formula_r1c1
        return
    except Exception:
        pass

    try:
        cell.formula2 = formula_r1c1
        return
    except Exception:
        pass

    cell.formula = formula_r1c1


def parse_file_metadata(file_path: Path) -> dict[str, str]:
    stem = file_path.stem
    parts = [part.strip() for part in stem.split(" - ") if part.strip()]

    ticker = ""
    if len(parts) >= 2:
        ticker = re.split(r"[\s_]", parts[1])[0].upper()
    if not ticker:
        fallback = re.search(r"\b[A-Z]{2,8}\b", stem)
        if fallback:
            ticker = fallback.group(0).upper()
    if not ticker:
        raise ValueError("could not parse ticker from filename")

    period_match = PERIOD_RE.search(stem)
    if not period_match:
        raise ValueError("could not parse model period from filename")

    phase_raw, month_raw, year_raw = period_match.groups()
    phase = phase_raw.capitalize()
    month_key = month_raw[:3].lower()
    if month_key not in MONTH_TO_NUM:
        raise ValueError(f"unsupported month token: {month_raw}")

    year = int(year_raw)
    month_num = MONTH_TO_NUM[month_key]
    month_abbrev = month_key.capitalize()
    day = PHASE_TO_DAY[phase]

    model_period = f"{phase}{month_abbrev}_{year}"
    model_date = date(year, month_num, day).isoformat()
    model = f"{ticker}_{model_period}"

    return {
        "ticker": ticker,
        "model_period": model_period,
        "model_date": model_date,
        "model": model,
    }


def get_output_path(input_folder: Path, target_output_dir: Path) -> Path:
    base_name = f"{input_folder.resolve().name}_PARAM"
    candidate = target_output_dir / f"{base_name}.xlsx"
    suffix = 1
    while candidate.exists():
        candidate = target_output_dir / f"{base_name}.{suffix}.xlsx"
        suffix += 1
    return candidate


def compute_range_width(forecast_max: Any, forecast_min: Any) -> float | None:
    max_num = to_number(forecast_max)
    min_num = to_number(forecast_min)
    if max_num is None or min_num is None:
        return None
    return max_num - min_num


def restore_cell_value(cell: Any, original_formula2: Any, original_formula: Any, original_value: Any) -> None:
    if original_formula2 not in (None, ""):
        try:
            cell.formula2 = original_formula2
            return
        except Exception:
            pass

    if original_formula not in (None, ""):
        try:
            cell.formula = original_formula
            return
        except Exception:
            pass

    cell.value = original_value


def extract_empirical_rows(wb: Any, metadata: dict[str, str], source_file: str) -> list[dict[str, Any]]:
    ws = wb.sheets[EMPIRICAL_SHEET_NAME]
    snapshot = build_snapshot(ws)
    anchor = find_anchor_max(snapshot)
    if anchor is None:
        raise ValueError("could not find 'max' anchor in Empirical Model")
    anchor_row, anchor_col = anchor

    max_value_cell = locate_value_cell(snapshot, ["max"], anchor_row, anchor_col) or (anchor_row, anchor_col + 1)
    min_value_cell = locate_value_cell(snapshot, ["min"], anchor_row, anchor_col)
    avg_pen_cell_ref = locate_value_cell(
        snapshot,
        ["avg penetration", "average penetration", "avg pen"],
        anchor_row,
        anchor_col,
    )
    if avg_pen_cell_ref is None:
        raise ValueError("could not find avg penetration value cell in Empirical Model")

    est_total_cell = locate_value_cell(
        snapshot,
        ["estimated total sold", "estimated total", "est total sold", "forecast total sold"],
        anchor_row,
        anchor_col,
    )
    reported_sales_cell = locate_value_cell(
        snapshot,
        ["reported sales", "actual sales", "sales reported"],
        anchor_row,
        anchor_col,
    )
    quarterly_sales_cell = locate_value_cell(
        snapshot,
        ["quarterly sales", "quarterly sale"],
        anchor_row,
        anchor_col,
    )
    growth_rate_cell = locate_value_cell(
        snapshot,
        ["growth rate %", "growth rate", "growth %"],
        anchor_row,
        anchor_col,
    )
    captured_pct_cell = locate_value_cell(
        snapshot,
        ["sales captured in db %", "sales captured in db", "captured in db %", "captured in db"],
        anchor_row,
        anchor_col,
    )
    last_quarter_cell = locate_value_cell(
        snapshot,
        ["last quarter used", "last qtr used", "last quarter"],
        anchor_row,
        anchor_col,
    )

    avg_pen_cell = ws.range(avg_pen_cell_ref)
    try:
        original_formula2 = avg_pen_cell.formula2
    except Exception:
        original_formula2 = None
    try:
        original_formula = avg_pen_cell.formula
    except Exception:
        original_formula = None
    original_value = avg_pen_cell.value

    rows: list[dict[str, Any]] = []
    for num_quarters_used in range(1, N_QUARTERS + 1):
        avg_formula = f'=IFERROR(AVERAGE(RC[-{num_quarters_used}]:RC[-1]),"")'
        set_formula_r1c1(avg_pen_cell, avg_formula)
        wb.app.calculate()

        avg_penetration_pct = avg_pen_cell.value
        forecast_value = read_cell(ws, est_total_cell)
        reported_sales = read_cell(ws, reported_sales_cell)
        forecast_max = read_cell(ws, max_value_cell)
        forecast_min = read_cell(ws, min_value_cell)
        quarterly_sales = read_cell(ws, quarterly_sales_cell)
        growth_rate_pct = read_cell(ws, growth_rate_cell)
        sales_captured_in_db_pct = read_cell(ws, captured_pct_cell)
        last_quarter_used = read_cell(ws, last_quarter_cell)
        if last_quarter_used in (None, ""):
            # Fallback to the header directly above the latest quarter input.
            last_quarter_used = ws.range((avg_pen_cell_ref[0] - 1, max(avg_pen_cell_ref[1] - 1, 1))).value

        rows.append(
            {
                "model": metadata["model"],
                "ticker": metadata["ticker"],
                "model_period": metadata["model_period"],
                "model_date": metadata["model_date"],
                "method": "empirical",
                "parameter_name": "avg_penetration_pct",
                "parameter_value": avg_penetration_pct,
                "num_quarters_used": num_quarters_used,
                "last_quarter_used": last_quarter_used,
                "forecast_value": forecast_value,
                "actual_value": reported_sales,
                "forecast_max": forecast_max,
                "forecast_min": forecast_min,
                "range_width": compute_range_width(forecast_max, forecast_min),
                "avg_penetration_pct": avg_penetration_pct,
                "quarterly_sales": quarterly_sales,
                "reported_sales": reported_sales,
                "growth_rate_pct": growth_rate_pct,
                "sales_captured_in_db_pct": sales_captured_in_db_pct,
                "source_file": source_file,
            }
        )

    restore_cell_value(avg_pen_cell, original_formula2, original_formula, original_value)
    wb.app.calculate()
    return rows


def rows_are_duplicates(left: dict[str, Any], right: dict[str, Any]) -> bool:
    numeric_keys = ["forecast_value", "forecast_max", "forecast_min", "intercept", "slope"]
    for key in numeric_keys:
        left_num = to_number(left.get(key))
        right_num = to_number(right.get(key))
        if left_num is None and right_num is None:
            continue
        if left_num is None or right_num is None:
            return False
        if abs(left_num - right_num) > 1e-10:
            return False
    return True


def extract_regression_rows(wb: Any, metadata: dict[str, str], source_file: str) -> list[dict[str, Any]]:
    ws = wb.sheets[REGRESSION_SHEET_NAME]
    snapshot = build_snapshot(ws)
    anchor = find_anchor_max(snapshot)
    if anchor is None:
        raise ValueError("could not find 'max' anchor in Regression Model")
    anchor_row, anchor_col = anchor

    y_col = anchor_col - 7
    x_col = anchor_col - 11
    if y_col < 1 or x_col < 1:
        raise ValueError("invalid regression x/y column offsets from max anchor")

    numeric_pairs: list[tuple[float, float]] = []
    for row_offset, _ in enumerate(snapshot.values):
        abs_row = snapshot.start_row + row_offset
        x_value = value_from_snapshot(snapshot, abs_row, x_col)
        y_value = value_from_snapshot(snapshot, abs_row, y_col)
        x_num = to_number(x_value)
        y_num = to_number(y_value)
        if x_num is None or y_num is None:
            continue
        numeric_pairs.append((x_num, y_num))

    if len(numeric_pairs) < 2:
        raise ValueError("not enough numeric x/y points for regression")

    numeric_pairs = numeric_pairs[-N_QUARTERS:]
    max_n = len(numeric_pairs)
    x_values = [pair[0] for pair in numeric_pairs]
    y_values = [pair[1] for pair in numeric_pairs]

    forecast_cell = locate_value_cell(
        snapshot,
        ["tot fcst w/o sa", "tot fcst wo sa", "total forecast w/o sa", "tot fcst without sa"],
        anchor_row,
        anchor_col,
    )
    max_value_cell = locate_value_cell(snapshot, ["max"], anchor_row, anchor_col) or (anchor_row, anchor_col + 1)
    min_value_cell = locate_value_cell(snapshot, ["min"], anchor_row, anchor_col)
    actual_value_cell = locate_value_cell(
        snapshot,
        ["actual value", "actual sales", "reported sales"],
        anchor_row,
        anchor_col,
    )

    temp_row = anchor_row + 2
    temp_col = anchor_col + 8
    ws.range((temp_row, temp_col)).value = [[x, y] for x, y in zip(x_values, y_values)]

    intercept_cell = ws.range((temp_row, temp_col + 3))
    slope_cell = ws.range((temp_row, temp_col + 4))
    forecast_calc_cell = ws.range((temp_row, temp_col + 5))

    rows: list[dict[str, Any]] = []
    for num_quarters_used in range(2, max_n + 1):
        range_start = temp_row + (max_n - num_quarters_used)
        range_end = temp_row + max_n - 1

        intercept_formula = (
            f"=INTERCEPT(R{range_start}C{temp_col + 1}:R{range_end}C{temp_col + 1},"
            f"R{range_start}C{temp_col}:R{range_end}C{temp_col})"
        )
        slope_formula = (
            f"=SLOPE(R{range_start}C{temp_col + 1}:R{range_end}C{temp_col + 1},"
            f"R{range_start}C{temp_col}:R{range_end}C{temp_col})"
        )
        # Predict one step ahead off the latest x to create TOT FCST w/o SA fallback.
        forecast_formula = (
            f"=R{temp_row}C{temp_col + 3}+R{temp_row}C{temp_col + 4}"
            f"*(R{range_end}C{temp_col}+1)"
        )

        set_formula_r1c1(intercept_cell, intercept_formula)
        set_formula_r1c1(slope_cell, slope_formula)
        set_formula_r1c1(forecast_calc_cell, forecast_formula)
        wb.app.calculate()

        intercept = intercept_cell.value
        slope = slope_cell.value
        forecast_value = read_cell(ws, forecast_cell)
        if forecast_value in (None, ""):
            forecast_value = forecast_calc_cell.value

        fallback_max = max(y_values[-num_quarters_used:])
        fallback_min = min(y_values[-num_quarters_used:])
        forecast_max = read_cell(ws, max_value_cell)
        forecast_min = read_cell(ws, min_value_cell)
        if forecast_max in (None, ""):
            forecast_max = fallback_max
        if forecast_min in (None, ""):
            forecast_min = fallback_min

        row = {
            "model": metadata["model"],
            "ticker": metadata["ticker"],
            "model_period": metadata["model_period"],
            "model_date": metadata["model_date"],
            "method": "regression",
            "parameter_name": "num_quarters_used",
            "parameter_value": num_quarters_used,
            "num_quarters_used": num_quarters_used,
            "forecast_value": forecast_value,
            "actual_value": read_cell(ws, actual_value_cell) or "",
            "forecast_max": forecast_max,
            "forecast_min": forecast_min,
            "range_width": compute_range_width(forecast_max, forecast_min),
            "intercept": intercept,
            "slope": slope,
            "source_file": source_file,
        }

        if rows and rows_are_duplicates(rows[-1], row):
            continue
        rows.append(row)

    return rows


def write_sheet(workbook: Workbook, name: str, columns: list[str], rows: list[dict[str, Any]]) -> None:
    ws = workbook.create_sheet(name)
    ws.append(columns)

    for row in rows:
        ws.append([row.get(column, "") for column in columns])

    for cell in ws[1]:
        cell.font = Font(bold=True)

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(ws.max_column)}{ws.max_row}"

    for col_idx, col_name in enumerate(columns, start=1):
        max_len = len(col_name)
        for row_idx in range(2, ws.max_row + 1):
            value = ws.cell(row=row_idx, column=col_idx).value
            if value is None:
                continue
            max_len = max(max_len, len(str(value)))
        ws.column_dimensions[get_column_letter(col_idx)].width = min(max_len + 2, 40)


def save_output(
    output_path: Path,
    empirical_rows: list[dict[str, Any]],
    regression_rows: list[dict[str, Any]],
) -> None:
    wb = Workbook()
    wb.remove(wb.active)
    write_sheet(wb, "empirical_candidates", EMPIRICAL_COLUMNS, empirical_rows)
    write_sheet(wb, "regression_candidates", REGRESSION_COLUMNS, regression_rows)
    wb.save(output_path)


def configure_excel_app(app: Any) -> None:
    app.visible = False
    try:
        app.display_alerts = False
    except Exception:
        pass
    try:
        app.screen_updating = False
    except Exception:
        pass
    try:
        app.calculation = "manual"
    except Exception:
        pass
    try:
        app.api.EnableEvents = False
    except Exception:
        pass


def main() -> None:
    if not input_dir.exists():
        print(f"Skipped all files: input directory does not exist: {input_dir}")
        return

    output_dir.mkdir(parents=True, exist_ok=True)
    output_path = get_output_path(input_dir, output_dir)

    empirical_rows: list[dict[str, Any]] = []
    regression_rows: list[dict[str, Any]] = []
    processed_files = 0

    files = sorted(input_dir.iterdir(), key=lambda p: p.name.lower())
    app = xw.App(visible=False, add_book=False)
    configure_excel_app(app)

    try:
        for file_path in files:
            if not file_path.is_file():
                continue
            if file_path.suffix.lower() != ".xlsx":
                print(f"Skipped {file_path.name}: not an .xlsx file")
                continue
            if file_path.name.startswith("~"):
                print(f"Skipped {file_path.name}: temporary file")
                continue

            try:
                metadata = parse_file_metadata(file_path)
            except ValueError as exc:
                print(f"Skipped {file_path.name}: {exc}")
                continue

            wb = None
            try:
                wb = app.books.open(str(file_path), update_links=False)

                file_empirical_rows: list[dict[str, Any]] = []
                file_regression_rows: list[dict[str, Any]] = []

                try:
                    file_empirical_rows = extract_empirical_rows(wb, metadata, file_path.name)
                except Exception as exc:
                    print(f"Skipped empirical extraction for {file_path.name}: {exc}")

                try:
                    file_regression_rows = extract_regression_rows(wb, metadata, file_path.name)
                except Exception as exc:
                    print(f"Skipped regression extraction for {file_path.name}: {exc}")

                empirical_rows.extend(file_empirical_rows)
                regression_rows.extend(file_regression_rows)
                processed_files += 1
                print(f"Processed {file_path.name}")
            except Exception as exc:
                print(f"Skipped {file_path.name}: workbook open/process error: {exc}")
            finally:
                if wb is not None:
                    safe_close_workbook(wb)
    finally:
        app.quit()

    save_output(output_path, empirical_rows, regression_rows)

    print(f"Output path: {output_path}")
    print(f"Number of files processed: {processed_files}")
    print(f"Number of empirical rows: {len(empirical_rows)}")
    print(f"Number of regression rows: {len(regression_rows)}")


if __name__ == "__main__":
    main()
