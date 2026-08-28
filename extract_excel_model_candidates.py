from __future__ import annotations

import math
import re
from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path
from typing import Any, Iterable

import xlwings as xw
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter


# Configure these two paths before running.
input_dir = "/workspace/input"
output_dir = "/workspace/output"


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


PERIOD_TO_DAY = {"Early": 5, "Mid": 15, "Late": 25}


@dataclass(frozen=True)
class FileLabels:
    model: str
    ticker: str
    model_period: str
    model_date: str


@dataclass(frozen=True)
class SheetSnapshot:
    start_row: int
    start_col: int
    end_row: int
    end_col: int
    values: list[list[Any]]

    def get(self, row: int, col: int) -> Any:
        if row < self.start_row or col < self.start_col:
            return None
        r_idx = row - self.start_row
        c_idx = col - self.start_col
        if r_idx < 0 or c_idx < 0:
            return None
        if r_idx >= len(self.values):
            return None
        row_values = self.values[r_idx]
        if c_idx >= len(row_values):
            return None
        value = row_values[c_idx]
        if value == "":
            return None
        return value


def normalize_2d(values: Any) -> list[list[Any]]:
    if values is None:
        return []
    if not isinstance(values, list):
        return [[values]]
    if not values:
        return []
    if isinstance(values[0], list):
        return values
    return [values]


def build_sheet_snapshot(sheet: xw.Sheet) -> SheetSnapshot:
    used = sheet.used_range
    values = normalize_2d(used.value)
    if not values:
        return SheetSnapshot(
            start_row=used.row,
            start_col=used.column,
            end_row=used.row,
            end_col=used.column,
            values=[],
        )
    max_cols = max(len(row) for row in values)
    normalized_rows: list[list[Any]] = []
    for row in values:
        padded = list(row) + [None] * (max_cols - len(row))
        normalized_rows.append(padded)
    return SheetSnapshot(
        start_row=used.row,
        start_col=used.column,
        end_row=used.row + len(normalized_rows) - 1,
        end_col=used.column + max_cols - 1,
        values=normalized_rows,
    )


def as_number(value: Any) -> float | None:
    if value is None or value == "":
        return None
    if isinstance(value, bool):
        return float(value)
    if isinstance(value, (int, float)):
        if isinstance(value, float) and math.isnan(value):
            return None
        return float(value)
    if isinstance(value, str):
        text = value.strip()
        if not text:
            return None
        text = text.replace(",", "")
        percent = text.endswith("%")
        if percent:
            text = text[:-1].strip()
        try:
            numeric = float(text)
        except ValueError:
            return None
        return numeric / 100 if percent else numeric
    return None


def as_int(value: Any) -> int | None:
    number = as_number(value)
    if number is None:
        return None
    return int(round(number))


def safe_value(value: Any) -> Any:
    if value == "":
        return None
    return value


def set_r1c1_formula2(cell: xw.Range, formula: str) -> None:
    try:
        cell.api.Formula2R1C1 = formula
        return
    except Exception:
        pass
    try:
        cell.formula2 = formula
        return
    except Exception:
        pass
    try:
        cell.api.FormulaR1C1 = formula
    except Exception:
        cell.formula = formula


def recalculate(workbook: xw.Book) -> None:
    try:
        workbook.app.calculate()
    except Exception:
        workbook.app.api.Calculate()


def safe_close_workbook(workbook: xw.Book) -> None:
    try:
        workbook.close(save=False)
        return
    except TypeError:
        try:
            workbook.close(False)
            return
        except Exception:
            pass
    except Exception:
        pass
    try:
        workbook.api.Close(SaveChanges=False)
        return
    except Exception:
        pass
    try:
        workbook.api.Close(False)
    except Exception:
        pass


def parse_file_labels(file_path: Path) -> FileLabels:
    stem = file_path.stem
    parts = [segment.strip() for segment in stem.split(" - ")]
    ticker = parts[1].upper() if len(parts) >= 2 and parts[1] else "UNKNOWN"

    period_source = parts[2] if len(parts) >= 3 else stem
    period_source = period_source.split("_")[0].strip()
    period_match = re.search(r"(Early|Mid|Late)([A-Za-z]{3,})(\d{4})", period_source, re.IGNORECASE)

    if period_match:
        period_part = period_match.group(1).title()
        month_abbrev = period_match.group(2)[:3].title()
        year = int(period_match.group(3))
        try:
            month_number = datetime.strptime(month_abbrev, "%b").month
            day = PERIOD_TO_DAY[period_part]
            model_period = f"{period_part}{month_abbrev}_{year}"
            model_date = date(year, month_number, day).isoformat()
        except ValueError:
            cleaned_period = re.sub(r"\s+", "_", period_source).strip("_")
            model_period = cleaned_period or "UnknownPeriod"
            model_date = ""
    else:
        cleaned_period = re.sub(r"\s+", "_", period_source).strip("_")
        model_period = cleaned_period or "UnknownPeriod"
        model_date = ""

    return FileLabels(
        model=f"{ticker}_{model_period}",
        ticker=ticker,
        model_period=model_period,
        model_date=model_date,
    )


def find_anchor(snapshot: SheetSnapshot, anchor_text: str = "max") -> tuple[int, int] | None:
    target = anchor_text.strip().lower()
    for r_idx, row in enumerate(snapshot.values):
        for c_idx, value in enumerate(row):
            if isinstance(value, str) and value.strip().lower() == target:
                return snapshot.start_row + r_idx, snapshot.start_col + c_idx
    return None


def extract_empirical_candidates(
    workbook: xw.Book,
    labels: FileLabels,
    source_file: str,
) -> list[dict[str, Any]]:
    try:
        sheet = workbook.sheets["Empirical Model"]
    except Exception:
        print(f"Skipped empirical extraction ({source_file}): sheet 'Empirical Model' not found")
        return []

    snapshot = build_sheet_snapshot(sheet)
    anchor = find_anchor(snapshot, "max")
    if anchor is None:
        print(f"Skipped empirical extraction ({source_file}): 'max' anchor not found")
        return []
    anchor_row, anchor_col = anchor

    # Offsets are anchored to the 'max' column to avoid repeated sheet scanning.
    offsets = {
        "sales_captured_in_db_pct": -8,
        "growth_rate_pct": -7,
        "quarterly_sales": -6,
        "num_quarters_used": -5,
        "last_quarter_used": -4,
        "avg_penetration_pct": -3,
        "reported_sales": -2,
        "forecast_value": -1,
        "forecast_max": 0,
        "forecast_min": 1,
    }

    first_data_row = anchor_row + 1
    n_quarters = 10

    helper_col = snapshot.end_col + 2
    penetration_source_col = anchor_col + offsets["sales_captured_in_db_pct"]
    source_rel_col = penetration_source_col - helper_col
    for idx in range(n_quarters):
        row = first_data_row + idx
        start_ref = f"R[{-idx}]C[{source_rel_col}]"
        end_ref = f"RC[{source_rel_col}]"
        formula = f'=IFERROR(AVERAGE({start_ref}:{end_ref}), "")'
        set_r1c1_formula2(sheet.range((row, helper_col)), formula)
    recalculate(workbook)

    helper_values = [
        safe_value(sheet.range((first_data_row + idx, helper_col)).value) for idx in range(n_quarters)
    ]
    sheet.range((first_data_row, helper_col), (first_data_row + n_quarters - 1, helper_col)).clear_contents()

    rows: list[dict[str, Any]] = []
    for idx in range(n_quarters):
        row = first_data_row + idx
        num_quarters = as_int(snapshot.get(row, anchor_col + offsets["num_quarters_used"])) or (idx + 1)
        last_quarter_used = safe_value(snapshot.get(row, anchor_col + offsets["last_quarter_used"]))
        forecast_value = safe_value(snapshot.get(row, anchor_col + offsets["forecast_value"]))
        actual_value = safe_value(snapshot.get(row, anchor_col + offsets["reported_sales"]))
        forecast_max = safe_value(snapshot.get(row, anchor_col + offsets["forecast_max"]))
        forecast_min = safe_value(snapshot.get(row, anchor_col + offsets["forecast_min"]))
        avg_penetration_existing = safe_value(snapshot.get(row, anchor_col + offsets["avg_penetration_pct"]))
        avg_penetration = (
            avg_penetration_existing if avg_penetration_existing is not None else helper_values[idx]
        )
        quarterly_sales = safe_value(snapshot.get(row, anchor_col + offsets["quarterly_sales"]))
        reported_sales = safe_value(snapshot.get(row, anchor_col + offsets["reported_sales"]))
        if avg_penetration is None:
            quarterly_sales_num = as_number(quarterly_sales)
            reported_sales_num = as_number(reported_sales)
            if quarterly_sales_num is not None and reported_sales_num not in (None, 0):
                avg_penetration = quarterly_sales_num / reported_sales_num
        growth_rate_pct = safe_value(snapshot.get(row, anchor_col + offsets["growth_rate_pct"]))
        sales_captured_in_db_pct = safe_value(
            snapshot.get(row, anchor_col + offsets["sales_captured_in_db_pct"])
        )

        if all(
            value is None
            for value in (
                forecast_value,
                actual_value,
                forecast_max,
                forecast_min,
                avg_penetration,
                quarterly_sales,
            )
        ):
            continue

        max_num = as_number(forecast_max)
        min_num = as_number(forecast_min)
        range_width = max_num - min_num if max_num is not None and min_num is not None else None

        rows.append(
            {
                "model": labels.model,
                "ticker": labels.ticker,
                "model_period": labels.model_period,
                "model_date": labels.model_date,
                "method": "empirical",
                "parameter_name": "avg_penetration_pct",
                "parameter_value": avg_penetration,
                "num_quarters_used": num_quarters,
                "last_quarter_used": last_quarter_used,
                "forecast_value": forecast_value,
                "actual_value": actual_value,
                "forecast_max": forecast_max,
                "forecast_min": forecast_min,
                "range_width": range_width,
                "avg_penetration_pct": avg_penetration,
                "quarterly_sales": quarterly_sales,
                "reported_sales": reported_sales,
                "growth_rate_pct": growth_rate_pct,
                "sales_captured_in_db_pct": sales_captured_in_db_pct,
                "source_file": source_file,
            }
        )
    return rows


def extract_regression_candidates(
    workbook: xw.Book,
    labels: FileLabels,
    source_file: str,
) -> list[dict[str, Any]]:
    try:
        sheet = workbook.sheets["Regression Model"]
    except Exception:
        print(f"Skipped regression extraction ({source_file}): sheet 'Regression Model' not found")
        return []

    snapshot = build_sheet_snapshot(sheet)
    anchor = find_anchor(snapshot, "max")
    if anchor is None:
        print(f"Skipped regression extraction ({source_file}): 'max' anchor not found")
        return []
    anchor_row, anchor_col = anchor

    # Required anchored source columns from existing model layout.
    y_col = anchor_col - 7
    x_col = anchor_col - 11

    history_rows: list[int] = []
    for row in range(anchor_row - 1, snapshot.start_row - 1, -1):
        y_val = as_number(snapshot.get(row, y_col))
        x_val = as_number(snapshot.get(row, x_col))
        if y_val is not None and x_val is not None:
            history_rows.append(row)
        elif history_rows:
            break
    history_rows.reverse()

    max_n = min(10, len(history_rows))
    if max_n < 2:
        print(f"Skipped regression extraction ({source_file}): not enough x/y history for INTERCEPT/SLOPE")
        return []

    offsets = {
        "num_quarters_used": -5,
        "actual_value": -2,
        "forecast_value": -1,  # TOT FCST w/o SA
        "forecast_max": 0,
        "forecast_min": 1,
    }

    first_data_row = anchor_row + 1
    intercept_col = snapshot.end_col + 2
    slope_col = snapshot.end_col + 3

    n_values = list(range(2, max_n + 1))
    for idx, n in enumerate(n_values):
        row = first_data_row + idx
        start_row = history_rows[-n]
        end_row = history_rows[-1]
        intercept_formula = (
            f'=IFERROR(INTERCEPT(R{start_row}C{y_col}:R{end_row}C{y_col}, '
            f'R{start_row}C{x_col}:R{end_row}C{x_col}), "")'
        )
        slope_formula = (
            f'=IFERROR(SLOPE(R{start_row}C{y_col}:R{end_row}C{y_col}, '
            f'R{start_row}C{x_col}:R{end_row}C{x_col}), "")'
        )
        set_r1c1_formula2(sheet.range((row, intercept_col)), intercept_formula)
        set_r1c1_formula2(sheet.range((row, slope_col)), slope_formula)
    recalculate(workbook)

    rows: list[dict[str, Any]] = []
    previous_signature: tuple[Any, ...] | None = None

    for idx, n in enumerate(n_values):
        row = first_data_row + idx

        num_quarters_used = as_int(snapshot.get(row, anchor_col + offsets["num_quarters_used"])) or n
        forecast_value = safe_value(snapshot.get(row, anchor_col + offsets["forecast_value"]))
        actual_value = safe_value(snapshot.get(row, anchor_col + offsets["actual_value"]))
        forecast_max = safe_value(snapshot.get(row, anchor_col + offsets["forecast_max"]))
        forecast_min = safe_value(snapshot.get(row, anchor_col + offsets["forecast_min"]))
        intercept = safe_value(sheet.range((row, intercept_col)).value)
        slope = safe_value(sheet.range((row, slope_col)).value)

        max_num = as_number(forecast_max)
        min_num = as_number(forecast_min)
        range_width = max_num - min_num if max_num is not None and min_num is not None else None

        signature = (
            num_quarters_used,
            as_number(forecast_value),
            as_number(forecast_max),
            as_number(forecast_min),
            as_number(intercept),
            as_number(slope),
        )
        if previous_signature is not None and signature == previous_signature:
            continue
        previous_signature = signature

        rows.append(
            {
                "model": labels.model,
                "ticker": labels.ticker,
                "model_period": labels.model_period,
                "model_date": labels.model_date,
                "method": "regression",
                "parameter_name": "num_quarters_used",
                "parameter_value": num_quarters_used,
                "num_quarters_used": num_quarters_used,
                "forecast_value": forecast_value,
                "actual_value": actual_value,
                "forecast_max": forecast_max,
                "forecast_min": forecast_min,
                "range_width": range_width,
                "intercept": intercept,
                "slope": slope,
                "source_file": source_file,
            }
        )

    if n_values:
        last_formula_row = first_data_row + len(n_values) - 1
        sheet.range((first_data_row, intercept_col), (last_formula_row, slope_col)).clear_contents()

    return rows


def choose_output_path(input_path: Path, output_path: Path) -> Path:
    stem = f"{input_path.name}_PARAM"
    candidate = output_path / f"{stem}.xlsx"
    if not candidate.exists():
        return candidate
    suffix = 1
    while True:
        candidate = output_path / f"{stem}.{suffix}.xlsx"
        if not candidate.exists():
            return candidate
        suffix += 1


def write_sheet(
    workbook: Workbook,
    name: str,
    headers: list[str],
    rows: Iterable[dict[str, Any]],
) -> None:
    ws = workbook.create_sheet(title=name)
    ws.append(headers)
    for row in rows:
        ws.append([row.get(header) for header in headers])

    for cell in ws[1]:
        cell.font = Font(bold=True)

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(ws.max_column)}{ws.max_row}"

    for col_idx in range(1, ws.max_column + 1):
        header = ws.cell(row=1, column=col_idx).value
        max_len = len(str(header)) if header is not None else 0
        for row_idx in range(2, ws.max_row + 1):
            value = ws.cell(row=row_idx, column=col_idx).value
            if value is None:
                continue
            max_len = max(max_len, len(str(value)))
        ws.column_dimensions[get_column_letter(col_idx)].width = min(max(max_len + 2, 12), 44)


def run() -> None:
    input_path = Path(input_dir).expanduser().resolve()
    output_path = Path(output_dir).expanduser().resolve()
    output_path.mkdir(parents=True, exist_ok=True)

    if not input_path.exists():
        raise FileNotFoundError(f"input_dir does not exist: {input_path}")
    if not input_path.is_dir():
        raise NotADirectoryError(f"input_dir is not a directory: {input_path}")

    empirical_rows: list[dict[str, Any]] = []
    regression_rows: list[dict[str, Any]] = []
    processed_files = 0

    app = xw.App(visible=False, add_book=False)
    try:
        for attr, value in (
            ("display_alerts", False),
            ("screen_updating", False),
            ("enable_events", False),
        ):
            try:
                setattr(app, attr, value)
            except Exception:
                pass
        try:
            app.calculation = "manual"
        except Exception:
            pass

        for file_path in sorted(input_path.iterdir()):
            if not file_path.is_file():
                continue
            if file_path.name.startswith("~"):
                print(f"Skipped file: {file_path.name} (temporary file)")
                continue
            if file_path.suffix.lower() != ".xlsx":
                print(f"Skipped file: {file_path.name} (not .xlsx)")
                continue

            workbook = None
            try:
                workbook = app.books.open(str(file_path), update_links=False)
                labels = parse_file_labels(file_path)

                empirical_rows.extend(
                    extract_empirical_candidates(
                        workbook=workbook,
                        labels=labels,
                        source_file=file_path.name,
                    )
                )
                regression_rows.extend(
                    extract_regression_candidates(
                        workbook=workbook,
                        labels=labels,
                        source_file=file_path.name,
                    )
                )
                processed_files += 1
                print(f"Processed file: {file_path.name}")
            except Exception as exc:
                print(f"Skipped file: {file_path.name} (error: {exc})")
            finally:
                if workbook is not None:
                    safe_close_workbook(workbook)
    finally:
        app.quit()

    destination = choose_output_path(input_path=input_path, output_path=output_path)
    out_wb = Workbook()
    default_sheet = out_wb.active
    out_wb.remove(default_sheet)
    write_sheet(out_wb, "empirical_candidates", EMPIRICAL_HEADERS, empirical_rows)
    write_sheet(out_wb, "regression_candidates", REGRESSION_HEADERS, regression_rows)
    out_wb.save(destination)

    print(f"Output path: {destination}")
    print(f"Number of files processed: {processed_files}")
    print(f"Number of empirical rows: {len(empirical_rows)}")
    print(f"Number of regression rows: {len(regression_rows)}")


if __name__ == "__main__":
    run()
