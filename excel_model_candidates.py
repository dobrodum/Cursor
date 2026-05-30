#!/usr/bin/env python3
"""
Builds one parameter workbook from all .xlsx files in an input directory.

The script opens each source workbook exactly once, processes both
"Empirical Model" and "Regression Model" sheets while the workbook is open,
and then closes the source workbook without saving any changes.
"""

from __future__ import annotations

import calendar
import math
import re
import statistics
from dataclasses import dataclass
from datetime import date
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Sequence, Tuple

import xlwings as xw
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

# -------- User-configurable paths --------
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

DAY_MAP = {"early": 5, "mid": 15, "late": 25}
PERIOD_RE = re.compile(r"(Early|Mid|Late)\s*([A-Za-z]{3,9})\s*(\d{4})", re.IGNORECASE)


@dataclass
class FileModelInfo:
    model: str
    ticker: str
    model_period: str
    model_date: str


@dataclass
class SheetSnapshot:
    matrix: List[List[Any]]
    start_row: int
    start_col: int

    @property
    def end_row(self) -> int:
        return self.start_row + len(self.matrix) - 1

    @property
    def end_col(self) -> int:
        width = len(self.matrix[0]) if self.matrix else 0
        return self.start_col + width - 1

    def get_value(self, row: int, col: int) -> Any:
        if row < self.start_row or col < self.start_col:
            return None
        r_idx = row - self.start_row
        c_idx = col - self.start_col
        if r_idx < 0 or r_idx >= len(self.matrix):
            return None
        row_data = self.matrix[r_idx]
        if c_idx < 0 or c_idx >= len(row_data):
            return None
        return row_data[c_idx]


def to_float(value: Any) -> Optional[float]:
    if value is None:
        return None
    if isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        if isinstance(value, float) and math.isnan(value):
            return None
        return float(value)
    if isinstance(value, str):
        cleaned = value.strip().replace(",", "")
        if not cleaned:
            return None
        try:
            return float(cleaned)
        except ValueError:
            return None
    return None


def normalize_2d(values: Any) -> List[List[Any]]:
    if values is None:
        return []
    if not isinstance(values, list):
        return [[values]]
    if values and not isinstance(values[0], list):
        return [values]
    return values


def get_sheet_snapshot(sheet: xw.Sheet) -> SheetSnapshot:
    used = sheet.used_range
    matrix = normalize_2d(used.value)
    return SheetSnapshot(matrix=matrix, start_row=used.row, start_col=used.column)


def find_anchor(snapshot: SheetSnapshot, target: str = "max") -> Optional[Tuple[int, int]]:
    lowered = target.strip().lower()
    for r_idx, row in enumerate(snapshot.matrix):
        for c_idx, value in enumerate(row):
            if isinstance(value, str) and value.strip().lower() == lowered:
                return snapshot.start_row + r_idx, snapshot.start_col + c_idx
    return None


def month_number(month_token: str) -> Optional[int]:
    token = month_token.strip().lower()
    token = token[:3] if len(token) > 3 else token
    abbr_lookup = {calendar.month_abbr[i].lower(): i for i in range(1, 13)}
    if token in abbr_lookup:
        return abbr_lookup[token]
    full_lookup = {calendar.month_name[i].lower(): i for i in range(1, 13)}
    return full_lookup.get(month_token.strip().lower())


def parse_model_metadata(file_name: str) -> FileModelInfo:
    stem = Path(file_name).stem
    parts = [part.strip() for part in stem.split(" - ")]

    ticker = ""
    if len(parts) >= 2:
        ticker = parts[1]
    if not ticker:
        ticker_match = re.search(r"\b[A-Z]{2,8}\b", stem)
        ticker = ticker_match.group(0) if ticker_match else ""

    period_source = parts[2] if len(parts) >= 3 else stem
    period_source = period_source.split("_")[0]

    period_match = PERIOD_RE.search(period_source)
    model_period = ""
    model_date = ""

    if period_match:
        phase = period_match.group(1).title()
        month_token = period_match.group(2)
        year_text = period_match.group(3)
        month_num = month_number(month_token)
        if month_num is not None:
            month_abbr = calendar.month_abbr[month_num]
            model_period = f"{phase}{month_abbr}_{year_text}"
            day = DAY_MAP[phase.lower()]
            model_date = date(int(year_text), month_num, day).isoformat()

    model = f"{ticker}_{model_period}" if ticker and model_period else (ticker or stem)
    return FileModelInfo(model=model, ticker=ticker, model_period=model_period, model_date=model_date)


def next_output_path(input_path: Path, out_dir: Path) -> Path:
    base_name = f"{input_path.name}_PARAM.xlsx"
    base_path = out_dir / base_name
    if not base_path.exists():
        return base_path

    counter = 1
    while True:
        candidate = out_dir / f"{input_path.name}_PARAM.{counter}.xlsx"
        if not candidate.exists():
            return candidate
        counter += 1


def safe_close_workbook(wb: xw.Book) -> None:
    try:
        wb.close(save=False)
        return
    except TypeError:
        pass
    except Exception:
        pass

    try:
        wb.api.Close(SaveChanges=False)  # type: ignore[attr-defined]
        return
    except Exception:
        pass

    try:
        wb.close()
    except Exception:
        pass


def value_or_blank(value: Optional[float]) -> Any:
    return "" if value is None else value


def pct_or_blank(value: Optional[float]) -> Any:
    return "" if value is None else value * 100.0


def is_close_or_blank(left: Any, right: Any, tol: float = 1e-9) -> bool:
    left_num = to_float(left)
    right_num = to_float(right)
    if left_num is None and right_num is None:
        return True
    if left_num is None or right_num is None:
        return False
    return abs(left_num - right_num) <= tol


def build_data_points(
    snapshot: SheetSnapshot,
    anchor_row: int,
    x_col: int,
    y_col: int,
    quarter_col: int,
    require_nonzero_y: bool,
) -> List[Dict[str, Any]]:
    points: List[Dict[str, Any]] = []
    for row in range(snapshot.start_row, anchor_row):
        x_val = to_float(snapshot.get_value(row, x_col))
        y_val = to_float(snapshot.get_value(row, y_col))
        if x_val is None or y_val is None:
            continue
        if require_nonzero_y and y_val == 0:
            continue
        quarter_value = snapshot.get_value(row, quarter_col)
        points.append(
            {
                "row": row,
                "quarter_label": quarter_value if quarter_value is not None else "",
                "x": x_val,
                "y": y_val,
            }
        )
    return points


def extract_empirical_candidates(
    wb: xw.Book,
    info: FileModelInfo,
    source_file: str,
) -> List[Dict[str, Any]]:
    try:
        sheet = wb.sheets["Empirical Model"]
    except Exception:
        print(f"  skipped empirical: missing sheet 'Empirical Model' ({source_file})")
        return []

    snapshot = get_sheet_snapshot(sheet)
    anchor = find_anchor(snapshot, "max")
    if anchor is None:
        print(f"  skipped empirical: no 'max' anchor found ({source_file})")
        return []

    anchor_row, anchor_col = anchor
    quarterly_col = anchor_col - 11
    reported_col = anchor_col - 7
    quarter_label_col = quarterly_col - 1
    points = build_data_points(
        snapshot=snapshot,
        anchor_row=anchor_row,
        x_col=quarterly_col,
        y_col=reported_col,
        quarter_col=quarter_label_col,
        require_nonzero_y=True,
    )

    if len(points) < 2:
        print(f"  skipped empirical: not enough numeric history ({source_file})")
        return []

    n_quarters = min(10, len(points))
    helper_col = anchor_col + 20
    helper_start_row = anchor_row + 2

    for idx, n_used in enumerate(range(1, n_quarters + 1)):
        subset = points[-n_used:]
        subset_start_row = subset[0]["row"]
        subset_end_row = subset[-1]["row"]
        avg_formula = (
            f"=AVERAGE(IFERROR("
            f"R{subset_start_row}C{quarterly_col}:R{subset_end_row}C{quarterly_col}/"
            f"R{subset_start_row}C{reported_col}:R{subset_end_row}C{reported_col},\"\"))"
        )
        sheet.range((helper_start_row + idx, helper_col)).formula2 = avg_formula

    wb.app.calculate()

    avg_values = sheet.range(
        (helper_start_row, helper_col),
        (helper_start_row + n_quarters - 1, helper_col),
    ).value
    avg_values_2d = normalize_2d(avg_values)

    rows: List[Dict[str, Any]] = []
    for idx, n_used in enumerate(range(1, n_quarters + 1)):
        subset = points[-n_used:]
        ratios = [pt["x"] / pt["y"] for pt in subset if pt["y"]]
        avg_penetration = to_float(avg_values_2d[idx][0]) if idx < len(avg_values_2d) else None
        if avg_penetration is None and ratios:
            avg_penetration = statistics.fmean(ratios)

        current = subset[-1]
        current_quarterly = current["x"]
        current_reported = current["y"]

        forecast_value = (
            current_quarterly / avg_penetration if avg_penetration not in (None, 0.0) else None
        )
        ratio_min = min(ratios) if ratios else None
        ratio_max = max(ratios) if ratios else None

        forecast_max = (
            current_quarterly / ratio_min if ratio_min not in (None, 0.0) else None
        )
        forecast_min = (
            current_quarterly / ratio_max if ratio_max not in (None, 0.0) else None
        )
        range_width = (
            forecast_max - forecast_min
            if forecast_max is not None and forecast_min is not None
            else None
        )

        prev_quarterly = subset[-2]["x"] if len(subset) >= 2 else None
        growth_rate = (
            (current_quarterly - prev_quarterly) / prev_quarterly
            if prev_quarterly not in (None, 0.0)
            else None
        )
        sales_captured = (
            current_quarterly / current_reported if current_reported not in (None, 0.0) else None
        )

        rows.append(
            {
                "model": info.model,
                "ticker": info.ticker,
                "model_period": info.model_period,
                "model_date": info.model_date,
                "method": "empirical",
                "parameter_name": "avg_penetration_pct",
                "parameter_value": pct_or_blank(avg_penetration),
                "num_quarters_used": n_used,
                "last_quarter_used": subset[0]["quarter_label"],
                "forecast_value": value_or_blank(forecast_value),
                "actual_value": value_or_blank(current_reported),
                "forecast_max": value_or_blank(forecast_max),
                "forecast_min": value_or_blank(forecast_min),
                "range_width": value_or_blank(range_width),
                "avg_penetration_pct": pct_or_blank(avg_penetration),
                "quarterly_sales": value_or_blank(current_quarterly),
                "reported_sales": value_or_blank(current_reported),
                "growth_rate_pct": pct_or_blank(growth_rate),
                "sales_captured_in_db_pct": pct_or_blank(sales_captured),
                "source_file": source_file,
            }
        )

    return rows


def extract_regression_candidates(
    wb: xw.Book,
    info: FileModelInfo,
    source_file: str,
) -> List[Dict[str, Any]]:
    try:
        sheet = wb.sheets["Regression Model"]
    except Exception:
        print(f"  skipped regression: missing sheet 'Regression Model' ({source_file})")
        return []

    snapshot = get_sheet_snapshot(sheet)
    anchor = find_anchor(snapshot, "max")
    if anchor is None:
        print(f"  skipped regression: no 'max' anchor found ({source_file})")
        return []

    anchor_row, anchor_col = anchor
    y_col = anchor_col - 7
    x_col = anchor_col - 11
    quarter_label_col = x_col - 1

    points = build_data_points(
        snapshot=snapshot,
        anchor_row=anchor_row,
        x_col=x_col,
        y_col=y_col,
        quarter_col=quarter_label_col,
        require_nonzero_y=False,
    )
    if len(points) < 2:
        print(f"  skipped regression: not enough numeric history ({source_file})")
        return []

    max_quarters = min(10, len(points))
    quarter_range = list(range(2, max_quarters + 1))
    if not quarter_range:
        return []

    helper_col = anchor_col + 30
    helper_start_row = anchor_row + 2

    for idx, n_used in enumerate(quarter_range):
        subset = points[-n_used:]
        start_row = subset[0]["row"]
        end_row = subset[-1]["row"]

        intercept_formula = (
            f"=INTERCEPT(R{start_row}C{y_col}:R{end_row}C{y_col},"
            f"R{start_row}C{x_col}:R{end_row}C{x_col})"
        )
        slope_formula = (
            f"=SLOPE(R{start_row}C{y_col}:R{end_row}C{y_col},"
            f"R{start_row}C{x_col}:R{end_row}C{x_col})"
        )
        sheet.range((helper_start_row + idx, helper_col)).formula2 = intercept_formula
        sheet.range((helper_start_row + idx, helper_col + 1)).formula2 = slope_formula

    wb.app.calculate()

    intercept_values = normalize_2d(
        sheet.range(
            (helper_start_row, helper_col),
            (helper_start_row + len(quarter_range) - 1, helper_col),
        ).value
    )
    slope_values = normalize_2d(
        sheet.range(
            (helper_start_row, helper_col + 1),
            (helper_start_row + len(quarter_range) - 1, helper_col + 1),
        ).value
    )

    rows: List[Dict[str, Any]] = []
    for idx, n_used in enumerate(quarter_range):
        subset = points[-n_used:]
        intercept = to_float(intercept_values[idx][0]) if idx < len(intercept_values) else None
        slope = to_float(slope_values[idx][0]) if idx < len(slope_values) else None
        latest_x = subset[-1]["x"]

        forecast_total_without_sa = (
            intercept + (slope * latest_x) if intercept is not None and slope is not None else None
        )

        residual_std = None
        if intercept is not None and slope is not None and len(subset) > 1:
            residuals = [(pt["y"] - (intercept + slope * pt["x"])) for pt in subset]
            if len(residuals) > 1:
                residual_std = statistics.stdev(residuals)
            else:
                residual_std = 0.0

        forecast_max = (
            forecast_total_without_sa + residual_std
            if forecast_total_without_sa is not None and residual_std is not None
            else None
        )
        forecast_min = (
            forecast_total_without_sa - residual_std
            if forecast_total_without_sa is not None and residual_std is not None
            else None
        )
        range_width = (
            forecast_max - forecast_min
            if forecast_max is not None and forecast_min is not None
            else None
        )

        row = {
            "model": info.model,
            "ticker": info.ticker,
            "model_period": info.model_period,
            "model_date": info.model_date,
            "method": "regression",
            "parameter_name": "num_quarters_used",
            "parameter_value": n_used,
            "num_quarters_used": n_used,
            "forecast_value": value_or_blank(forecast_total_without_sa),
            "actual_value": "",
            "forecast_max": value_or_blank(forecast_max),
            "forecast_min": value_or_blank(forecast_min),
            "range_width": value_or_blank(range_width),
            "intercept": value_or_blank(intercept),
            "slope": value_or_blank(slope),
            "source_file": source_file,
        }

        if rows:
            prev = rows[-1]
            if (
                is_close_or_blank(prev["forecast_value"], row["forecast_value"])
                and is_close_or_blank(prev["forecast_max"], row["forecast_max"])
                and is_close_or_blank(prev["forecast_min"], row["forecast_min"])
                and is_close_or_blank(prev["intercept"], row["intercept"])
                and is_close_or_blank(prev["slope"], row["slope"])
            ):
                continue

        rows.append(row)

    return rows


def write_sheet(ws: Any, headers: Sequence[str], rows: Sequence[Dict[str, Any]]) -> None:
    ws.append(list(headers))
    for row in rows:
        ws.append([row.get(header, "") for header in headers])

    for cell in ws[1]:
        cell.font = Font(bold=True)

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions

    for idx, header in enumerate(headers, start=1):
        max_len = len(header)
        for row_idx in range(2, ws.max_row + 1):
            value = ws.cell(row=row_idx, column=idx).value
            value_len = len(str(value)) if value is not None else 0
            max_len = max(max_len, value_len)
        ws.column_dimensions[get_column_letter(idx)].width = min(max_len + 2, 48)


def save_output_workbook(
    output_path: Path,
    empirical_rows: Sequence[Dict[str, Any]],
    regression_rows: Sequence[Dict[str, Any]],
) -> None:
    wb = Workbook()
    ws_empirical = wb.active
    ws_empirical.title = "empirical_candidates"
    write_sheet(ws_empirical, EMPIRICAL_HEADERS, empirical_rows)

    ws_regression = wb.create_sheet("regression_candidates")
    write_sheet(ws_regression, REGRESSION_HEADERS, regression_rows)

    wb.save(output_path)


def iter_input_files(folder: Path) -> Iterable[Path]:
    for path in sorted(folder.iterdir()):
        if not path.is_file():
            print(f"skipped {path.name}: not a file")
            continue
        if path.name.startswith("~"):
            print(f"skipped {path.name}: temp file")
            continue
        if path.suffix.lower() != ".xlsx":
            print(f"skipped {path.name}: not an .xlsx file")
            continue
        if re.search(r"_PARAM(\.\d+)?$", path.stem, flags=re.IGNORECASE):
            print(f"skipped {path.name}: output workbook pattern")
            continue
        yield path


def main() -> None:
    in_path = Path(input_dir).expanduser().resolve()
    out_path = Path(output_dir).expanduser().resolve()

    if not in_path.exists() or not in_path.is_dir():
        raise FileNotFoundError(f"input_dir does not exist or is not a directory: {in_path}")

    out_path.mkdir(parents=True, exist_ok=True)
    output_file = next_output_path(in_path, out_path)

    empirical_rows: List[Dict[str, Any]] = []
    regression_rows: List[Dict[str, Any]] = []
    files_processed = 0

    app: Optional[xw.App] = None
    try:
        app = xw.App(visible=False, add_book=False)
        app.display_alerts = False
        app.screen_updating = False
        try:
            app.calculation = "manual"
        except Exception:
            pass

        for file_path in iter_input_files(in_path):
            print(f"processing {file_path.name}")
            wb = None
            try:
                wb = app.books.open(str(file_path), update_links=False)
                metadata = parse_model_metadata(file_path.name)

                empirical_rows.extend(
                    extract_empirical_candidates(
                        wb=wb,
                        info=metadata,
                        source_file=file_path.name,
                    )
                )
                regression_rows.extend(
                    extract_regression_candidates(
                        wb=wb,
                        info=metadata,
                        source_file=file_path.name,
                    )
                )
                files_processed += 1
            except Exception as exc:
                print(f"skipped {file_path.name}: failed to process ({exc})")
            finally:
                if wb is not None:
                    safe_close_workbook(wb)
    finally:
        if app is not None:
            try:
                app.quit()
            except Exception:
                pass

    save_output_workbook(output_file, empirical_rows, regression_rows)

    print(f"output path: {output_file}")
    print(f"files processed: {files_processed}")
    print(f"empirical rows: {len(empirical_rows)}")
    print(f"regression rows: {len(regression_rows)}")


if __name__ == "__main__":
    main()
