from __future__ import annotations

from dataclasses import dataclass
from datetime import date
from pathlib import Path
import re
from typing import Any, Dict, Iterable, List, Optional, Sequence, Tuple

import xlwings as xw
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter


# -------------------------
# User-configurable paths
# -------------------------
input_dir = Path("/workspace/input")
output_dir = Path("/workspace/output")


EMPIRICAL_COLUMNS: List[str] = [
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

REGRESSION_COLUMNS: List[str] = [
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

MONTH_MAP = {
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

PERIOD_DAY = {
    "early": 5,
    "mid": 15,
    "late": 25,
}


@dataclass
class FileMetadata:
    model: str
    ticker: str
    model_period: str
    model_date: str


@dataclass
class SheetSnapshot:
    values: List[List[Any]]
    base_row: int
    base_col: int


def to_2d(values: Any) -> List[List[Any]]:
    if values is None:
        return []
    if isinstance(values, tuple):
        values = list(values)
    if not isinstance(values, list):
        return [[values]]
    if not values:
        return []
    if isinstance(values[0], tuple):
        return [list(row) if isinstance(row, tuple) else row for row in values]
    if not isinstance(values[0], list):
        return [values]
    return values


def normalize_label(value: Any) -> str:
    if value is None:
        return ""
    text = str(value).strip().lower()
    text = re.sub(r"[\s/\-]+", "_", text)
    text = re.sub(r"[^a-z0-9_]", "", text)
    return text


def get_unique_output_path(in_dir: Path, out_dir: Path) -> Path:
    out_dir.mkdir(parents=True, exist_ok=True)
    base_name = f"{in_dir.name}_PARAM"
    candidate = out_dir / f"{base_name}.xlsx"
    if not candidate.exists():
        return candidate

    idx = 1
    while True:
        candidate = out_dir / f"{base_name}.{idx}.xlsx"
        if not candidate.exists():
            return candidate
        idx += 1


def parse_filename_metadata(file_name: str) -> FileMetadata:
    stem = Path(file_name).stem
    parts = [p.strip() for p in stem.split("-")]
    ticker = "UNKNOWN"
    if len(parts) >= 2 and parts[1]:
        ticker = re.sub(r"[^A-Za-z0-9]", "", parts[1]).upper() or "UNKNOWN"

    period_match = re.search(
        r"(Early|Mid|Late)\s*([A-Za-z]{3,9})\s*[_\-]?\s*(\d{4})",
        stem,
        flags=re.IGNORECASE,
    )
    model_period = "UnknownPeriod"
    model_date = ""
    if period_match:
        period_token = period_match.group(1).strip().lower()
        month_token = period_match.group(2).strip().lower()[:3]
        year = int(period_match.group(3))
        day = PERIOD_DAY.get(period_token, 15)
        month_num = MONTH_MAP.get(month_token, 1)
        month_name = month_token.title()
        period_prefix = period_token.title()
        model_period = f"{period_prefix}{month_name}_{year}"
        model_date = date(year, month_num, day).isoformat()

    model = f"{ticker}_{model_period}"
    return FileMetadata(
        model=model,
        ticker=ticker,
        model_period=model_period,
        model_date=model_date,
    )


def safe_float(value: Any) -> Optional[float]:
    if value is None or value == "":
        return None
    if isinstance(value, bool):
        return float(value)
    if isinstance(value, (int, float)):
        return float(value)
    text = str(value).strip().replace(",", "")
    if text.endswith("%"):
        try:
            return float(text[:-1]) / 100.0
        except ValueError:
            return None
    try:
        return float(text)
    except ValueError:
        return None


def safe_value(value: Any) -> Any:
    if isinstance(value, str):
        value = value.strip()
        return value if value else None
    return value


def write_formula_r1c1(cell: xw.main.Range, formula_r1c1: str) -> None:
    try:
        cell.formula2 = formula_r1c1
    except Exception:
        try:
            cell.api.Formula2R1C1 = formula_r1c1
        except Exception:
            cell.api.FormulaR1C1 = formula_r1c1


def close_workbook_safe(wb: xw.main.Book) -> None:
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
    except Exception:
        # Last-resort close. Workbook is never saved in this workflow.
        wb.close()


def capture_snapshot(sheet: xw.main.Sheet) -> SheetSnapshot:
    used = sheet.used_range
    values = to_2d(used.value)
    return SheetSnapshot(values=values, base_row=used.row, base_col=used.column)


def find_anchor(snapshot: SheetSnapshot, label: str = "max") -> Optional[Tuple[int, int]]:
    target = normalize_label(label)
    best: Optional[Tuple[int, int, int, int]] = None  # (score, abs_row, abs_col, idx)
    idx = 0
    for r_idx, row in enumerate(snapshot.values):
        for c_idx, val in enumerate(row):
            if normalize_label(val) != target:
                idx += 1
                continue
            score = 0
            neighbors: List[Tuple[int, int]] = [
                (r_idx, c_idx + 1),
                (r_idx, c_idx - 1),
                (r_idx - 1, c_idx),
                (r_idx + 1, c_idx),
            ]
            for nr, nc in neighbors:
                if nr < 0 or nr >= len(snapshot.values):
                    continue
                if nc < 0 or nc >= len(snapshot.values[nr]):
                    continue
                neighbor_norm = normalize_label(snapshot.values[nr][nc])
                if neighbor_norm == "min":
                    score += 2
                if "forecast" in neighbor_norm:
                    score += 1
            abs_row = snapshot.base_row + r_idx
            abs_col = snapshot.base_col + c_idx
            candidate = (score, abs_row, abs_col, idx)
            if best is None or candidate > best:
                best = candidate
            idx += 1
    if best is None:
        return None
    return best[1], best[2]


def build_header_map(snapshot: SheetSnapshot, header_row_abs: int) -> Dict[str, int]:
    row_idx = header_row_abs - snapshot.base_row
    if row_idx < 0 or row_idx >= len(snapshot.values):
        return {}

    row_values = snapshot.values[row_idx]
    header_map: Dict[str, int] = {}
    for c_idx, value in enumerate(row_values):
        key = normalize_label(value)
        if not key:
            continue
        header_map[key] = snapshot.base_col + c_idx
    return header_map


def resolve_col(
    header_map: Dict[str, int],
    aliases: Sequence[str],
    anchor_col: int,
    default_offset: int,
) -> int:
    for alias in aliases:
        key = normalize_label(alias)
        if key in header_map:
            return header_map[key]
    return anchor_col + default_offset


def is_numeric(value: Any) -> bool:
    return safe_float(value) is not None


def find_contiguous_numeric_rows(
    sheet: xw.main.Sheet,
    data_col: int,
    end_row: int,
    max_scan: int = 120,
) -> Tuple[int, int]:
    if end_row < 1:
        return 1, 0

    r = end_row
    scanned = 0
    while r >= 1 and scanned < max_scan:
        v = sheet.range((r, data_col)).value
        if not is_numeric(v):
            break
        r -= 1
        scanned += 1
    start_row = r + 1
    return start_row, end_row


def pull_row_values(sheet: xw.main.Sheet, row: int, col_map: Dict[str, int]) -> Dict[str, Any]:
    out: Dict[str, Any] = {}
    for key, col in col_map.items():
        out[key] = safe_value(sheet.range((row, col)).value)
    return out


def extract_empirical_rows(
    wb: xw.main.Book,
    metadata: FileMetadata,
    source_file: str,
) -> List[Dict[str, Any]]:
    if "Empirical Model" not in [s.name for s in wb.sheets]:
        print(f"  - skipped empirical (missing sheet): {source_file}")
        return []

    sheet = wb.sheets["Empirical Model"]
    snapshot = capture_snapshot(sheet)
    anchor = find_anchor(snapshot, label="max")
    if anchor is None:
        print(f"  - skipped empirical (no max anchor): {source_file}")
        return []

    anchor_row, anchor_col = anchor
    header_map = build_header_map(snapshot, anchor_row)

    col_map = {
        "num_quarters_used": resolve_col(
            header_map,
            aliases=("num_quarters_used", "quarters_used", "n_quarters"),
            anchor_col=anchor_col,
            default_offset=-9,
        ),
        "last_quarter_used": resolve_col(
            header_map,
            aliases=("last_quarter_used", "last_quarter", "last_qtr"),
            anchor_col=anchor_col,
            default_offset=-8,
        ),
        "forecast_value": resolve_col(
            header_map,
            aliases=("estimated_total_sold", "tot_fcst", "forecast_value"),
            anchor_col=anchor_col,
            default_offset=-2,
        ),
        "actual_value": resolve_col(
            header_map,
            aliases=("reported_sales", "actual_value", "actual"),
            anchor_col=anchor_col,
            default_offset=-1,
        ),
        "forecast_max": resolve_col(
            header_map,
            aliases=("max", "forecast_max"),
            anchor_col=anchor_col,
            default_offset=0,
        ),
        "forecast_min": resolve_col(
            header_map,
            aliases=("min", "forecast_min"),
            anchor_col=anchor_col,
            default_offset=1,
        ),
        "avg_penetration_pct": resolve_col(
            header_map,
            aliases=("avg_penetration_pct", "avg_penetration", "avg_pen"),
            anchor_col=anchor_col,
            default_offset=-6,
        ),
        "quarterly_sales": resolve_col(
            header_map,
            aliases=("quarterly_sales", "qtr_sales"),
            anchor_col=anchor_col,
            default_offset=-12,
        ),
        "reported_sales": resolve_col(
            header_map,
            aliases=("reported_sales",),
            anchor_col=anchor_col,
            default_offset=-1,
        ),
        "growth_rate_pct": resolve_col(
            header_map,
            aliases=("growth_rate_pct", "growth_rate", "growth_pct"),
            anchor_col=anchor_col,
            default_offset=-5,
        ),
        "sales_captured_in_db_pct": resolve_col(
            header_map,
            aliases=("sales_captured_in_db_pct", "captured_in_db_pct", "penetration_pct"),
            anchor_col=anchor_col,
            default_offset=-7,
        ),
    }

    n_quarters = 10
    first_output_row = anchor_row + 1

    # Compute avg penetration with R1C1 formula2 using trailing quarters in sales-captured column.
    hist_start, hist_end = find_contiguous_numeric_rows(
        sheet=sheet,
        data_col=col_map["sales_captured_in_db_pct"],
        end_row=anchor_row - 1,
    )
    hist_count = max(0, hist_end - hist_start + 1)

    formulas_written = False
    if hist_count > 0:
        for i in range(n_quarters):
            row = first_output_row + i
            used = min(i + 1, hist_count)
            start_row = hist_end - used + 1

            # Keep num_quarters_used explicit in the candidate table row.
            sheet.range((row, col_map["num_quarters_used"])).value = i + 1
            avg_cell = sheet.range((row, col_map["avg_penetration_pct"]))
            avg_formula = (
                f"=AVERAGE(R{start_row}C{col_map['sales_captured_in_db_pct']}:"
                f"R{hist_end}C{col_map['sales_captured_in_db_pct']})"
            )
            write_formula_r1c1(avg_cell, avg_formula)
            formulas_written = True

    if formulas_written:
        wb.app.calculate()

    rows: List[Dict[str, Any]] = []
    for i in range(n_quarters):
        row_num = first_output_row + i
        pulled = pull_row_values(sheet, row_num, col_map)

        forecast_max = safe_float(pulled["forecast_max"])
        forecast_min = safe_float(pulled["forecast_min"])
        range_width = None
        if forecast_max is not None and forecast_min is not None:
            range_width = forecast_max - forecast_min

        avg_pen = safe_float(pulled["avg_penetration_pct"])
        forecast_val = safe_float(pulled["forecast_value"])
        actual_val = safe_float(pulled["actual_value"])

        # Drop empty lines with no signal from workbook.
        if (
            safe_float(pulled["num_quarters_used"]) is None
            and forecast_val is None
            and forecast_max is None
            and forecast_min is None
            and avg_pen is None
        ):
            continue

        out = {
            "model": metadata.model,
            "ticker": metadata.ticker,
            "model_period": metadata.model_period,
            "model_date": metadata.model_date,
            "method": "empirical",
            "parameter_name": "avg_penetration_pct",
            "parameter_value": avg_pen,
            "num_quarters_used": pulled["num_quarters_used"],
            "last_quarter_used": pulled["last_quarter_used"],
            "forecast_value": forecast_val,
            "actual_value": actual_val,
            "forecast_max": forecast_max,
            "forecast_min": forecast_min,
            "range_width": range_width,
            "avg_penetration_pct": avg_pen,
            "quarterly_sales": safe_float(pulled["quarterly_sales"]),
            "reported_sales": safe_float(pulled["reported_sales"]),
            "growth_rate_pct": safe_float(pulled["growth_rate_pct"]),
            "sales_captured_in_db_pct": safe_float(pulled["sales_captured_in_db_pct"]),
            "source_file": source_file,
        }
        rows.append(out)

    return rows


def extract_regression_rows(
    wb: xw.main.Book,
    metadata: FileMetadata,
    source_file: str,
) -> List[Dict[str, Any]]:
    if "Regression Model" not in [s.name for s in wb.sheets]:
        print(f"  - skipped regression (missing sheet): {source_file}")
        return []

    sheet = wb.sheets["Regression Model"]
    snapshot = capture_snapshot(sheet)
    anchor = find_anchor(snapshot, label="max")
    if anchor is None:
        print(f"  - skipped regression (no max anchor): {source_file}")
        return []

    anchor_row, anchor_col = anchor
    header_map = build_header_map(snapshot, anchor_row)

    # Required anchor-based offsets.
    y_col = anchor_col - 7
    x_col = anchor_col - 11

    num_quarters_col = resolve_col(
        header_map,
        aliases=("num_quarters_used", "quarters_used", "n_quarters"),
        anchor_col=anchor_col,
        default_offset=-9,
    )
    forecast_col = resolve_col(
        header_map,
        aliases=("tot_fcst_wo_sa", "tot_fcst_w_o_sa", "tot_fcst_without_sa", "forecast_value"),
        anchor_col=anchor_col,
        default_offset=-2,
    )
    max_col = resolve_col(
        header_map,
        aliases=("max", "forecast_max"),
        anchor_col=anchor_col,
        default_offset=0,
    )
    min_col = resolve_col(
        header_map,
        aliases=("min", "forecast_min"),
        anchor_col=anchor_col,
        default_offset=1,
    )
    actual_col = resolve_col(
        header_map,
        aliases=("actual_value", "actual", "reported_sales"),
        anchor_col=anchor_col,
        default_offset=-1,
    )

    # If intercept/slope columns exist in-table, use them. Otherwise use helper columns.
    intercept_target_col = resolve_col(
        header_map,
        aliases=("intercept",),
        anchor_col=anchor_col,
        default_offset=3,
    )
    slope_target_col = resolve_col(
        header_map,
        aliases=("slope",),
        anchor_col=anchor_col,
        default_offset=4,
    )

    hist_start, hist_end = find_contiguous_numeric_rows(
        sheet=sheet,
        data_col=y_col,
        end_row=anchor_row - 1,
    )
    hist_count = max(0, hist_end - hist_start + 1)
    n_quarters = min(10, hist_count) if hist_count > 0 else 10
    first_output_row = anchor_row + 1

    # Write intercept/slope formulas with R1C1 .formula2, then calculate once.
    formulas_written = False
    for i in range(n_quarters):
        row = first_output_row + i
        used = i + 1
        if hist_count > 0:
            used = min(used, hist_count)
            start_row = hist_end - used + 1
            y_range = f"R{start_row}C{y_col}:R{hist_end}C{y_col}"
            x_range = f"R{start_row}C{x_col}:R{hist_end}C{x_col}"
            write_formula_r1c1(
                sheet.range((row, intercept_target_col)),
                f"=INTERCEPT({y_range},{x_range})",
            )
            write_formula_r1c1(
                sheet.range((row, slope_target_col)),
                f"=SLOPE({y_range},{x_range})",
            )
            formulas_written = True
        sheet.range((row, num_quarters_col)).value = used

    if formulas_written:
        wb.app.calculate()

    rows: List[Dict[str, Any]] = []
    prev_signature: Optional[Tuple[Optional[float], ...]] = None
    for i in range(n_quarters):
        row = first_output_row + i
        num_q = safe_float(sheet.range((row, num_quarters_col)).value)
        intercept_val = safe_float(sheet.range((row, intercept_target_col)).value)
        slope_val = safe_float(sheet.range((row, slope_target_col)).value)
        forecast_val = safe_float(sheet.range((row, forecast_col)).value)
        forecast_max = safe_float(sheet.range((row, max_col)).value)
        forecast_min = safe_float(sheet.range((row, min_col)).value)
        actual_val = safe_float(sheet.range((row, actual_col)).value)

        # Fallbacks if sheet doesn't expose explicit rows for these fields.
        if hist_count > 0 and (forecast_max is None or forecast_min is None):
            used = min(i + 1, hist_count)
            start_row = hist_end - used + 1
            y_values = sheet.range((start_row, y_col), (hist_end, y_col)).value
            y_vals = [safe_float(v) for v in to_2d(y_values)[0] if safe_float(v) is not None]
            if y_vals:
                if forecast_max is None:
                    forecast_max = max(y_vals)
                if forecast_min is None:
                    forecast_min = min(y_vals)

        if (
            forecast_val is None
            and intercept_val is not None
            and slope_val is not None
            and hist_count > 0
        ):
            x_last = safe_float(sheet.range((hist_end, x_col)).value)
            if x_last is not None:
                forecast_val = intercept_val + slope_val * (x_last + 1)

        range_width = None
        if forecast_max is not None and forecast_min is not None:
            range_width = forecast_max - forecast_min

        if (
            num_q is None
            and forecast_val is None
            and intercept_val is None
            and slope_val is None
            and forecast_max is None
            and forecast_min is None
        ):
            continue

        signature = (
            num_q,
            round(intercept_val, 10) if intercept_val is not None else None,
            round(slope_val, 10) if slope_val is not None else None,
            round(forecast_val, 10) if forecast_val is not None else None,
            round(forecast_max, 10) if forecast_max is not None else None,
            round(forecast_min, 10) if forecast_min is not None else None,
        )

        # Duplicate-final-row guard: skip if final row matches the previous calculation.
        is_final = i == (n_quarters - 1)
        if is_final and prev_signature is not None and signature == prev_signature:
            continue
        prev_signature = signature

        rows.append(
            {
                "model": metadata.model,
                "ticker": metadata.ticker,
                "model_period": metadata.model_period,
                "model_date": metadata.model_date,
                "method": "regression",
                "parameter_name": "num_quarters_used",
                "parameter_value": num_q,
                "num_quarters_used": num_q,
                "forecast_value": forecast_val,
                "actual_value": actual_val,
                "forecast_max": forecast_max,
                "forecast_min": forecast_min,
                "range_width": range_width,
                "intercept": intercept_val,
                "slope": slope_val,
                "source_file": source_file,
            }
        )

    return rows


def style_output_sheet(ws, headers: Sequence[str], rows: List[Dict[str, Any]]) -> None:
    ws.append(list(headers))
    for cell in ws[1]:
        cell.font = Font(bold=True)

    for row in rows:
        ws.append([row.get(col) for col in headers])

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions

    for col_idx, header in enumerate(headers, start=1):
        max_len = len(header)
        for value in ws.iter_cols(
            min_col=col_idx,
            max_col=col_idx,
            min_row=2,
            max_row=ws.max_row,
            values_only=True,
        ):
            for cell_value in value:
                if cell_value is None:
                    continue
                max_len = max(max_len, len(str(cell_value)))
        ws.column_dimensions[get_column_letter(col_idx)].width = min(max_len + 2, 60)


def collect_input_files(in_dir: Path) -> List[Path]:
    files: List[Path] = []
    if not in_dir.exists():
        print(f"skipped input directory (not found): {in_dir}")
        return files
    if not in_dir.is_dir():
        print(f"skipped input path (not a directory): {in_dir}")
        return files

    for path in sorted(in_dir.iterdir(), key=lambda p: p.name.lower()):
        if not path.is_file():
            print(f"skipped file: {path.name} (not a file)")
            continue
        if path.name.startswith("~"):
            print(f"skipped file: {path.name} (temporary file)")
            continue
        if path.suffix.lower() != ".xlsx":
            print(f"skipped file: {path.name} (not .xlsx)")
            continue
        files.append(path)
    return files


def write_output_workbook(
    out_path: Path,
    empirical_rows: List[Dict[str, Any]],
    regression_rows: List[Dict[str, Any]],
) -> None:
    wb = Workbook()
    default_sheet = wb.active
    wb.remove(default_sheet)

    empirical_ws = wb.create_sheet("empirical_candidates")
    regression_ws = wb.create_sheet("regression_candidates")

    style_output_sheet(empirical_ws, EMPIRICAL_COLUMNS, empirical_rows)
    style_output_sheet(regression_ws, REGRESSION_COLUMNS, regression_rows)

    wb.save(out_path)


def main() -> None:
    files = collect_input_files(input_dir)
    out_path = get_unique_output_path(input_dir, output_dir)

    empirical_rows: List[Dict[str, Any]] = []
    regression_rows: List[Dict[str, Any]] = []
    processed_count = 0

    app = xw.App(visible=False, add_book=False)
    app.display_alerts = False
    app.screen_updating = False

    try:
        for file_path in files:
            print(f"processed file: {file_path.name}")
            wb: Optional[xw.main.Book] = None
            try:
                # Required safe open mode.
                wb = app.books.open(str(file_path), update_links=False)
                meta = parse_filename_metadata(file_path.name)

                empirical_rows.extend(
                    extract_empirical_rows(
                        wb=wb,
                        metadata=meta,
                        source_file=file_path.name,
                    )
                )
                regression_rows.extend(
                    extract_regression_rows(
                        wb=wb,
                        metadata=meta,
                        source_file=file_path.name,
                    )
                )
                processed_count += 1
            except Exception as exc:
                print(f"skipped file: {file_path.name} (error: {exc})")
            finally:
                if wb is not None:
                    close_workbook_safe(wb)
    finally:
        app.quit()

    write_output_workbook(
        out_path=out_path,
        empirical_rows=empirical_rows,
        regression_rows=regression_rows,
    )

    print(f"output path: {out_path}")
    print(f"number of files processed: {processed_count}")
    print(f"number of empirical rows: {len(empirical_rows)}")
    print(f"number of regression rows: {len(regression_rows)}")


if __name__ == "__main__":
    main()
