#!/usr/bin/env python3
"""Run a rolling-average penetration model against an Excel workbook.

This script keeps the Excel-based calculation flow from the original snippet,
but fixes the brittle parts:

- validates workbook paths and row settings
- includes the final partial window back to the configured start row
- handles more date formats safely
- closes Excel/workbooks cleanly
- saves the output workbook next to the source workbook by default
- exposes the settings through a small CLI instead of hard-coded values
"""

from __future__ import annotations

import argparse
import sys
from dataclasses import dataclass
from datetime import date, datetime, time, timedelta
from pathlib import Path
from typing import Iterable, Optional


EXCEL_EPOCH = datetime(1899, 12, 30)


@dataclass(frozen=True)
class ModelConfig:
    sheet_name: str = "Empirical Model"
    param_cell: str = "J271"
    total_cell: str = "F265"
    max_cell: str = "F266"
    min_cell: str = "F267"
    date_column: str = "A"
    column_b: str = "B"
    column_c: str = "C"
    data_column: str = "I"
    actual_column: Optional[str] = None
    rows_per_quarter: int = 3
    start_row: int = 7
    last_row: int = 209


@dataclass(frozen=True)
class IterationResult:
    iteration: int
    date_value: object
    quarter: Optional[str]
    column_b: object
    column_c: object
    avg_penetration: object
    total_sold: object
    max_value: object
    min_value: object
    actual: object
    error: object
    error_pct: object

    def as_row(self) -> list[object]:
        return [
            self.iteration,
            self.date_value,
            self.quarter,
            self.column_b,
            self.column_c,
            self.avg_penetration,
            self.total_sold,
            self.max_value,
            self.min_value,
            self.actual,
            self.error,
            self.error_pct,
        ]


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


def validate_config(config: ModelConfig) -> None:
    if config.rows_per_quarter <= 0:
        raise ValueError("rows_per_quarter must be greater than 0")
    if config.start_row <= 0 or config.last_row <= 0:
        raise ValueError("start_row and last_row must be positive integers")
    if config.start_row > config.last_row:
        raise ValueError("start_row cannot be greater than last_row")


def iter_window_starts(start_row: int, last_row: int, rows_per_quarter: int) -> list[int]:
    """Return the rolling window start rows.

    The original code only emitted full 3-row steps and skipped the final
    partial window when the start row was not aligned to that step size.
    This version always includes the configured start row as the final window.
    """

    if rows_per_quarter <= 0:
        raise ValueError("rows_per_quarter must be greater than 0")
    if start_row > last_row:
        raise ValueError("start_row cannot be greater than last_row")

    starts: list[int] = []
    current_start = max(start_row, last_row - rows_per_quarter + 1)

    while current_start > start_row:
        starts.append(current_start)
        current_start -= rows_per_quarter

    starts.append(start_row)
    return starts


def coerce_excel_datetime(value: object) -> Optional[datetime]:
    if value in (None, ""):
        return None
    if isinstance(value, datetime):
        return value
    if isinstance(value, date):
        return datetime.combine(value, time.min)
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        return EXCEL_EPOCH + timedelta(days=float(value))
    if isinstance(value, str):
        stripped = value.strip()
        if not stripped:
            return None

        try:
            return datetime.fromisoformat(stripped)
        except ValueError:
            pass

        for fmt in (
            "%m/%d/%Y",
            "%m/%d/%y",
            "%d/%m/%Y",
            "%d/%m/%y",
            "%Y/%m/%d",
            "%Y-%m-%d",
        ):
            try:
                return datetime.strptime(stripped, fmt)
            except ValueError:
                continue

    return None


def detect_quarter(value: object) -> Optional[str]:
    parsed = coerce_excel_datetime(value)
    if parsed is None:
        return None if value in (None, "") else "Unknown"
    return f"Q{((parsed.month - 1) // 3) + 1}"


def has_value(value: object) -> bool:
    return value not in (None, "")


def build_output_path(
    file_path: Path,
    output_dir: Optional[Path] = None,
    generated_at: Optional[datetime] = None,
) -> Path:
    timestamp = (generated_at or datetime.now()).strftime("%Y%m%d_%H%M%S")
    target_dir = output_dir or file_path.parent
    return target_dir / f"{file_path.stem}_ENHANCED_{timestamp}.xlsx"


def build_summary_rows(
    file_path: Path,
    output_path: Path,
    config: ModelConfig,
    iteration_count: int,
    generated_at: datetime,
) -> list[list[object]]:
    return [
        ["RUN SUMMARY", ""],
        ["File", file_path.name],
        ["Source Path", str(file_path)],
        ["Output Path", str(output_path)],
        ["Sheet", config.sheet_name],
        ["Method", "Rolling average penetration model"],
        ["Rows per quarter", config.rows_per_quarter],
        ["Start row", config.start_row],
        ["Last row", config.last_row],
        ["Iterations", iteration_count],
        ["Parameter cell", config.param_cell],
        ["Data column", config.data_column],
        ["Actual column", config.actual_column or "Not provided"],
        ["Generated at", generated_at.strftime("%Y-%m-%d %H:%M:%S")],
    ]


def restore_parameter_cell(target_range: object, original_formula: object, original_value: object) -> None:
    if isinstance(original_formula, str) and original_formula.startswith("="):
        target_range.formula = original_formula
    else:
        target_range.value = original_value


def run_model(
    file_path: str | Path,
    config: ModelConfig,
    *,
    output_dir: Optional[str | Path] = None,
    visible: bool = False,
) -> Path:
    validate_config(config)

    source_path = Path(file_path).expanduser().resolve()
    if not source_path.exists():
        raise FileNotFoundError(f"Excel file not found: {source_path}")

    if sys.platform.startswith("linux"):
        raise RuntimeError(
            "xlwings needs a local Excel desktop app. Run this script on the "
            "Mac or Windows machine that has Microsoft Excel and the workbook."
        )

    try:
        import xlwings as xw
    except ImportError as exc:
        raise RuntimeError(
            "xlwings is not installed. Install it with 'pip install xlwings'."
        ) from exc

    target_dir = Path(output_dir).expanduser().resolve() if output_dir else source_path.parent
    target_dir.mkdir(parents=True, exist_ok=True)
    generated_at = datetime.now()
    output_path = build_output_path(source_path, target_dir, generated_at)

    with xw.App(visible=visible, add_book=False) as app:
        app.display_alerts = False
        try:
            app.screen_updating = False
        except Exception:
            pass
        try:
            app.calculation = "automatic"
        except Exception:
            pass

        workbook = None
        output_workbook = None
        parameter_range = None
        original_formula = None
        original_value = None

        try:
            workbook = app.books.open(str(source_path), update_links=False, read_only=False)
            worksheet = workbook.sheets[config.sheet_name]
            parameter_range = worksheet.range(config.param_cell)
            original_formula = parameter_range.formula
            original_value = parameter_range.value

            results: list[IterationResult] = []

            for iteration, first_row in enumerate(
                iter_window_starts(config.start_row, config.last_row, config.rows_per_quarter),
                start=1,
            ):
                range_used = f"{config.data_column}{first_row}:{config.data_column}{config.last_row}"
                raw_date = worksheet.range(f"{config.date_column}{first_row}").value

                parameter_range.formula = f"=AVERAGE({range_used})"
                app.calculate()

                total_sold = worksheet.range(config.total_cell).value
                max_value = worksheet.range(config.max_cell).value
                min_value = worksheet.range(config.min_cell).value
                actual = None
                error = None
                error_pct = None

                if config.actual_column:
                    actual = worksheet.range(f"{config.actual_column}{first_row}").value
                    if has_value(actual):
                        error = total_sold - actual
                        if actual != 0:
                            error_pct = error / actual

                results.append(
                    IterationResult(
                        iteration=iteration,
                        date_value=raw_date,
                        quarter=detect_quarter(raw_date),
                        column_b=worksheet.range(f"{config.column_b}{first_row}").value,
                        column_c=worksheet.range(f"{config.column_c}{first_row}").value,
                        avg_penetration=parameter_range.value,
                        total_sold=total_sold,
                        max_value=max_value,
                        min_value=min_value,
                        actual=actual,
                        error=error,
                        error_pct=error_pct,
                    )
                )

            restore_parameter_cell(parameter_range, original_formula, original_value)
            output_workbook = app.books.add()
            results_sheet = output_workbook.sheets[0]
            results_sheet.name = "Results"
            results_sheet.range("A1").value = RESULT_HEADERS
            results_sheet.range("A2").value = [result.as_row() for result in results]

            summary_sheet = output_workbook.sheets.add("Run_Summary", after=results_sheet)
            summary_sheet.range("A1").value = build_summary_rows(
                source_path,
                output_path,
                config,
                len(results),
                generated_at,
            )

            try:
                results_sheet.autofit()
                summary_sheet.autofit()
            except Exception:
                pass

            output_workbook.save(str(output_path))
            return output_path
        finally:
            if parameter_range is not None:
                try:
                    restore_parameter_cell(parameter_range, original_formula, original_value)
                except Exception:
                    pass
            if output_workbook is not None:
                output_workbook.close()
            if workbook is not None:
                workbook.close()


def parse_args(argv: Optional[Iterable[str]] = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("file_path", help="Path to the source Excel workbook")
    parser.add_argument(
        "--sheet-name",
        default=ModelConfig.sheet_name,
        help=f"Worksheet name to process (default: {ModelConfig.sheet_name})",
    )
    parser.add_argument("--param-cell", default=ModelConfig.param_cell)
    parser.add_argument("--total-cell", default=ModelConfig.total_cell)
    parser.add_argument("--max-cell", default=ModelConfig.max_cell)
    parser.add_argument("--min-cell", default=ModelConfig.min_cell)
    parser.add_argument("--date-column", default=ModelConfig.date_column)
    parser.add_argument("--column-b", default=ModelConfig.column_b)
    parser.add_argument("--column-c", default=ModelConfig.column_c)
    parser.add_argument("--data-column", default=ModelConfig.data_column)
    parser.add_argument("--actual-column", default=None)
    parser.add_argument("--rows-per-quarter", type=int, default=ModelConfig.rows_per_quarter)
    parser.add_argument("--start-row", type=int, default=ModelConfig.start_row)
    parser.add_argument("--last-row", type=int, default=ModelConfig.last_row)
    parser.add_argument(
        "--output-dir",
        default=None,
        help="Directory for the generated workbook. Defaults to the source workbook directory.",
    )
    parser.add_argument(
        "--visible",
        action="store_true",
        help="Show the Excel application while the script runs.",
    )
    return parser.parse_args(list(argv) if argv is not None else None)


def main(argv: Optional[Iterable[str]] = None) -> int:
    args = parse_args(argv)
    config = ModelConfig(
        sheet_name=args.sheet_name,
        param_cell=args.param_cell,
        total_cell=args.total_cell,
        max_cell=args.max_cell,
        min_cell=args.min_cell,
        date_column=args.date_column,
        column_b=args.column_b,
        column_c=args.column_c,
        data_column=args.data_column,
        actual_column=args.actual_column,
        rows_per_quarter=args.rows_per_quarter,
        start_row=args.start_row,
        last_row=args.last_row,
    )

    output_path = run_model(
        args.file_path,
        config,
        output_dir=args.output_dir,
        visible=args.visible,
    )
    print(f"Done - saved as {output_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
