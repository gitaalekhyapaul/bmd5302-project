from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
import re
import shutil
import sys
import time
import traceback
from typing import Iterable

import pandas as pd
from PIL import ImageGrab
import xlwings as xw
from xlwings import XlwingsError

try:
    from xlwings._xlmac import kw as mac_kw
    from xlwings._xlmac import posix_to_hfs_path
except ImportError:  # pragma: no cover - non-macOS backend
    mac_kw = None
    posix_to_hfs_path = None


@dataclass(frozen=True, slots=True)
class WorkbookContract:
    sheet_name: str = "NormalTest"
    input_cell: str = "B1"
    chart_name: str = "NormalDataChart"
    macro_name: str = "GenerateNormalData"
    sample_table_header_range: str = "C1:D1"
    sample_table_start_row: int = 2
    sample_table_start_column: str = "C"
    sample_table_end_column: str = "D"

    def sample_table_range(self, sample_count: int) -> str:
        last_row = self.sample_table_start_row + sample_count - 1
        return (
            f"{self.sample_table_start_column}{self.sample_table_start_row}:"
            f"{self.sample_table_end_column}{last_row}"
        )


@dataclass(frozen=True, slots=True)
class WorkflowPaths:
    workbook_path: Path
    output_dir: Path
    workbook_output_dir: Path
    chart_output_dir: Path
    workbook_copy: Path

    @classmethod
    def create(
        cls,
        workbook_path: str | Path = "Test.xlsm",
        output_dir: str | Path = "notebook_outputs",
    ) -> WorkflowPaths:
        resolved_workbook_path = Path(workbook_path).expanduser().resolve()
        resolved_output_dir = Path(output_dir).expanduser().resolve()
        workbook_output_dir = resolved_output_dir / "workbooks"
        chart_output_dir = resolved_output_dir / "charts"
        workbook_output_dir.mkdir(parents=True, exist_ok=True)
        chart_output_dir.mkdir(parents=True, exist_ok=True)

        return cls(
            workbook_path=resolved_workbook_path,
            output_dir=resolved_output_dir,
            workbook_output_dir=workbook_output_dir,
            chart_output_dir=chart_output_dir,
            workbook_copy=workbook_output_dir / resolved_workbook_path.name,
        )

    def chart_path_for_run(self, run_index: int, sample_count: int) -> Path:
        run_label = f"run_{run_index:02d}_b1_{sample_count}"
        return self.chart_output_dir / f"{self.workbook_path.stem}_{run_label}.png"


@dataclass(slots=True)
class RunResult:
    sample_count: int
    workbook_copy: Path
    chart_path: Path
    sample_table: pd.DataFrame


class ExcelChartExporter:
    def __init__(self, contract: WorkbookContract) -> None:
        self.contract = contract

    def export(self, book: xw.Book, chart_path: Path) -> None:
        chart_path.parent.mkdir(parents=True, exist_ok=True)
        chart_path.unlink(missing_ok=True)
        sheet = book.sheets[self.contract.sheet_name]
        original_visibility = book.app.visible

        try:
            chart = sheet.charts[self.contract.chart_name]
        except Exception as exc:
            raise RuntimeError(
                f"Failed to access chart '{self.contract.chart_name}' on sheet "
                f"'{self.contract.sheet_name}': {exc}"
            ) from exc

        export_errors: list[str] = []
        self._prepare_chart_export(book, sheet, export_errors)

        try:
            if self._try_native_export(chart, chart_path, export_errors):
                return
            if self._try_clipboard_export(chart, chart_path, export_errors):
                return
            if self._try_appscript_export(chart, chart_path, export_errors):
                return
        finally:
            try:
                book.app.visible = original_visibility
            except Exception:
                pass

        raise RuntimeError(
            f"Chart export did not produce a PNG at {chart_path}. "
            f"Tried: {' | '.join(export_errors)}"
        )

    def _prepare_chart_export(
        self,
        book: xw.Book,
        sheet: xw.Sheet,
        export_errors: list[str],
    ) -> None:
        try:
            book.app.visible = True
            book.app.activate(steal_focus=True)
            sheet.activate()
            time.sleep(0.5)
        except Exception as exc:
            export_errors.append(f"app/sheet activation: {exc}")

    def _try_native_export(
        self,
        chart: xw.Chart,
        chart_path: Path,
        export_errors: list[str],
    ) -> bool:
        try:
            chart.to_png(str(chart_path))
            if chart_path.exists():
                return True
            export_errors.append("chart.to_png did not create the file")
        except XlwingsError as exc:
            export_errors.append(f"chart.to_png: {exc}")
        except Exception as exc:
            export_errors.append(f"chart.to_png: {exc}")
        return False

    def _try_clipboard_export(
        self,
        chart: xw.Chart,
        chart_path: Path,
        export_errors: list[str],
    ) -> bool:
        api = chart.api
        for index, label in ((1, "chart.api[1].copy_picture"), (0, "chart.api[0].copy_picture")):
            try:
                if not (isinstance(api, tuple) and len(api) > index and mac_kw):
                    continue
                api[index].copy_picture(appearance=mac_kw.screen, format=mac_kw.bitmap)
                time.sleep(0.75)
                clipboard_image = ImageGrab.grabclipboard()
                if clipboard_image is None:
                    export_errors.append(f"{label} did not place an image on the clipboard")
                    continue
                clipboard_image.save(chart_path)
                if chart_path.exists():
                    return True
                export_errors.append(f"{label} created a clipboard image but no file")
            except Exception as exc:
                export_errors.append(f"{label}: {exc}")
        return False

    def _try_appscript_export(
        self,
        chart: xw.Chart,
        chart_path: Path,
        export_errors: list[str],
    ) -> bool:
        api = chart.api

        try:
            if isinstance(api, tuple) and len(api) >= 1 and posix_to_hfs_path and mac_kw:
                api[0].save_as_picture(
                    file_name=posix_to_hfs_path(str(chart_path.resolve())),
                    picture_type=mac_kw.save_as_PNG_file,
                )
                if chart_path.exists():
                    return True
                export_errors.append("chart.api[0].save_as_picture did not create the file")
        except Exception as exc:
            export_errors.append(f"chart.api[0].save_as_picture: {exc}")

        try:
            if isinstance(api, tuple) and len(api) >= 2 and posix_to_hfs_path:
                api[1].save_as(filename=posix_to_hfs_path(str(chart_path.resolve())))
                if chart_path.exists():
                    return True
                export_errors.append("chart.api[1].save_as did not create the file")
        except Exception as exc:
            export_errors.append(f"chart.api[1].save_as: {exc}")
        return False


class ExcelWorkbookRunner:
    def __init__(
        self,
        paths: WorkflowPaths,
        contract: WorkbookContract | None = None,
        chart_exporter: ExcelChartExporter | None = None,
    ) -> None:
        self.paths = paths
        self.contract = contract or WorkbookContract()
        self.chart_exporter = chart_exporter or ExcelChartExporter(self.contract)

    @classmethod
    def from_paths(
        cls,
        workbook_path: str | Path = "Test.xlsm",
        output_dir: str | Path = "notebook_outputs",
        contract: WorkbookContract | None = None,
    ) -> ExcelWorkbookRunner:
        return cls(
            paths=WorkflowPaths.create(workbook_path=workbook_path, output_dir=output_dir),
            contract=contract,
        )

    def run_for_inputs(
        self,
        sample_counts: Iterable[int],
        *,
        visible: bool = False,
    ) -> list[RunResult]:
        if not self.paths.workbook_path.exists():
            raise FileNotFoundError(f"Workbook not found: {self.paths.workbook_path}")

        normalized_counts = [_normalize_sample_count(value) for value in sample_counts]
        if not normalized_counts:
            raise ValueError("Provide at least one value to write into cell B1.")

        results: list[RunResult] = []

        try:
            with xw.App(visible=visible, add_book=False) as app:
                app.display_alerts = False
                app.screen_updating = False

                for index, sample_count in enumerate(normalized_counts, start=1):
                    results.append(
                        self._run_single_sample(
                            app=app,
                            sample_count=sample_count,
                            run_index=index,
                        )
                    )
        except Exception as exc:
            self._log_excel_exception(exc)
            raise RuntimeError(
                "Excel automation failed. Make sure Microsoft Excel is installed, "
                "macOS has granted automation access to Python/Terminal, and macros "
                "are enabled when the copied workbook opens."
            ) from exc

        return results

    def _run_single_sample(
        self,
        *,
        app: xw.App,
        sample_count: int,
        run_index: int,
    ) -> RunResult:
        chart_path = self.paths.chart_path_for_run(run_index, sample_count)
        shutil.copy2(self.paths.workbook_path, self.paths.workbook_copy)

        book = app.books.open(str(self.paths.workbook_copy))
        try:
            sheet = book.sheets[self.contract.sheet_name]
            sheet.range(self.contract.input_cell).value = sample_count

            self._call_macro(book, self.contract.macro_name)
            book.save()
            self.chart_exporter.export(book, chart_path)
            sample_table = self._read_sample_table(sheet, sample_count)

            return RunResult(
                sample_count=sample_count,
                workbook_copy=self.paths.workbook_copy,
                chart_path=chart_path,
                sample_table=sample_table,
            )
        finally:
            book.close()

    def _call_macro(self, book: xw.Book, macro_name: str, *args: object) -> object:
        attempts: list[str] = []
        workbook_names = [book.name]
        if "." in book.name:
            workbook_names.append(book.name.rsplit(".", 1)[0])

        scoped_names = [macro_name, f"Module1.{macro_name}"]

        for scoped_name in scoped_names:
            try:
                return book.macro(scoped_name)(*args)
            except Exception as exc:
                attempts.append(f"{scoped_name}: {exc}")

        for workbook_name in workbook_names:
            for scoped_name in scoped_names:
                for candidate in (
                    f"{workbook_name}!{scoped_name}",
                    f"'{workbook_name}'!{scoped_name}",
                ):
                    try:
                        return book.app.macro(candidate)(*args)
                    except Exception as exc:
                        attempts.append(f"{candidate}: {exc}")

        raise RuntimeError(
            f"Unable to call VBA macro '{macro_name}'. Tried: " + " | ".join(attempts)
        )

    def _read_sample_table(self, sheet: xw.Sheet, sample_count: int) -> pd.DataFrame:
        headers = sheet.range(self.contract.sample_table_header_range).options(ndim=1).value
        rows = sheet.range(self.contract.sample_table_range(sample_count)).options(ndim=2).value

        if rows is None:
            rows = []

        return pd.DataFrame(rows, columns=headers)

    @staticmethod
    def _log_excel_exception(exc: Exception) -> None:
        print("Excel automation error:", file=sys.stderr)
        traceback.print_exception(type(exc), exc, exc.__traceback__, file=sys.stderr)


def parse_sample_counts(raw_values: str) -> list[int]:
    tokens = [token for token in re.split(r"[\s,]+", raw_values.strip()) if token]
    if not tokens:
        raise ValueError("Enter at least one positive integer for cell B1.")

    sample_counts: list[int] = []
    for token in tokens:
        try:
            value = int(token)
        except ValueError as exc:
            raise ValueError(f"'{token}' is not a valid integer for cell B1.") from exc

        sample_counts.append(_normalize_sample_count(value))

    return sample_counts


def run_workbook_for_inputs(
    sample_counts: Iterable[int],
    workbook_path: str | Path = "Test.xlsm",
    output_dir: str | Path = "notebook_outputs",
    *,
    visible: bool = False,
) -> list[RunResult]:
    runner = ExcelWorkbookRunner.from_paths(
        workbook_path=workbook_path,
        output_dir=output_dir,
    )
    return runner.run_for_inputs(sample_counts, visible=visible)


def _normalize_sample_count(value: int) -> int:
    sample_count = int(value)
    if sample_count < 1:
        raise ValueError("Each value written to cell B1 must be at least 1.")
    return sample_count
