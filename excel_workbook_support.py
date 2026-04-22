from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
import sys
import time
import traceback

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
class ExcelChartSpec:
    sheet_name: str
    chart_name: str


class ExcelChartExporter:
    def __init__(self, chart_spec: ExcelChartSpec) -> None:
        self.chart_spec = chart_spec

    def export(self, book: xw.Book, chart_path: Path) -> None:
        chart_path.parent.mkdir(parents=True, exist_ok=True)
        chart_path.unlink(missing_ok=True)
        sheet = book.sheets[self.chart_spec.sheet_name]
        original_visibility = book.app.visible

        try:
            chart = sheet.charts[self.chart_spec.chart_name]
        except Exception as exc:
            raise RuntimeError(
                f"Failed to access chart '{self.chart_spec.chart_name}' on sheet "
                f"'{self.chart_spec.sheet_name}': {exc}"
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


def call_vba_macro(book: xw.Book, macro_name: str, *args: object) -> object:
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


def log_excel_exception(exc: Exception) -> None:
    print("Excel automation error:", file=sys.stderr)
    traceback.print_exception(type(exc), exc, exc.__traceback__, file=sys.stderr)
