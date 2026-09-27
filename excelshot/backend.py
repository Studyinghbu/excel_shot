"""Windows Excel/WPS cell-only capture engine.

The engine deliberately copies a *rectangular Range* to a temporary chart and
exports that chart.  This avoids row/column headers, window chrome, and the
unreliable partial-row behaviour of mouse wheel screenshots.  Pages are made
of complete worksheet rows and are scrolled to before each export.
"""
from __future__ import annotations

import hashlib
import json
import os
import re
import sys
import time
import traceback
from dataclasses import dataclass, asdict
from datetime import datetime
from pathlib import Path
from threading import Event
from typing import Callable, Optional

from .pagination import RowSpan, plan_pages


LOG = Callable[[str], None]


@dataclass
class CaptureOptions:
    path: str
    sheet: str = ""
    address: str = ""
    software: str = "auto"       # auto | excel | wps
    zoom: int = 100
    settle_seconds: float = 0.6


@dataclass
class PageRecord:
    index: int
    address: str
    start_row: int
    end_row: int
    file: str
    sha256: str
    ocr_text: str = ""
    ocr_warning: str = ""


_A1 = re.compile(r"^\$?([A-Za-z]{1,3})\$?([1-9][0-9]*)$")
_A1RANGE = re.compile(
    r"^\$?([A-Za-z]{1,3})\$?([1-9][0-9]*)"
    r"(?::\$?([A-Za-z]{1,3})\$?([1-9][0-9]*))?$"
)


def _root_dir() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parents[1]


def shot_directory() -> Path:
    out = _root_dir() / "shot"
    out.mkdir(parents=True, exist_ok=True)
    return out


def _log(log: LOG, message: str) -> None:
    try:
        log(message)
    except Exception:
        pass


def _safe_name(value: str, fallback: str = "sheet") -> str:
    value = re.sub(r"[<>:\"/\\|?*\x00-\x1f]", "_", str(value)).strip(" .")
    return (value or fallback)[:80]


def _column_number(value: str) -> int:
    n = 0
    for char in value.upper():
        n = n * 26 + ord(char) - 64
    return n


def _parse_address(value: str) -> tuple[str, str, int, int, int, int]:
    match = _A1RANGE.fullmatch(value.strip())
    if not match:
        raise ValueError("区域必须是 A1 或 A1:D200 这样的单元格范围")
    c1, r1, c2, r2 = match.groups()
    col1, row1 = _column_number(c1), int(r1)
    c2 = c2 or c1
    r2 = r2 or r1
    col2, row2 = _column_number(c2), int(r2)
    if col2 < col1 or row2 < row1:
        raise ValueError("单元格范围的起点必须在终点左上方")
    return c1.upper(), c2.upper(), row1, row2, col1, col2


def _is_hidden(value) -> bool:
    try:
        return bool(value)
    except Exception:
        return False


def _range_address(ws, r1: int, c1: int, r2: int, c2: int):
    return ws.Range(ws.Cells(r1, c1), ws.Cells(r2, c2))


def _page_rows(ws, start: int, end: int, c1: int, c2: int):
    heights: dict[int, float] = {}
    hidden: set[int] = set()
    for row in range(start, end + 1):
        obj = ws.Rows(row)
        if _is_hidden(getattr(obj, "Hidden", False)):
            hidden.add(row)
            heights[row] = 0.0
            continue
        try:
            h = float(obj.RowHeight or 0)
        except Exception:
            h = 15.0
        heights[row] = max(0.0, h)
    return heights, hidden


def _prog_id(software: str) -> str:
    if software.lower() == "wps":
        return "ket.Application"
    return "Excel.Application"


def _get_dispatch(prog_id: str, existing_path: str = "", *, force_new: bool = False):
    import win32com.client as win32
    # If the requested workbook is already open, attaching avoids a second
    # instance and preserves the user's existing view.
    if existing_path and not force_new:
        try:
            attached = win32.GetObject(os.path.abspath(existing_path))
            # GetObject(file) commonly returns a Workbook rather than the
            # Application.  Promote it to its owner so callers can use the
            # same code for a newly created and an already-open workbook.
            try:
                owner = attached.Application
                if hasattr(owner, "Workbooks"):
                    return owner, False
            except Exception:
                pass
            if hasattr(attached, "Workbooks"):
                return attached, False
        except Exception:
            pass
    try:
        return win32.DispatchEx(prog_id), True
    except Exception as exc:
        if prog_id != "Excel.Application":
            raise RuntimeError("没有找到 WPS 表格 COM 接口，请在软件选项选择 Excel") from exc
        raise RuntimeError("没有找到 Microsoft Excel COM 接口") from exc


def _find_workbook(app, path: str):
    norm = os.path.normcase(os.path.abspath(path))
    for wb in app.Workbooks:
        try:
            if os.path.normcase(os.path.abspath(str(wb.FullName))) == norm:
                return wb, False
        except Exception:
            continue
    wb = app.Workbooks.Open(os.path.abspath(path), ReadOnly=True, UpdateLinks=0, AddToMru=False)
    return wb, True


def _worksheet(app, wb, name: str):
    if name:
        try:
            return wb.Worksheets(name)
        except Exception as exc:
            raise ValueError(f"找不到工作表：{name}") from exc
    try:
        active_wb = app.ActiveWorkbook
        if active_wb is not None and str(active_wb.Name) == str(wb.Name):
            return app.ActiveSheet
    except Exception:
        pass
    return wb.Worksheets(1)


def _used_bounds(ws) -> tuple[int, int, int, int]:
    used = ws.UsedRange
    start_row = int(used.Row)
    start_col = int(used.Column)
    end_row = start_row + int(used.Rows.Count) - 1
    end_col = start_col + int(used.Columns.Count) - 1
    return start_row, start_col, end_row, end_col


def _data_bounds(ws) -> tuple[int, int, int, int]:
    """Return the first/last non-empty cells, avoiding formatting-only UsedRange."""
    fallback = _used_bounds(ws)
    try:
        # Excel Find constants: xlValues=-4163, xlByRows=1, xlByColumns=2,
        # xlNext=1, xlPrevious=2.  Named arguments work in both Excel and
        # the compatible WPS automation server.
        # Search from opposite corners so Find's wrap-around behaviour is
        # deterministic (some localized Excel builds ignore omitted After).
        first_after = ws.Cells(ws.Rows.Count, ws.Columns.Count)
        last_after = ws.Cells(1, 1)
        find_kwargs = {"What": "*", "LookAt": 2, "LookIn": -4123, "MatchCase": False, "SearchFormat": False}
        first_row_cell = ws.Cells.Find(After=first_after, SearchOrder=1, SearchDirection=1, **find_kwargs)
        last_row_cell = ws.Cells.Find(After=last_after, SearchOrder=1, SearchDirection=2, **find_kwargs)
        first_col_cell = ws.Cells.Find(After=first_after, SearchOrder=2, SearchDirection=1, **find_kwargs)
        last_col_cell = ws.Cells.Find(After=last_after, SearchOrder=2, SearchDirection=2, **find_kwargs)
        if all((first_row_cell, last_row_cell, first_col_cell, last_col_cell)):
            return (
                int(first_row_cell.Row), int(first_col_cell.Column),
                int(last_row_cell.Row), int(last_col_cell.Column),
            )
    except Exception:
        pass
    return fallback


def _wait_or_cancel(seconds: float, cancelled: Event) -> bool:
    return cancelled.wait(max(0.0, seconds))


def _prepare_excel_window(app, wb=None) -> None:
    """Make the workbook window an interactive, maximized desktop window.

    Excel launched through COM can inherit a 1-pixel/thumbnail window state.
    CopyPicture then returns a generic RPC error because there is no drawable
    sheet surface.  This helper is intentionally called after opening the
    workbook, when ActiveWindow actually exists.
    """
    import ctypes
    from ctypes import wintypes

    # A COM-launched Excel process can take a few seconds to create its first
    # top-level window.  Accessing ActiveWindow immediately during that period
    # either returns None or (on some Office builds) raises RPC_S_CALL_FAILED.
    # Polling here also gives Excel time to finish loading the workbook before
    # CopyPicture touches the drawing surface.
    try:
        app.Visible = True
    except Exception:
        pass
    try:
        app.UserControl = True
    except Exception:
        pass
    window = None
    deadline = time.monotonic() + 12.0
    while time.monotonic() < deadline and window is None:
        try:
            if wb is not None:
                wb.Activate()
        except Exception:
            pass
        try:
            window = app.ActiveWindow
        except Exception:
            window = None
        if window is not None:
            break
        try:
            # In a fresh instance ActiveWindow can lag behind Windows(1).
            windows = app.Windows
            if windows.Count:
                window = windows(1)
        except Exception:
            pass
        time.sleep(0.15)
    if window is None:
        raise RuntimeError("Excel 没有可用的工作簿窗口（Excel 可能仍在启动）")

    # Restore first, then maximize.  Setting WindowState alone is unreliable
    # when Excel starts with a tiny saved geometry, so repeat it through the
    # native window handle as well.
    try:
        window.Activate()
    except Exception:
        pass
    for state in (-4143, -4137):  # xlNormal, xlMaximized
        try:
            window.WindowState = state
        except Exception:
            pass
    try:
        app.WindowState = -4137  # Application-level state on newer Excel
    except Exception:
        pass
    hwnd = 0
    try:
        hwnd = int(window.Hwnd)
    except Exception:
        try:
            hwnd = int(app.Hwnd)
        except Exception:
            hwnd = 0
    if hwnd:
        try:
            user32 = ctypes.windll.user32
            user32.IsWindow.argtypes = [ctypes.c_void_p]
            user32.IsWindow.restype = ctypes.c_bool
            if user32.IsWindow(hwnd):
                user32.ShowWindow(hwnd, 9)  # SW_RESTORE
                user32.ShowWindow(hwnd, 3)  # SW_MAXIMIZE
                user32.BringWindowToTop(hwnd)
                user32.SetForegroundWindow(hwnd)
                # SWP_FRAMECHANGED makes a stale, saved 1-pixel frame redraw.
                user32.SetWindowPos(hwnd, 0, 0, 0, 0, 0, 0x0003 | 0x0020 | 0x0040)
                user32.ShowWindow(hwnd, 3)
        except Exception:
            # Locked-down desktops can reject SetForegroundWindow.  Excel can
            # still be used through COM when it is already visible.
            pass
    # Do not return until the window has a real client area.  This check is
    # deliberately native because querying Width/Height through a busy COM
    # server can itself produce another RPC error.
    for _ in range(20):
        try:
            if hwnd:
                rect = wintypes.RECT()
                if ctypes.windll.user32.GetWindowRect(hwnd, ctypes.byref(rect)):
                    if (rect.right - rect.left) >= 300 and (rect.bottom - rect.top) >= 200:
                        break
            elif float(window.Width) >= 300 and float(window.Height) >= 200:
                break
        except Exception:
            pass
        time.sleep(0.1)
    # Wait for Excel's automation queue to become idle before the first
    # clipboard operation.  Without this, CopyPicture can race workbook
    # painting and return the generic RPC_S_CALL_FAILED HRESULT.
    for _ in range(30):
        try:
            if bool(app.Ready):
                break
        except Exception:
            # Older WPS/Excel versions do not expose Ready; the fixed delay
            # below still gives their window time to paint.
            break
        time.sleep(0.1)
    time.sleep(0.35)


def _copy_range_png(ws, rng, filename: Path):
    # CopyPicture is kept inside the already prepared Excel window.  Repeatedly
    # activating/maximizing/selecting here made every page visibly flash.
    copy_error = None
    app = ws.Application
    try:
        last_error = None
        for attempt in range(2):
            try:
                if attempt:
                    _prepare_excel_window(app, ws.Parent)
                    ws.Activate()
                    time.sleep(0.4)
                rng.CopyPicture(Appearance=1, Format=2)  # xlScreen, xlBitmap
                # Prefer saving the copied bitmap directly.  It avoids adding
                # and deleting a temporary chart on the worksheet, which is
                # the main source of the visible flash.
                try:
                    from PIL import ImageGrab
                    time.sleep(0.06)  # let Excel publish the bitmap to CF_DIB
                    clipboard_image = ImageGrab.grabclipboard()
                    if hasattr(clipboard_image, "save"):
                        clipboard_image.save(str(filename), "PNG")
                        last_error = None
                        break
                except Exception:
                    pass
                width = max(8.0, min(10000.0, float(rng.Width)))
                height = max(8.0, min(20000.0, float(rng.Height)))
                old_screen = None
                try:
                    old_screen = bool(app.ScreenUpdating)
                    app.ScreenUpdating = False
                    charts = ws.ChartObjects()
                    chart_obj = charts.Add(0, 0, width, height)
                    try:
                        chart = chart_obj.Chart
                        chart.Paste()
                        chart.Export(str(filename), "PNG")
                    finally:
                        try:
                            chart_obj.Delete()
                        except Exception:
                            pass
                finally:
                    if old_screen is not None:
                        try:
                            app.ScreenUpdating = old_screen
                        except Exception:
                            pass
                last_error = None
                break
            except Exception as exc:
                last_error = exc
        if last_error is not None:
            raise last_error
    except Exception as exc:
        # Clipboard/ChartObjects are disabled in some Office policies.  The
        # fallback captures the exact visible Range rectangle, still excluding
        # all row/column headers and window chrome.
        copy_error = exc
        try:
            _screen_capture_range(ws, rng, filename)
        except Exception as fallback_exc:
            raise RuntimeError(f"Excel 单元格截图失败：{exc}；屏幕区域备用截图也失败：{fallback_exc}") from exc
    if not filename.exists() or filename.stat().st_size < 64:
        raise RuntimeError(f"Excel 没有生成截图：{copy_error or '未知错误'}")


def _is_rpc_failure(exc: BaseException) -> bool:
    """Return True for the COM failures that indicate a dead Excel server.

    Excel reports both RPC_S_CALL_FAILED and RPC_S_SERVER_UNAVAILABLE with
    different pywin32 exception strings depending on the Office version and
    locale.  Looking at the HRESULTs as well as the text lets capture() make
    one clean reconnect attempt instead of sending a stale Range to the
    screen-capture fallback.
    """
    text = " ".join(str(part) for part in (exc, getattr(exc, "__cause__", None), getattr(exc, "__context__", None)))
    return any(token in text.lower() for token in (
        "-2147023170", "-2147023174", "0x800706be", "0x800706ba",
        "rpc_s_call_failed", "rpc_s_server_unavailable", "rpc 服务器",
        "rpc server", "远程过程调用失败", "服务器不可用",
    ))


def _screen_capture_range(ws, rng, filename: Path) -> None:
    """Capture a visible cell rectangle when Excel's clipboard is unavailable."""
    import ctypes
    from PIL import ImageGrab

    app = ws.Application
    try:
        hwnd = int(app.ActiveWindow.Hwnd)
        user32 = ctypes.windll.user32
        # The normal path leaves Excel in the foreground.  Only request
        # focus when it was lost; doing this unconditionally caused flashing
        # between every page.
        if int(user32.GetForegroundWindow()) != hwnd:
            user32.ShowWindow(hwnd, 9)  # SW_RESTORE
            user32.BringWindowToTop(hwnd)
            user32.SetForegroundWindow(hwnd)
            time.sleep(0.12)
        if not user32.IsWindowVisible(hwnd):
            raise RuntimeError("Excel 窗口不可见")
        # ImageGrab reads the interactive desktop.  Refuse to save another
        # application's pixels if Windows did not grant Excel the foreground.
        if int(user32.GetForegroundWindow()) != hwnd:
            raise RuntimeError("Excel 窗口没有获得前台焦点")
    except Exception:
        raise RuntimeError("Excel 窗口不可见或未获得前台焦点")
    window = app.ActiveWindow
    left = float(rng.Left)
    top = float(rng.Top)
    right = left + float(rng.Width)
    bottom = top + float(rng.Height)
    x1 = int(window.PointsToScreenPixelsX(left))
    y1 = int(window.PointsToScreenPixelsY(top))
    x2 = int(window.PointsToScreenPixelsX(right))
    y2 = int(window.PointsToScreenPixelsY(bottom))
    if x2 <= x1 or y2 <= y1:
        raise RuntimeError("Excel 单元格没有完全显示在屏幕上")
    image = ImageGrab.grab(bbox=(x1, y1, x2, y2), all_screens=True)
    image.save(str(filename), "PNG")


def _sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def capture(options: CaptureOptions, log: LOG = print, cancelled: Optional[Event] = None) -> Path:
    """Capture every populated row in complete-cell pages and return shot dir."""
    cancelled = cancelled or Event()
    source = Path(options.path).expanduser()
    if not source.exists():
        raise FileNotFoundError(f"找不到 Excel 文件：{source}")
    if source.suffix.lower() not in {".xls", ".xlsx", ".xlsm", ".xlsb"}:
        raise ValueError("请选择 .xls、.xlsx、.xlsm 或 .xlsb 文件")
    out = shot_directory()
    started_at = datetime.now().strftime("%Y%m%d_%H%M%S")
    run_dir = out  # all files stay directly beside the EXE in shot/
    prog = _prog_id(options.software)
    app = wb = ws = None
    own_app = own_wb = False
    old_state = {}
    records: list[PageRecord] = []
    ocr = None
    try:
        _log(log, f"正在连接 {prog}…")
        # Use a clean COM instance.  Attaching with GetObject can inherit a
        # tiny/modal Excel window and a stale RPC channel from another Excel.
        app, own_app = _get_dispatch(prog, "", force_new=True)
        try:
            old_state["visible"] = bool(app.Visible)
            old_state["display_alerts"] = bool(app.DisplayAlerts)
            old_state["screen_updating"] = bool(app.ScreenUpdating)
            old_state["enable_events"] = bool(app.EnableEvents)
        except Exception:
            pass
        app.Visible = True
        app.DisplayAlerts = False
        try:
            app.EnableEvents = False
        except Exception:
            pass
        wb, own_wb = _find_workbook(app, str(source))
        # Must happen after Open: a fresh COM Excel instance has no
        # ActiveWindow before its first workbook is loaded.
        _prepare_excel_window(app, wb)
        ws = _worksheet(app, wb, options.sheet.strip())
        ws.Activate()
        try:
            old_state.update({
                "zoom": app.ActiveWindow.Zoom,
                "scroll_row": app.ActiveWindow.ScrollRow,
                "scroll_col": app.ActiveWindow.ScrollColumn,
            })
            if options.zoom:
                app.ActiveWindow.Zoom = max(10, min(400, int(options.zoom)))
        except Exception:
            pass
        ur1, uc1, ur2, uc2 = _data_bounds(ws)
        if options.address.strip():
            _, _, r1, r2, c1, c2 = _parse_address(options.address)
            if r1 < ur1 or c1 < uc1 or r2 > ur2 or c2 > uc2:
                _log(log, "提示：指定区域超出 UsedRange，仍按你指定的单元格范围截图。")
        else:
            r1, c1, r2, c2 = ur1, uc1, ur2, uc2
        heights, _hidden = _page_rows(ws, r1, r2, c1, c2)
        merged_spans: list[RowSpan] = []
        try:
            target_for_merges = _range_address(ws, r1, c1, r2, c2)
            seen_merges: set[tuple[int, int]] = set()
            for area in target_for_merges.MergeAreas:
                area_start = int(area.Row)
                area_end = area_start + int(area.Rows.Count) - 1
                if area_end > area_start and (area_start, area_end) not in seen_merges:
                    seen_merges.add((area_start, area_end))
                    merged_spans.append(RowSpan(area_start, area_end))
        except Exception:
            # MergeAreas is not exposed by some older WPS builds.  In that
            # case the normal complete-row planner still provides safe pages.
            merged_spans = []
        # A normal Excel window can display around 700–900 points at 100%;
        # 900 gives a comfortable PNG while keeping full rows and no overlap.
        max_height = 900.0
        try:
            visible = app.ActiveWindow.VisibleRange
            max_height = max(120.0, min(1800.0, float(visible.Height)))
        except Exception:
            pass
        pages = list(plan_pages(heights, r1, r2, max_height, merged_spans=merged_spans))
        if not pages:
            raise ValueError("所选区域没有可见的单元格行")
        stem = _safe_name(source.stem, "workbook") + "_" + _safe_name(ws.Name, "sheet")
        _log(log, f"工作表：{ws.Name}；范围：{r1}:{r2} 行 × {c1}:{c2} 列；共 {len(pages)} 页")
        try:
            from .ocr import OfflineOCR
            ocr = OfflineOCR()
            _log(log, "离线 OCR 已加载（只用于检查页间衔接）")
        except Exception as exc:
            _log(log, f"OCR 加载失败，将继续截图：{exc}")
        previous_ocr = None
        previous_hash = None
        for index, page in enumerate(pages, 1):
            if cancelled.is_set():
                _log(log, "已停止，已生成的截图保留在 shot 文件夹。")
                break
            start, end = page.start, page.end
            try:
                app.ActiveWindow.ScrollRow = start
            except Exception:
                try:
                    ws.Cells(start, c1).Select()
                except Exception:
                    pass
            if _wait_or_cancel(float(options.settle_seconds), cancelled):
                break
            rng = _range_address(ws, start, c1, end, c2)
            try:
                address = str(rng.Address(False, False))
            except TypeError:
                # Some pywin32 Excel builds expose Address as a property when
                # optional arguments are omitted.
                address = str(rng.Address)
            filename = run_dir / f"{stem}_p{index:04d}_r{start:06d}-{end:06d}.png"
            _log(log, f"第 {index}/{len(pages)} 页：{address}")
            try:
                _copy_range_png(ws, rng, filename)
            except Exception as capture_exc:
                # A minimized COM-launched Excel can lose its drawing server
                # during the first CopyPicture call.  Reusing the same Range
                # after RPC_S_CALL_FAILED only repeats the failure, so make
                # one clean Excel instance and retry the current page.
                if not _is_rpc_failure(capture_exc):
                    raise
                _log(log, "Excel 图形服务响应失败，正在重新连接并重试当前页…")
                try:
                    if own_wb and wb is not None:
                        wb.Close(SaveChanges=False)
                except Exception:
                    pass
                try:
                    if own_app and app is not None:
                        app.Quit()
                except Exception:
                    pass
                try:
                    app, own_app = _get_dispatch(prog, "", force_new=True)
                    app.Visible = True
                    app.DisplayAlerts = False
                    wb, own_wb = _find_workbook(app, str(source))
                    _prepare_excel_window(app, wb)
                    ws = _worksheet(app, wb, options.sheet.strip())
                    ws.Activate()
                    if options.zoom:
                        app.ActiveWindow.Zoom = max(10, min(400, int(options.zoom)))
                    rng = _range_address(ws, start, c1, end, c2)
                    try:
                        address = str(rng.Address(False, False))
                    except TypeError:
                        address = str(rng.Address)
                    _copy_range_png(ws, rng, filename)
                except Exception as retry_exc:
                    raise RuntimeError(f"Excel 重新连接后截图仍失败：{retry_exc}") from retry_exc
            digest = _sha256(filename)
            ocr_text, ocr_warning = "", ""
            if ocr is not None:
                try:
                    result = ocr.recognize(filename)
                    ocr_text = result.get("text", "")
                    if previous_ocr:
                        check = ocr.compare_adjacent(previous_ocr, result)
                        ocr_warning = check.get("warning", "")
                        if ocr_warning:
                            _log(log, f"  OCR 衔接提示：{ocr_warning}")
                    previous_ocr = result
                except Exception as exc:
                    ocr_warning = f"OCR 检查失败：{exc}"
            # Identical pixels indicate an automation glitch. Remove that
            # accidental duplicate, while keeping legitimate repeated values.
            if previous_hash and digest == previous_hash:
                filename.unlink(missing_ok=True)
                _log(log, "  检测到完全相同的图片，已跳过重复页。")
                continue
            previous_hash = digest
            records.append(PageRecord(index, address, start, end, filename.name, digest, ocr_text, ocr_warning))
        manifest = run_dir / f"{stem}_{started_at}_manifest.json"
        manifest.write_text(json.dumps({
            "created_at": datetime.now().isoformat(timespec="seconds"),
            "source": str(source), "sheet": str(ws.Name),
            "range": {"start_row": r1, "end_row": r2, "start_col": c1, "end_col": c2},
            "pages": [asdict(record) for record in records],
            "ocr": "offline RapidOCR (optional runtime module)",
        }, ensure_ascii=False, indent=2), encoding="utf-8")
        _log(log, f"完成：{len(records)} 张截图已保存到 {run_dir}")
        return run_dir
    except Exception as exc:
        _log(log, traceback.format_exc())
        raise RuntimeError(str(exc)) from exc
    finally:
        try:
            if app is not None:
                if "display_alerts" in old_state:
                    app.DisplayAlerts = old_state["display_alerts"]
                if "screen_updating" in old_state:
                    app.ScreenUpdating = old_state["screen_updating"]
                if "enable_events" in old_state:
                    app.EnableEvents = old_state["enable_events"]
                if "visible" in old_state and not own_app:
                    app.Visible = old_state["visible"]
            if app is not None and old_state and getattr(app, "ActiveWindow", None):
                app.ActiveWindow.Zoom = old_state.get("zoom", app.ActiveWindow.Zoom)
                app.ActiveWindow.ScrollRow = old_state.get("scroll_row", app.ActiveWindow.ScrollRow)
                app.ActiveWindow.ScrollColumn = old_state.get("scroll_col", app.ActiveWindow.ScrollColumn)
        except Exception:
            pass
        try:
            if own_wb and wb is not None:
                wb.Close(SaveChanges=False)
        except Exception:
            pass
        try:
            if own_app and app is not None:
                app.Quit()
        except Exception:
            pass
