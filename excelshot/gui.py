"""Tkinter desktop interface for Excel 连续截图.

The GUI deliberately contains no COM or OCR code.  The capture backend is imported
inside the worker thread so importing this module remains safe on machines without
Office (and so COM is initialised and owned by the capture thread).
"""

from __future__ import annotations

import os
import queue
import subprocess
import sys
import threading
import time
import traceback
from pathlib import Path
from typing import Any, Callable, Optional

import tkinter as tk
from tkinter import filedialog, messagebox, ttk


APP_TITLE = "Excel 连续截图"
_SUPPORTED = {
    ".xlsx": "Excel 工作簿",
    ".xls": "Excel 97-2003 工作簿",
    ".xlsm": "启用宏的 Excel 工作簿",
    ".xlsb": "Excel 二进制工作簿",
}


def app_dir() -> Path:
    """Return the directory beside the executable (or the project when developing)."""

    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    # gui.py lives in ``excelshot/``; use the repository root in source mode.
    return Path(__file__).resolve().parent.parent


class ExcelShotApp:
    """Small, keyboard-friendly front end around :mod:`excelshot.backend`."""

    def __init__(self, master: Optional[tk.Tk] = None) -> None:
        self.master = master or tk.Tk()
        self.master.title(APP_TITLE)
        self.master.minsize(680, 460)
        self.master.geometry("760x560")
        self.master.protocol("WM_DELETE_WINDOW", self._on_close)
        self.master.bind_all("<Escape>", self._on_escape)

        self._events: queue.Queue[tuple[str, Any]] = queue.Queue()
        self._cancelled = threading.Event()
        self._worker: Optional[threading.Thread] = None
        self._closing = False
        self._iconified_for_capture = False
        self._last_result: Optional[Path] = None

        self.path_var = tk.StringVar()
        self.sheet_var = tk.StringVar()
        self.address_var = tk.StringVar()
        self.software_var = tk.StringVar(value="auto")
        self.zoom_var = tk.StringVar(value="100")
        self.delay_var = tk.StringVar(value="0.6")
        self.status_var = tk.StringVar(value="请选择一个 Excel 文件开始")

        self._build_style()
        self._build_widgets()
        self._poll_events()

    # ------------------------------------------------------------------ UI
    def _build_style(self) -> None:
        style = ttk.Style(self.master)
        try:
            style.theme_use("vista")
        except tk.TclError:
            # ``vista`` is present on normal Windows 10 installations.  The
            # fallback keeps the program usable under Wine and in test runners.
            try:
                style.theme_use("clam")
            except tk.TclError:
                pass
        style.configure("Title.TLabel", font=("Microsoft YaHei UI", 17, "bold"))
        style.configure("Hint.TLabel", foreground="#5b6573")
        style.configure("Status.TLabel", foreground="#1e5b96")
        style.configure("Primary.TButton", padding=(16, 6))

    def _build_widgets(self) -> None:
        outer = ttk.Frame(self.master, padding=(22, 18, 22, 14))
        outer.pack(fill="both", expand=True)
        outer.columnconfigure(0, weight=1)
        outer.rowconfigure(4, weight=1)

        ttk.Label(outer, text=APP_TITLE, style="Title.TLabel").grid(
            row=0, column=0, sticky="w"
        )
        ttk.Label(
            outer,
            text="按完整的 Excel 单元格行逐页滚动截图，用离线 OCR 检查相邻页衔接。",
            style="Hint.TLabel",
        ).grid(row=1, column=0, sticky="w", pady=(4, 15))

        file_frame = ttk.LabelFrame(outer, text="Excel 文件", padding=12)
        file_frame.grid(row=2, column=0, sticky="ew", pady=(0, 10))
        file_frame.columnconfigure(0, weight=1)
        self.file_entry = ttk.Entry(file_frame, textvariable=self.path_var)
        self.file_entry.grid(row=0, column=0, sticky="ew", padx=(0, 8))
        self.browse_button = ttk.Button(file_frame, text="选择文件…", command=self._browse)
        self.browse_button.grid(row=0, column=1)

        options = ttk.LabelFrame(outer, text="截图设置（留空使用自动值）", padding=12)
        options.grid(row=3, column=0, sticky="ew", pady=(0, 10))
        for col, weight in ((1, 1), (3, 1)):
            options.columnconfigure(col, weight=weight)

        ttk.Label(options, text="工作表名称").grid(row=0, column=0, sticky="w", padx=(0, 7), pady=4)
        ttk.Entry(options, textvariable=self.sheet_var).grid(row=0, column=1, sticky="ew", padx=(0, 20), pady=4)
        ttk.Label(options, text="A1 范围").grid(row=0, column=2, sticky="w", padx=(0, 7), pady=4)
        ttk.Entry(options, textvariable=self.address_var).grid(row=0, column=3, sticky="ew", pady=4)

        ttk.Label(options, text="打开方式").grid(row=1, column=0, sticky="w", padx=(0, 7), pady=4)
        self.software_combo = ttk.Combobox(
            options,
            textvariable=self.software_var,
            state="readonly",
            values=("auto", "excel", "wps"),
            width=10,
        )
        self.software_combo.grid(row=1, column=1, sticky="w", padx=(0, 20), pady=4)
        ttk.Label(options, text="缩放比例 (%)").grid(row=1, column=2, sticky="w", padx=(0, 7), pady=4)
        ttk.Entry(options, textvariable=self.zoom_var, width=10).grid(row=1, column=3, sticky="w", pady=4)

        ttk.Label(options, text="翻页后等待 (秒)").grid(row=2, column=0, sticky="w", padx=(0, 7), pady=4)
        ttk.Entry(options, textvariable=self.delay_var, width=10).grid(row=2, column=1, sticky="w", padx=(0, 20), pady=4)
        ttk.Label(
            options,
            text="仅截取单元格区域；相邻页面自动上下衔接，不重复。",
            style="Hint.TLabel",
        ).grid(row=2, column=2, columnspan=2, sticky="w", pady=4)

        log_frame = ttk.LabelFrame(outer, text="运行日志", padding=(8, 8, 8, 7))
        log_frame.grid(row=4, column=0, sticky="nsew")
        log_frame.columnconfigure(0, weight=1)
        log_frame.rowconfigure(0, weight=1)
        self.log_text = tk.Text(
            log_frame,
            height=8,
            wrap="word",
            state="disabled",
            bg="#fbfcfe",
            relief="flat",
            padx=8,
            pady=6,
        )
        self.log_text.grid(row=0, column=0, sticky="nsew")
        scroll = ttk.Scrollbar(log_frame, orient="vertical", command=self.log_text.yview)
        scroll.grid(row=0, column=1, sticky="ns")
        self.log_text.configure(yscrollcommand=scroll.set)

        bottom = ttk.Frame(outer)
        bottom.grid(row=5, column=0, sticky="ew", pady=(12, 0))
        bottom.columnconfigure(0, weight=1)
        ttk.Label(bottom, textvariable=self.status_var, style="Status.TLabel").grid(row=0, column=0, sticky="w")
        self.progress = ttk.Progressbar(bottom, mode="indeterminate", length=115)
        self.progress.grid(row=0, column=1, padx=10)
        self.start_button = ttk.Button(bottom, text="开始截图", style="Primary.TButton", command=self._start)
        self.start_button.grid(row=0, column=2, padx=(0, 7))
        self.stop_button = ttk.Button(bottom, text="停止 (Esc)", command=self._stop, state="disabled")
        self.stop_button.grid(row=0, column=3, padx=(0, 7))
        self.open_button = ttk.Button(bottom, text="打开 shot 文件夹", command=self._open_shot_folder)
        self.open_button.grid(row=0, column=4)

        ttk.Label(
            outer,
            text="提示：截图过程中窗口会最小化以避免遮挡 Excel。也可以按 Esc 请求停止。",
            style="Hint.TLabel",
        ).grid(row=6, column=0, sticky="w", pady=(9, 0))

    # ----------------------------------------------------------- interactions
    def _browse(self) -> None:
        value = filedialog.askopenfilename(
            title="选择 Excel 文件",
            filetypes=(
                ("Excel 文件", "*.xlsx *.xls *.xlsm *.xlsb"),
                ("所有文件", "*.*"),
            ),
        )
        if value:
            self.path_var.set(value)
            self._append_log(f"已选择：{value}")

    def _parse_options(self) -> tuple[Path, int, float] | None:
        raw_path = self.path_var.get().strip().strip('"')
        if not raw_path:
            messagebox.showwarning(APP_TITLE, "请先选择 Excel 文件。", parent=self.master)
            return None
        path = Path(raw_path)
        if not path.is_file():
            messagebox.showerror(APP_TITLE, f"找不到文件：\n{path}", parent=self.master)
            return None
        if path.suffix.lower() not in _SUPPORTED:
            messagebox.showerror(APP_TITLE, "请选择 .xlsx、.xls、.xlsm 或 .xlsb 文件。", parent=self.master)
            return None
        try:
            zoom = int(float(self.zoom_var.get().strip() or "100"))
        except ValueError:
            messagebox.showerror(APP_TITLE, "缩放比例需要是数字（例如 100）。", parent=self.master)
            return None
        if zoom < 10 or zoom > 400:
            messagebox.showerror(APP_TITLE, "缩放比例应在 10 到 400 之间。", parent=self.master)
            return None
        try:
            delay = float(self.delay_var.get().strip() or "0.6")
        except ValueError:
            messagebox.showerror(APP_TITLE, "等待时间需要是数字（例如 0.6）。", parent=self.master)
            return None
        if delay < 0 or delay > 30:
            messagebox.showerror(APP_TITLE, "等待时间应在 0 到 30 秒之间。", parent=self.master)
            return None
        return path, zoom, delay

    def _start(self) -> None:
        parsed = self._parse_options()
        if parsed is None or self._worker is not None and self._worker.is_alive():
            return
        path, zoom, delay = parsed
        self._cancelled.clear()
        self._last_result = None
        self._set_running(True)
        self._append_log("准备开始，后台 OCR 衔接检查已启用…")
        # Excel/WPS owns the screen while it is captured.  Iconify instead of
        # withdrawing so the user can restore this window and press Stop.
        try:
            self.master.iconify()
            self._iconified_for_capture = True
        except tk.TclError:
            self._iconified_for_capture = False
        self._worker = threading.Thread(
            target=self._capture_worker,
            args=(path, self.sheet_var.get().strip(), self.address_var.get().strip(), self.software_var.get(), zoom, delay),
            daemon=True,
            name="excelshot-capture",
        )
        self._worker.start()
        # The window is iconified while Excel owns the screen, so Tk may not
        # receive Escape.  A tiny Windows-only watcher keeps Escape available
        # as a global stop key; on non-Windows it simply does nothing.
        threading.Thread(target=self._watch_global_escape, daemon=True, name="excelshot-escape").start()

    def _watch_global_escape(self) -> None:
        try:
            import ctypes

            user32 = ctypes.windll.user32
            previous = False
            while self._worker is not None and self._worker.is_alive() and not self._cancelled.is_set():
                pressed = bool(user32.GetAsyncKeyState(0x1B) & 0x8000)  # VK_ESCAPE
                if pressed and not previous:
                    self._cancelled.set()
                    self._events.put(("log", "检测到 Esc，正在停止…"))
                    self._events.put(("stop_requested", None))
                    return
                previous = pressed
                time.sleep(0.08)
        except (AttributeError, OSError):
            # No user32 (for example in a Linux test runner); the Tk binding
            # remains available when the window has focus.
            return

    def _capture_worker(self, path: Path, sheet: str, address: str, software: str, zoom: int, delay: float) -> None:
        """Import and invoke the backend in this thread (COM must stay here)."""

        try:
            try:
                from . import backend  # type: ignore
            except ImportError:
                import excelshot.backend as backend  # type: ignore
            options = backend.CaptureOptions(
                path=str(path),
                sheet=sheet,
                address=address,
                software=software,
                zoom=zoom,
                settle_seconds=delay,
            )
            result = backend.capture(options, log=self._worker_log, cancelled=self._cancelled)
            self._events.put(("done", result))
        except Exception as exc:  # report on the main thread without crashing Tk
            details = "".join(traceback.format_exception_only(type(exc), exc)).strip()
            self._events.put(("error", details or repr(exc)))
            self._events.put(("debug", traceback.format_exc()))

    def _worker_log(self, message: str) -> None:
        self._events.put(("log", str(message)))

    def _stop(self) -> None:
        if self._worker is not None and self._worker.is_alive():
            self._cancelled.set()
            self._append_log("正在请求停止；请等待 Excel 完成当前操作…")
            self.stop_button.configure(state="disabled")
            self.status_var.set("正在停止…")

    def _on_escape(self, _event: Optional[tk.Event] = None) -> str:
        self._stop()
        return "break"

    def _on_close(self) -> None:
        if self._worker is not None and self._worker.is_alive():
            self._cancelled.set()
            self._closing = True
            self.status_var.set("正在停止，窗口稍后关闭…")
            self.master.iconify()
            return
        self.master.destroy()

    def _set_running(self, running: bool) -> None:
        state = "disabled" if running else "normal"
        self.start_button.configure(state=state)
        self.browse_button.configure(state=state)
        self.file_entry.configure(state=state)
        self.stop_button.configure(state="normal" if running else "disabled")
        if running:
            self.progress.start(12)
            self.status_var.set("截图进行中…")
        else:
            self.progress.stop()

    def _restore_window(self) -> None:
        if self._iconified_for_capture:
            try:
                self.master.deiconify()
                self.master.lift()
            except tk.TclError:
                pass
            self._iconified_for_capture = False

    def _poll_events(self) -> None:
        try:
            while True:
                kind, payload = self._events.get_nowait()
                if kind == "log":
                    self._append_log(str(payload))
                elif kind == "done":
                    self._last_result = Path(payload) if payload else None
                    self._restore_window()
                    self._set_running(False)
                    if self._cancelled.is_set():
                        self.status_var.set("已停止")
                        self._append_log("任务已停止。")
                    else:
                        shown = str(self._last_result) if self._last_result else str(app_dir() / "shot")
                        self.status_var.set("截图完成")
                        self._append_log(f"截图完成，输出目录：{shown}")
                    self._worker = None
                    if self._closing:
                        self.master.destroy()
                elif kind == "error":
                    self._restore_window()
                    self._set_running(False)
                    self.status_var.set("运行失败")
                    self._append_log(f"错误：{payload}")
                    self._worker = None
                    if not self._closing:
                        messagebox.showerror(APP_TITLE, str(payload), parent=self.master)
                    else:
                        self.master.destroy()
                elif kind == "debug":
                    # Keep details in the log for troubleshooting, but do not
                    # pop up a second noisy dialog.
                    self._append_log(str(payload).rstrip())
                elif kind == "stop_requested":
                    self.stop_button.configure(state="disabled")
                    self.status_var.set("正在停止…")
        except queue.Empty:
            pass
        if self.master.winfo_exists():
            self.master.after(100, self._poll_events)

    def _append_log(self, message: str) -> None:
        stamp = time.strftime("%H:%M:%S")
        self.log_text.configure(state="normal")
        self.log_text.insert("end", f"[{stamp}] {message}\n")
        self.log_text.see("end")
        self.log_text.configure(state="disabled")

    def _open_shot_folder(self) -> None:
        folder = app_dir() / "shot"
        folder.mkdir(parents=True, exist_ok=True)
        try:
            os.startfile(str(folder))  # type: ignore[attr-defined]
        except AttributeError:
            subprocess.Popen(["explorer", str(folder)])
        except OSError as exc:
            messagebox.showerror(APP_TITLE, f"无法打开文件夹：\n{exc}", parent=self.master)


def main(_argv: Optional[list[str]] = None) -> None:
    """Launch the GUI.  Kept argument-free for the frozen executable entrypoint."""

    root = tk.Tk()
    ExcelShotApp(root)
    root.mainloop()


# Useful aliases for a small runner and for embedders/tests.
App = ExcelShotApp
run_gui = main
run = main
launch = main


if __name__ == "__main__":  # pragma: no cover - convenience for source users
    main()
