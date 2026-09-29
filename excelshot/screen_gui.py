"""Tk interface for BrowserScreenShot.exe.

The selection overlay is a one-time desktop snapshot. During capture the Tk
window is withdrawn, so neither the app nor a translucent overlay is included
in the cell-only screenshots.
"""
from __future__ import annotations

import os
import queue
import threading
import time
import traceback
from pathlib import Path
from typing import Any, Optional

import tkinter as tk
from tkinter import messagebox, ttk
from PIL import ImageTk

from .screen_capture import CaptureOptions, capture_screen
from .screen_io import enable_dpi_awareness, grab_region, virtual_screen_bbox

APP_TITLE = "Browser Excel 连续截图"


class BrowserScreenApp:
    def __init__(self, master: Optional[tk.Tk] = None) -> None:
        enable_dpi_awareness()
        self.master = master or tk.Tk()
        self.master.title(APP_TITLE)
        self.master.minsize(720, 510)
        self.master.geometry("790x610")
        self.master.protocol("WM_DELETE_WINDOW", self._close)
        self.master.bind_all("<Escape>", lambda _event: self._stop())
        self.events: queue.Queue[tuple[str, Any]] = queue.Queue()
        self.cancelled = threading.Event()
        self.worker: Optional[threading.Thread] = None
        self.closing = False
        self.region: Optional[tuple[int, int, int, int]] = None
        self.last_run: Optional[Path] = None

        self.name_var = tk.StringVar(value="browser_sheet")
        self.delay_var = tk.StringVar(value="1.0")
        self.max_steps_var = tk.StringVar(value="200")
        self.wheel_var = tk.StringVar(value="1")
        self.countdown_var = tk.StringVar(value="5")
        self.region_var = tk.StringVar(value="尚未选择；只框选会滚动的单元格区域")
        self.status_var = tk.StringVar(value="请先框选单元格区域")
        self._build()
        self.master.after(100, self._poll)

    def _build(self) -> None:
        style = ttk.Style(self.master)
        try:
            style.theme_use("vista")
        except tk.TclError:
            pass
        style.configure("Title.TLabel", font=("Microsoft YaHei UI", 17, "bold"))
        style.configure("Hint.TLabel", foreground="#596777")
        style.configure("Status.TLabel", foreground="#155a91")
        outer = ttk.Frame(self.master, padding=(22, 18, 22, 14))
        outer.pack(fill="both", expand=True)
        outer.columnconfigure(0, weight=1)
        outer.rowconfigure(4, weight=1)
        ttk.Label(outer, text=APP_TITLE, style="Title.TLabel").grid(row=0, column=0, sticky="w")
        ttk.Label(outer, text="适用于浏览器/KVM 中显示的远程 Excel；程序只读取屏幕像素，不访问链接。", style="Hint.TLabel").grid(row=1, column=0, sticky="w", pady=(5, 14))

        region_box = ttk.LabelFrame(outer, text="1. 截图范围", padding=12)
        region_box.grid(row=2, column=0, sticky="ew", pady=(0, 10))
        region_box.columnconfigure(0, weight=1)
        ttk.Label(region_box, textvariable=self.region_var, style="Hint.TLabel").grid(row=0, column=0, sticky="ew", padx=(0, 10))
        self.select_button = ttk.Button(region_box, text="框选屏幕区域…", command=self._select_region)
        self.select_button.grid(row=0, column=1)
        ttk.Label(region_box, text="请排除冻结表头、行列标题、滚动条和浏览器工具栏。", style="Hint.TLabel").grid(row=1, column=0, columnspan=2, sticky="w", pady=(8, 0))

        opts = ttk.LabelFrame(outer, text="2. 采集设置", padding=12)
        opts.grid(row=3, column=0, sticky="ew", pady=(0, 10))
        for col in (1, 3):
            opts.columnconfigure(col, weight=1)
        labels = [("任务名称", self.name_var), ("翻页等待 (秒)", self.delay_var), ("最大滚动次数", self.max_steps_var), ("每步滚轮格数", self.wheel_var), ("开始倒计时 (秒)", self.countdown_var)]
        for i, (label, var) in enumerate(labels):
            row, col = divmod(i, 2)
            base = col * 2
            ttk.Label(opts, text=label).grid(row=row, column=base, sticky="w", padx=(0, 7), pady=4)
            ttk.Entry(opts, textvariable=var, width=16).grid(row=row, column=base + 1, sticky="ew", padx=(0, 20 if col == 0 else 0), pady=4)
        ttk.Label(opts, text="建议等待 1 秒；远程画面较慢时可调到 2–3 秒。", style="Hint.TLabel").grid(row=3, column=0, columnspan=4, sticky="w", pady=(4, 0))

        log_box = ttk.LabelFrame(outer, text="运行日志", padding=8)
        log_box.grid(row=4, column=0, sticky="nsew")
        log_box.columnconfigure(0, weight=1)
        log_box.rowconfigure(0, weight=1)
        self.log = tk.Text(log_box, height=10, wrap="word", state="disabled", bg="#fbfcfe", relief="flat", padx=8, pady=6)
        self.log.grid(row=0, column=0, sticky="nsew")
        scroll = ttk.Scrollbar(log_box, orient="vertical", command=self.log.yview)
        scroll.grid(row=0, column=1, sticky="ns")
        self.log.configure(yscrollcommand=scroll.set)

        bottom = ttk.Frame(outer)
        bottom.grid(row=5, column=0, sticky="ew", pady=(12, 0))
        bottom.columnconfigure(0, weight=1)
        ttk.Label(bottom, textvariable=self.status_var, style="Status.TLabel").grid(row=0, column=0, sticky="w")
        self.progress = ttk.Progressbar(bottom, mode="indeterminate", length=120)
        self.progress.grid(row=0, column=1, padx=10)
        self.start_button = ttk.Button(bottom, text="开始采集", command=self._start)
        self.start_button.grid(row=0, column=2, padx=(0, 7))
        self.stop_button = ttk.Button(bottom, text="停止 (Esc)", command=self._stop, state="disabled")
        self.stop_button.grid(row=0, column=3, padx=(0, 7))
        ttk.Button(bottom, text="打开 shot 文件夹", command=self._open_shot).grid(row=0, column=4)
        ttk.Label(outer, text="开始后请不要移动浏览器窗口；程序会把鼠标移到选区中心发送滚轮，不会点击单元格。", style="Hint.TLabel").grid(row=6, column=0, sticky="w", pady=(9, 0))

    def _append(self, text: str) -> None:
        stamp = time.strftime("%H:%M:%S")
        self.log.configure(state="normal")
        self.log.insert("end", f"[{stamp}] {text}\n")
        self.log.see("end")
        self.log.configure(state="disabled")

    def _select_region(self) -> None:
        if self.worker and self.worker.is_alive():
            return
        self.master.withdraw()
        try:
            bbox = virtual_screen_bbox()
            if bbox == (0, 0, 0, 0):
                raise RuntimeError("无法取得 Windows 虚拟桌面范围")
            left, top, right, bottom = bbox
            background = grab_region(bbox)
            overlay = tk.Toplevel(self.master)
            overlay.overrideredirect(True)
            overlay.attributes("-topmost", True)
            overlay.geometry(f"{background.width}x{background.height}{left:+d}{top:+d}")
            canvas = tk.Canvas(overlay, width=background.width, height=background.height, highlightthickness=0, cursor="crosshair")
            canvas.pack()
            image = ImageTk.PhotoImage(background)
            canvas.create_image(0, 0, image=image, anchor="nw")
            canvas.create_text(20, 20, anchor="nw", text="拖动框选只包含单元格内容的区域，松开鼠标确认；Esc 取消", fill="#ffeb3b", font=("Microsoft YaHei UI", 14, "bold"))
            start: list[int] = []
            rect: list[int] = []
            result: list[tuple[int, int, int, int]] = []
            def down(event: tk.Event) -> None:
                start[:] = [max(0, min(background.width, event.x)), max(0, min(background.height, event.y))]
                if rect:
                    canvas.delete(rect[0])
                    rect.clear()
            def move(event: tk.Event) -> None:
                if not start:
                    return
                x = max(0, min(background.width, event.x)); y = max(0, min(background.height, event.y))
                if rect:
                    canvas.coords(rect[0], start[0], start[1], x, y)
                else:
                    rect.append(canvas.create_rectangle(start[0], start[1], x, y, outline="#00ff66", width=3))
            def up(event: tk.Event) -> None:
                if not start:
                    return
                x = max(0, min(background.width, event.x)); y = max(0, min(background.height, event.y))
                x0, x1 = sorted((start[0], x)); y0, y1 = sorted((start[1], y))
                if x1 - x0 >= 80 and y1 - y0 >= 80:
                    result.append((left + x0, top + y0, left + x1, top + y1))
                    overlay.destroy()
                else:
                    self._append("选区至少需要 80×80 像素，请重新框选。")
            def cancel(_event: Optional[tk.Event] = None) -> None:
                if overlay.winfo_exists():
                    overlay.destroy()
            canvas.bind("<ButtonPress-1>", down); canvas.bind("<B1-Motion>", move); canvas.bind("<ButtonRelease-1>", up)
            overlay.bind("<Escape>", cancel)
            overlay.focus_force()
            self.master.wait_window(overlay)
            if result:
                self.region = result[0]
                self.region_var.set(f"已选择：{self.region[0]},{self.region[1]} → {self.region[2]},{self.region[3]}（{self.region[2]-self.region[0]}×{self.region[3]-self.region[1]}）")
                self.status_var.set("区域已选择，可以开始采集")
                self._append("已选择截图区域。请将远程 Excel 滚动到首屏后开始。")
        except Exception as exc:
            messagebox.showerror(APP_TITLE, str(exc), parent=self.master)
        finally:
            self.master.deiconify()
            self.master.lift()

    def _get_options(self) -> CaptureOptions | None:
        if not self.region:
            messagebox.showwarning(APP_TITLE, "请先框选远程 Excel 的单元格区域。", parent=self.master)
            return None
        try:
            delay = float(self.delay_var.get() or "1")
            steps = int(self.max_steps_var.get() or "200")
            wheel = int(self.wheel_var.get() or "1")
            countdown = float(self.countdown_var.get() or "5")
        except ValueError:
            messagebox.showerror(APP_TITLE, "等待时间、滚动次数和倒计时必须是数字。", parent=self.master)
            return None
        if not 0.2 <= delay <= 15 or not 1 <= steps <= 5000 or not 1 <= wheel <= 20 or not 0 <= countdown <= 60:
            messagebox.showerror(APP_TITLE, "请检查参数范围：等待 0.2–15 秒、次数 1–5000、滚轮 1–20、倒计时 0–60。", parent=self.master)
            return None
        return CaptureOptions(self.region, self.name_var.get().strip() or "browser_sheet", steps, delay, countdown, wheel)

    def _start(self) -> None:
        options = self._get_options()
        if options is None or (self.worker and self.worker.is_alive()):
            return
        self.cancelled.clear(); self.last_run = None
        self._set_running(True)
        self._append("开始准备；将隐藏本程序窗口，避免它出现在截图中。")
        self.master.withdraw()
        self.worker = threading.Thread(target=self._worker_run, args=(options,), daemon=True, name="browser-screen-capture")
        self.worker.start()
        threading.Thread(target=self._global_escape, daemon=True, name="browser-screen-escape").start()

    def _worker_run(self, options: CaptureOptions) -> None:
        try:
            path = capture_screen(options, log=lambda msg: self.events.put(("log", msg)), cancelled=self.cancelled)
            self.events.put(("done", path))
        except Exception as exc:
            self.events.put(("error", "".join(traceback.format_exception_only(type(exc), exc)).strip()))
            self.events.put(("debug", traceback.format_exc()))

    def _global_escape(self) -> None:
        try:
            import ctypes
            user32 = ctypes.windll.user32
            previous = False
            while self.worker and self.worker.is_alive() and not self.cancelled.is_set():
                pressed = bool(user32.GetAsyncKeyState(0x1B) & 0x8000)
                if pressed and not previous:
                    self.cancelled.set(); self.events.put(("log", "检测到 Esc，正在停止…")); return
                previous = pressed; time.sleep(0.08)
        except (AttributeError, OSError):
            return

    def _stop(self) -> None:
        if self.worker and self.worker.is_alive():
            self.cancelled.set(); self.status_var.set("正在停止…"); self._append("已请求停止，等待当前截图完成…")

    def _set_running(self, running: bool) -> None:
        state = "disabled" if running else "normal"
        self.select_button.configure(state=state); self.start_button.configure(state=state)
        self.stop_button.configure(state="normal" if running else "disabled")
        if running: self.progress.start(12)
        else: self.progress.stop()

    def _poll(self) -> None:
        try:
            while True:
                kind, payload = self.events.get_nowait()
                if kind == "log": self._append(str(payload))
                elif kind == "done":
                    self.last_run = Path(payload); self._finish(False)
                    self._append(f"输出目录：{self.last_run}")
                elif kind == "error":
                    self._finish(True); self._append(f"错误：{payload}"); messagebox.showerror(APP_TITLE, str(payload), parent=self.master)
                elif kind == "debug": self._append(str(payload).rstrip())
        except queue.Empty:
            pass
        try:
            if self.master.winfo_exists(): self.master.after(100, self._poll)
        except tk.TclError:
            return

    def _finish(self, failed: bool) -> None:
        self.master.deiconify(); self.master.lift(); self._set_running(False); self.worker = None
        self.status_var.set("运行失败" if failed else "采集完成")
        if self.closing: self.master.destroy()

    def _open_shot(self) -> None:
        folder = Path(__file__).resolve().parents[1] / "shot"
        if getattr(__import__("sys"), "frozen", False):
            import sys
            folder = Path(sys.executable).resolve().parent / "shot"
        folder.mkdir(parents=True, exist_ok=True)
        try: os.startfile(str(folder))
        except AttributeError: pass

    def _close(self) -> None:
        if self.worker and self.worker.is_alive():
            self.closing = True; self.cancelled.set(); self.master.withdraw(); return
        self.master.destroy()


def run_gui() -> None:
    enable_dpi_awareness()
    root = tk.Tk()
    BrowserScreenApp(root)
    root.mainloop()


__all__ = ["BrowserScreenApp", "run_gui"]

