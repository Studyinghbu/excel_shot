"""Small Windows screen I/O helpers for the browser/KVM capture mode."""
from __future__ import annotations

import ctypes
import sys
import time

from PIL import Image, ImageGrab


def enable_dpi_awareness() -> None:
    """Make Tk coordinates and ImageGrab coordinates use the same pixels."""
    if sys.platform != "win32":
        return
    try:
        user32 = ctypes.windll.user32
        if user32.SetProcessDpiAwarenessContext(ctypes.c_void_p(-4)):
            return
    except (AttributeError, OSError):
        pass
    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(2)
    except (AttributeError, OSError):
        try:
            ctypes.windll.user32.SetProcessDPIAware()
        except (AttributeError, OSError):
            pass


def virtual_screen_bbox() -> tuple[int, int, int, int]:
    """Return Windows' physical-pixel virtual desktop bbox."""
    if sys.platform != "win32":
        return (0, 0, 0, 0)
    u = ctypes.windll.user32
    left, top = int(u.GetSystemMetrics(76)), int(u.GetSystemMetrics(77))
    width, height = int(u.GetSystemMetrics(78)), int(u.GetSystemMetrics(79))
    return left, top, left + width, top + height


def grab_region(region: tuple[int, int, int, int]) -> Image.Image:
    """Capture a physical-pixel bbox, including a monitor left of the primary."""
    left, top, right, bottom = (int(x) for x in region)
    if right <= left or bottom <= top:
        raise ValueError("截图区域必须有正的宽度和高度")
    try:
        return ImageGrab.grab(bbox=(left, top, right, bottom), all_screens=True)
    except TypeError:
        return ImageGrab.grab(bbox=(left, top, right, bottom))


def scroll_at(region: tuple[int, int, int, int], notches: int = 1) -> None:
    """Place the pointer in the cell region and send a bounded wheel step."""
    if sys.platform != "win32":
        raise RuntimeError("屏幕滚动模式只能在 Windows 上运行")
    left, top, right, bottom = (int(x) for x in region)
    x = (left + right) // 2
    y = (top + bottom) // 2
    user32 = ctypes.windll.user32
    if not user32.SetCursorPos(x, y):
        raise RuntimeError("无法定位鼠标到截图区域")

    class MOUSEINPUT(ctypes.Structure):
        _fields_ = [("dx", ctypes.c_long), ("dy", ctypes.c_long),
                    ("mouseData", ctypes.c_ulong), ("dwFlags", ctypes.c_ulong),
                    ("time", ctypes.c_ulong), ("dwExtraInfo", ctypes.POINTER(ctypes.c_ulong))]

    class INPUT(ctypes.Structure):
        _fields_ = [("type", ctypes.c_ulong), ("mi", MOUSEINPUT)]

    delta = max(-20, min(20, int(notches))) * 120
    inp = INPUT(0, MOUSEINPUT(0, 0, ctypes.c_ulong(delta), 0x0800, 0, None))
    if user32.SendInput(1, ctypes.byref(inp), ctypes.sizeof(INPUT)) != 1:
        raise RuntimeError("无法向远程桌面发送滚轮输入")


def wait_settle(seconds: float, cancelled=None) -> bool:
    """Sleep in short pieces so Esc can stop a long remote redraw delay."""
    deadline = time.monotonic() + max(0.0, float(seconds))
    while time.monotonic() < deadline:
        if cancelled is not None and cancelled.is_set():
            return False
        time.sleep(min(0.05, max(0.0, deadline - time.monotonic())))
    return not (cancelled is not None and cancelled.is_set())


__all__ = ["enable_dpi_awareness", "virtual_screen_bbox", "grab_region", "scroll_at", "wait_settle"]
