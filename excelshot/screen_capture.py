"""Controller for browser/KVM based Excel screenshots.

The controller never opens a workbook and never accesses a browser URL.  It works
with the rectangle the user selected on the desktop, scrolls inside that rectangle,
and writes the first viewport plus only each newly exposed strip.  A manifest keeps
all measured shifts and OCR/alignment warnings for review.
"""
from __future__ import annotations

import hashlib
import json
import re
import sys
import time
from dataclasses import asdict, dataclass
from datetime import datetime
from pathlib import Path
from threading import Event
from typing import Callable, Optional

from PIL import Image

from .ocr import OfflineOCR
from .screen_io import enable_dpi_awareness, grab_region, scroll_at, wait_settle
from .screen_alignment import align_vertical, extract_novel_strip

LOG = Callable[[str], None]


@dataclass
class CaptureOptions:
    region: tuple[int, int, int, int]
    name: str = "browser_sheet"
    max_steps: int = 200
    settle_seconds: float = 1.0
    countdown_seconds: float = 5.0
    wheel_notches: int = 1
    output_dir: Optional[str] = None


def _root_dir() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parents[1]


def _safe_name(value: str) -> str:
    value = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", str(value)).strip(" .")
    return (value or "browser_sheet")[:80]


def _log(log: LOG, message: str) -> None:
    try:
        log(message)
    except Exception:
        pass


def _sha256(image: Image.Image) -> str:
    return hashlib.sha256(image.convert("RGB").tobytes()).hexdigest()


def _save(image: Image.Image, path: Path) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    image.convert("RGB").save(path, "PNG", optimize=False)


def _sleep_countdown(seconds: float, log: LOG, cancelled: Event) -> bool:
    seconds = max(0.0, float(seconds))
    if seconds <= 0:
        return True
    _log(log, f"请把远程 Excel 滚动到起始位置；{seconds:g} 秒后开始，期间不要移动窗口…")
    end = time.monotonic() + seconds
    last = None
    while time.monotonic() < end:
        if cancelled.is_set():
            return False
        remaining = max(0, int(end - time.monotonic() + 0.999))
        if remaining != last:
            _log(log, f"开始倒计时：{remaining} 秒")
            last = remaining
        time.sleep(min(0.1, max(0.0, end - time.monotonic())))
    return True


def capture_screen(options: CaptureOptions, log: LOG = print, cancelled: Optional[Event] = None) -> Path:
    """Capture a scrollable cell-only rectangle and return its run directory."""
    enable_dpi_awareness()
    cancelled = cancelled or Event()
    left, top, right, bottom = tuple(int(x) for x in options.region)
    width, height = right - left, bottom - top
    if width < 80 or height < 80:
        raise ValueError("截图区域太小；请只框选包含单元格内容的较大区域")
    max_steps = int(options.max_steps)
    if max_steps < 1 or max_steps > 5000:
        raise ValueError("最大滚动次数应在 1 到 5000 之间")
    settle = float(options.settle_seconds)
    if settle < 0.1 or settle > 30:
        raise ValueError("翻页等待应在 0.1 到 30 秒之间")
    wheel = max(1, min(20, int(options.wheel_notches)))

    root = Path(options.output_dir) if options.output_dir else _root_dir() / "shot"
    root.mkdir(parents=True, exist_ok=True)
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S_%f")
    run_dir = root / f"browser_{stamp}_{_safe_name(options.name)}"
    run_dir.mkdir(parents=True, exist_ok=False)
    manifest: dict = {
        "mode": "browser_screen",
        "started_at": datetime.now().isoformat(timespec="seconds"),
        "region": [left, top, right, bottom],
        "options": asdict(options),
        "pages": [],
        "warnings": [],
    }
    manifest["options"]["output_dir"] = str(root)

    def finish() -> None:
        (run_dir / "manifest.json").write_text(json.dumps(manifest, ensure_ascii=False, indent=2, default=str), encoding="utf-8")

    if not _sleep_countdown(options.countdown_seconds, log, cancelled):
        manifest["status"] = "cancelled_before_first_capture"
        finish()
        return run_dir

    try:
        ocr = OfflineOCR()
        _log(log, "离线 OCR 已加载；将用 OCR 和画面纹理验证相邻截图。")
    except Exception as exc:
        ocr = None
        manifest["warnings"].append(f"OCR 加载失败：{exc}")
        _log(log, f"OCR 加载失败，将依靠图像重叠检查：{exc}")

    current = grab_region((left, top, right, bottom)).convert("RGB")
    previous_ocr = ocr.recognize(current) if ocr else None
    first_path = run_dir / "page_0001_start.png"
    _save(current, first_path)
    manifest["pages"].append({
        "index": 1, "file": first_path.name, "kind": "viewport",
        "sha256": _sha256(current), "novel_height": current.height,
        "region": [left, top, right, bottom],
        "ocr_text": (previous_ocr or {}).get("text", ""),
    })
    _log(log, f"第 1 页已保存：{first_path.name}")

    previous = current
    page_index = 1
    stitch = current.copy()
    no_motion = 0
    try:
        for step in range(1, max_steps + 1):
            if cancelled.is_set():
                manifest["status"] = "cancelled"
                break
            scroll_at((left, top, right, bottom), -wheel)
            if not wait_settle(settle, cancelled):
                manifest["status"] = "cancelled"
                break
            current = grab_region((left, top, right, bottom)).convert("RGB")
            if _sha256(current) == _sha256(previous):
                no_motion += 1
                _log(log, f"第 {step} 次滚动画面未变化（{no_motion}/2）；远程桌面可能已到末尾。")
                if no_motion >= 2:
                    manifest["status"] = "end_of_content_or_no_focus"
                    manifest["warnings"].append("连续两次滚动后画面没有变化；已停止以避免重复截图。")
                    break
                continue
            no_motion = 0
            current_ocr = ocr.recognize(current) if ocr else None
            alignment = align_vertical(
                previous, current,
                ocr_previous=previous_ocr,
                ocr_current=current_ocr,
            )
            if not alignment.aligned or alignment.shift is None:
                reason = alignment.reason or alignment.status
                manifest["warnings"].append(f"第 {step} 次滚动未能可靠对齐：{reason}")
                _log(log, f"第 {step} 次滚动未能可靠对齐：{reason}；停止并保留已完成截图。")
                manifest["status"] = "alignment_failed"
                break
            dy = int(alignment.shift)
            if dy <= 0 or dy >= height:
                manifest["warnings"].append(f"第 {step} 次滚动得到异常位移 {dy}px")
                manifest["status"] = "alignment_failed"
                break
            # Current top through current top+overlap repeats the previous bottom.
            from PIL import Image
            novel = Image.fromarray(extract_novel_strip(current, alignment), mode="RGB")
            if novel.height < 12:
                manifest["warnings"].append(f"第 {step} 次滚动新内容只有 {novel.height}px，停止。")
                manifest["status"] = "alignment_failed"
                break
            page_index += 1
            path = run_dir / f"page_{page_index:04d}_new_{novel.height:04d}px.png"
            _save(novel, path)
            stitch = Image.new("RGB", (width, stitch.height + novel.height), "white")
            # Rebuilding avoids an in-place alias when Pillow versions share buffers.
            stitch.paste(Image.open(run_dir / "page_0001_start.png").convert("RGB"), (0, 0)) if page_index == 2 else None
            if page_index == 2:
                prior = Image.open(run_dir / "page_0001_start.png").convert("RGB")
                stitch = Image.new("RGB", (width, prior.height + novel.height), "white")
                stitch.paste(prior, (0, 0))
                stitch.paste(novel, (0, prior.height))
                prior.close()
            else:
                old = Image.open(run_dir / "stitched.png").convert("RGB")
                stitch = Image.new("RGB", (width, old.height + novel.height), "white")
                stitch.paste(old, (0, 0))
                stitch.paste(novel, (0, old.height))
                old.close()
            _save(stitch, run_dir / "stitched.png")
            warning = ""
            if ocr and previous_ocr and current_ocr:
                comparison = ocr.compare_adjacent(previous_ocr, current_ocr)
                warning = str(comparison.get("warning", ""))
                if warning:
                    # OCR duplicates can be legitimate identical cell values; this is a review signal.
                    manifest["warnings"].append(f"第 {step} 次滚动：{warning}")
                    _log(log, "  OCR 发现边界文字重复；已按像素对齐只保留新区域，并记录提示。")
            record = {
                "index": page_index, "file": path.name, "kind": "novel_strip",
                "step": step, "shift_y": dy, "overlap_height": height - dy,
                "novel_height": novel.height, "sha256": _sha256(novel),
                "alignment": asdict(alignment), "ocr_warning": warning,
                "ocr_text": (current_ocr or {}).get("text", ""),
            }
            manifest["pages"].append(record)
            _log(log, f"第 {page_index} 页已保存：{path.name}（滚动 {dy}px，连接无重复像素）")
            previous, previous_ocr = current, current_ocr
        else:
            manifest["status"] = "max_steps_reached"
            manifest["warnings"].append("达到最大滚动次数；如表格尚未结束，请提高最大滚动次数后再运行。")
    finally:
        if "status" not in manifest:
            manifest["status"] = "completed"
        manifest["finished_at"] = datetime.now().isoformat(timespec="seconds")
        finish()
    _log(log, f"完成：{len(manifest['pages'])} 个连续截图，输出目录：{run_dir}")
    return run_dir


__all__ = ["CaptureOptions", "capture_screen"]


