"""Small offline OCR adapter used to audit adjacent screenshots.

RapidOCR ships its PP-OCRv4 detection/recognition models in the wheel.  This
module never downloads a model and does not send the workbook or images over
the network.  OCR is deliberately a *review signal*: identical text in two
different worksheet cells is valid and is never removed.
"""
from __future__ import annotations

import re
from pathlib import Path
from typing import Any


class OfflineOCR:
    def __init__(self) -> None:
        try:
            from rapidocr_onnxruntime import RapidOCR
        except Exception as exc:  # pragma: no cover - exercised on unpackaged hosts
            raise RuntimeError("离线 OCR 模块未安装") from exc
        # RapidOCR loads the bundled ONNX models lazily.  Keeping one instance
        # avoids the 10–20 MB model initialization on every page.
        self._engine = RapidOCR()

    @staticmethod
    def _normalise(value: str) -> str:
        return re.sub(r"\s+", "", str(value or "")).casefold()

    def recognize(self, image: Any) -> dict[str, Any]:
        """Return JSON-safe OCR boxes and text for a path/PIL image/ndarray."""
        result, elapsed = self._engine(image)
        items: list[dict[str, Any]] = []
        if result:
            for row in result:
                if not row or len(row) < 3:
                    continue
                box, text, score = row[0], str(row[1]), float(row[2])
                points = [[float(point[0]), float(point[1])] for point in box]
                ymin = min((point[1] for point in points), default=0.0)
                ymax = max((point[1] for point in points), default=0.0)
                items.append({
                    "box": points, "text": text, "confidence": score,
                    "ymin": ymin, "ymax": ymax,
                })
        items.sort(key=lambda item: (item["ymin"], item["box"][0][0] if item["box"] else 0))
        return {
            "text": "\n".join(item["text"] for item in items),
            "items": items,
            "elapsed": [float(x) for x in (elapsed or [])],
        }

    def compare_adjacent(self, previous: dict[str, Any], current: dict[str, Any]) -> dict[str, Any]:
        """Detect likely duplicated boundary lines without deleting any page."""
        prev_items = list(previous.get("items") or [])
        curr_items = list(current.get("items") or [])
        if not prev_items or not curr_items:
            return {"duplicate": False, "warning": ""}
        # OCR orders boxes top-to-bottom.  Compare a small boundary window so
        # a repeated value in the middle of a page is not called an overlap.
        previous_boundary = {
            self._normalise(item.get("text", ""))
            for item in prev_items[-4:]
            if self._normalise(item.get("text", ""))
        }
        current_boundary = {
            self._normalise(item.get("text", ""))
            for item in curr_items[:4]
            if self._normalise(item.get("text", ""))
        }
        duplicate = bool(previous_boundary & current_boundary)
        if duplicate:
            return {
                "duplicate": True,
                "warning": "相邻页边界检测到相同 OCR 文字；程序仍保留两页，请核对表格是否有重复行。",
            }
        return {"duplicate": False, "warning": ""}


__all__ = ["OfflineOCR"]

