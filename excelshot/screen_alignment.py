"""Pixel/OCR alignment for screenshots of a worksheet shown through a browser/KVM.

The routine accepts two equal-sized desktop captures. For a downward scroll d,
previous[d:] should equal current[:-d]; shift is that positive d. OCR is an
optional independent signal and is never required.
"""
from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Any, Mapping, Sequence

import numpy as np


class AlignmentError(ValueError):
    """Raised only by strict callers; align_vertical itself returns evidence."""


@dataclass(frozen=True)
class AlignmentResult:
    status: str
    shift: int | None = None
    overlap_height: int = 0
    score: float = float("inf")
    margin: float = 0.0
    texture: float = 0.0
    ocr_shift: float | None = None
    reason: str = ""
    shape: tuple[int, int] | None = None

    @property
    def aligned(self) -> bool:
        return self.status == "aligned" and self.shift is not None

    @property
    def is_aligned(self) -> bool:
        return self.aligned

    @property
    def novel_start(self) -> int | None:
        if not self.aligned or self.shape is None:
            return None
        return self.shape[0] - int(self.shift or 0)

    @property
    def new_region(self) -> tuple[int, int, int, int] | None:
        start = self.novel_start
        if start is None or self.shape is None:
            return None
        return 0, start, self.shape[1], self.shape[0]


def _as_array(image: Any) -> np.ndarray:
    if isinstance(image, (str, Path)):
        from PIL import Image
        with Image.open(image) as opened:
            return np.asarray(opened.convert("RGB"), dtype=np.uint8).copy()
    if hasattr(image, "convert") and hasattr(image, "size"):
        return np.asarray(image.convert("RGB"), dtype=np.uint8).copy()
    array = np.asarray(image)
    if array.ndim == 2:
        array = np.repeat(array[:, :, None], 3, axis=2)
    elif array.ndim == 3 and array.shape[2] == 4:
        array = array[:, :, :3]
    if array.ndim != 3 or array.shape[2] not in (1, 3):
        raise ValueError("截图必须是 HxW、HxWx3 或 HxWx4 图像")
    if array.shape[2] == 1:
        array = np.repeat(array, 3, axis=2)
    if np.issubdtype(array.dtype, np.floating):
        if not np.isfinite(array).all():
            raise ValueError("截图含有 NaN 或无穷值")
        if float(array.max(initial=0)) <= 1.0:
            array = array * 255.0
    return np.clip(array, 0, 255).astype(np.uint8, copy=False)


def _crop(array: np.ndarray, roi: Sequence[int] | None) -> np.ndarray:
    if roi is None:
        return array
    if len(roi) != 4:
        raise ValueError("roi 必须是 (left, top, right, bottom)")
    left, top, right, bottom = (int(v) for v in roi)
    h, w = array.shape[:2]
    if not (0 <= left < right <= w and 0 <= top < bottom <= h):
        raise ValueError("roi 超出截图范围")
    return array[top:bottom, left:right]


def _features(array: np.ndarray, bins: int = 96) -> np.ndarray:
    from PIL import Image
    gray = np.dot(array[:, :, :3].astype(np.float32), [0.299, 0.587, 0.114])
    width = min(max(24, bins), gray.shape[1])
    reduced = np.asarray(
        Image.fromarray(np.clip(gray, 0, 255).astype(np.uint8)).resize(
            (width, gray.shape[0]), Image.Resampling.BILINEAR
        ), dtype=np.float32)
    edge = np.diff(reduced, axis=1, prepend=reduced[:, :1])
    result = np.concatenate((reduced, edge), axis=1)
    result -= result.mean(axis=1, keepdims=True)
    result /= np.maximum(result.std(axis=1, keepdims=True), 3.0)
    return result.astype(np.float32, copy=False)


def _ocr_shift(previous: Mapping[str, Any] | None, current: Mapping[str, Any] | None) -> float | None:
    if not previous or not current:
        return None

    def items(value: Mapping[str, Any]) -> dict[str, list[float]]:
        found: dict[str, list[float]] = {}
        for item in value.get("items") or []:
            text = "".join(str(item.get("text", "")).split()).casefold()
            if not text:
                continue
            try:
                if "ymin" in item and "ymax" in item:
                    center = (float(item["ymin"]) + float(item["ymax"])) / 2.0
                else:
                    box = item.get("box") or []
                    center = sum(float(p[1]) for p in box) / len(box)
            except (TypeError, ValueError, ZeroDivisionError):
                continue
            found.setdefault(text, []).append(center)
        return found

    before, after = items(previous), items(current)
    deltas: list[float] = []
    for text, ys in before.items():
        # Repeated labels are excluded; otherwise a periodic grid can look
        # confidently aligned at the wrong row.
        if len(ys) == 1 and len(after.get(text, ())) == 1:
            delta = ys[0] - after[text][0]
            if delta >= -3.0:
                deltas.append(delta)
    if not deltas:
        return None
    median = float(np.median(deltas))
    inliers = [d for d in deltas if abs(d - median) <= max(4.0, abs(median) * 0.08)]
    if len(inliers) < max(1, min(2, len(deltas))):
        return None
    return float(np.median(inliers))


def _scores(before: np.ndarray, after: np.ndarray, low: int, high: int) -> list[tuple[int, float]]:
    result: list[tuple[int, float]] = []
    for shift in range(low, high + 1):
        error = np.mean((before[shift:] - after[:-shift]) ** 2, axis=1)
        result.append((shift, float(np.quantile(error, 0.75))))
    return result


def _raw_error(before: np.ndarray, after: np.ndarray, shift: int) -> float:
    """Downsampled grayscale pixel tie breaker for repeated grid lines."""
    a = before.astype(np.float32)
    b = after.astype(np.float32)
    if a.shape[1] > 160:
        from PIL import Image
        width = 160
        a = np.asarray(Image.fromarray(a.astype(np.uint8)).resize((width, a.shape[0]), Image.Resampling.BILINEAR), dtype=np.float32)
        b = np.asarray(Image.fromarray(b.astype(np.uint8)).resize((width, b.shape[0]), Image.Resampling.BILINEAR), dtype=np.float32)
    return float(np.mean(np.abs(a[shift:] - b[:-shift])))

def align_vertical(
    previous: Any,
    current: Any,
    ocr_previous: Mapping[str, Any] | None = None,
    ocr_current: Mapping[str, Any] | None = None,
    *,
    roi: Sequence[int] | None = None,
    min_shift: int = 8,
    max_shift_ratio: float = 0.85,
    min_overlap_ratio: float = 0.20,
) -> AlignmentResult:
    """Estimate a positive downward scroll, rejecting weak/periodic matches."""
    try:
        before = _crop(_as_array(previous), roi)
        after = _crop(_as_array(current), roi)
    except (TypeError, ValueError, OSError) as exc:
        return AlignmentResult("invalid", reason=str(exc))
    if before.shape != after.shape:
        return AlignmentResult("invalid", reason="两张截图尺寸不同")
    height, width = before.shape[:2]
    shape = (height, width)
    if height < 32 or width < 32:
        return AlignmentResult("invalid", shape=shape, reason="截图区域过小")

    change = float(np.mean(np.abs(before.astype(np.int16) - after.astype(np.int16))))
    if change <= 1.25:
        return AlignmentResult("unchanged", shape=shape, score=change, reason="截图没有发生可见变化")

    try:
        first = _features(before)
        second = _features(after)
    except Exception as exc:
        return AlignmentResult("invalid", shape=shape, reason=f"图像特征提取失败: {exc}")
    texture = float(np.median(np.std(first, axis=1)))
    if texture < 0.30:
        return AlignmentResult("no_overlap", shape=shape, texture=texture, reason="重叠区域纹理不足")

    low = max(1, int(min_shift))
    high = min(height - 1, int(height * float(max_shift_ratio)))
    high = min(high, height - max(1, int(height * float(min_overlap_ratio))))
    if high < low:
        return AlignmentResult("invalid", shape=shape, texture=texture, reason="滚动范围参数无效")

    ordered = sorted(_scores(first, second, low, high), key=lambda pair: pair[1])
    best_shift, best_score = ordered[0]
    # Use raw pixels to resolve ties caused by repeated horizontal grid lines.
    # If raw errors tie too, keep the ambiguity and stop rather than guessing.
    feature_floor = best_score + max(1e-5, abs(best_score) * 0.20)
    candidate_pool = [shift for shift, score in ordered if score <= feature_floor][:32]
    raw_pairs = sorted(((_raw_error(before, after, shift), shift) for shift in candidate_pool), key=lambda pair: pair[0])
    raw_best, raw_shift = raw_pairs[0]
    if len(raw_pairs) > 1:
        raw_second = raw_pairs[1][0]
        if raw_best + 0.10 < raw_second:
            best_shift = raw_shift
            best_score = next(score for shift, score in ordered if shift == best_shift)
    nearby = max(3, height // 100)
    competitor = next((pair for pair in ordered[1:] if abs(pair[0] - best_shift) > nearby), None)
    second_score = competitor[1] if competitor else float("inf")
    margin = ((second_score - best_score) / max(second_score, 1e-6)) if np.isfinite(second_score) else 1.0
    ocr_delta = _ocr_shift(ocr_previous, ocr_current)
    # When several pixel candidates tie (periodic grid lines), a unique OCR
    # displacement may choose the matching candidate.  Only accept it when
    # its image score is effectively tied; OCR cannot override a bad image.
    if ocr_delta is not None:
        ocr_candidate = min(ordered, key=lambda pair: abs(pair[0] - ocr_delta))
        if (abs(float(ocr_candidate[0]) - ocr_delta) <= max(4.0, height * 0.02)
                and ocr_candidate[1] <= best_score + max(0.005, abs(best_score) * 0.25)):
            best_shift, best_score = ocr_candidate

    errors = np.mean((first[best_shift:] - second[:-best_shift]) ** 2, axis=1)
    q75 = float(np.quantile(errors, 0.75))
    if not np.isfinite(best_score) or q75 > 0.03:
        return AlignmentResult("scaled", shape=shape, score=best_score, margin=margin,
                               texture=texture, ocr_shift=ocr_delta,
                               reason="未找到可信的像素重叠，可能发生缩放或页面变化")
    if competitor is not None and margin < 0.055:
        # A unique OCR boundary match may break a repeated grid tie. Without
        # it, stopping is safer than choosing a duplicate row by pixels alone.
        if ocr_delta is None or abs(float(best_shift) - ocr_delta) > max(8.0, height * 0.04):
            return AlignmentResult("ambiguous", shape=shape, score=best_score, margin=margin,
                                   texture=texture, ocr_shift=ocr_delta,
                                   reason="多个滚动位移同样匹配，疑似重复网格")
    if ocr_delta is not None and abs(float(best_shift) - ocr_delta) > max(8.0, height * 0.04):
        return AlignmentResult("ambiguous", shape=shape, score=best_score, margin=margin,
                               texture=texture, ocr_shift=ocr_delta,
                               reason="像素位移与 OCR 位移不一致")
    if margin < 0.015 and ocr_delta is None:
        return AlignmentResult("ambiguous", shape=shape, score=best_score, margin=margin,
                               texture=texture, reason="位移置信度不足")
    return AlignmentResult("aligned", shift=int(best_shift), overlap_height=height - best_shift,
                           score=best_score, margin=margin, texture=texture,
                           ocr_shift=ocr_delta, shape=shape, reason="像素重叠和滚动方向通过校验")


def extract_novel_strip(image: Any, alignment: AlignmentResult) -> np.ndarray:
    """Return current[height-shift:] after a verified alignment."""
    if not alignment.aligned or alignment.shift is None:
        raise AlignmentError("只有 aligned 结果才能提取新区域")
    array = _as_array(image)
    start = alignment.novel_start
    if start is None or start <= 0 or start >= array.shape[0]:
        raise AlignmentError("位移不在截图范围内")
    return array[start:, :, :].copy()


__all__ = ["AlignmentError", "AlignmentResult", "align_vertical", "extract_novel_strip"]












