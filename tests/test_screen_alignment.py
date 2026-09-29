from __future__ import annotations

import numpy as np

from excelshot.screen_alignment import AlignmentResult, align_vertical, extract_novel_strip
from PIL import Image


def _source(height: int = 520, width: int = 180) -> np.ndarray:
    """Deterministic worksheet-like texture: cells, borders, and text strokes."""
    image = np.full((height, width, 3), 242, dtype=np.uint8)
    # Grid rows/columns
    for y in range(0, height, 31):
        image[y : y + 2, :, :] = 120
    for x in range(0, width, 45):
        image[:, x : x + 2, :] = 135
    # Distinct row labels and short text-like strokes.
    for row in range(8, height, 31):
        length = 8 + (row * 7) % 30
        image[row : row + 5, 7 : 7 + length, :] = 35
        image[row + 8 : row + 12, 11 : 11 + max(4, length // 2), :] = 75
    return image


def test_alignment_returns_downward_shift_and_new_bottom_strip():
    source = _source()
    shift = 37
    previous = source[:260]
    current = source[shift : shift + 260]
    result = align_vertical(previous, current)
    assert result.aligned
    assert abs(result.shift - shift) <= 1
    assert result.overlap_height == 260 - result.shift
    strip = extract_novel_strip(current, result)
    assert strip.shape[0] == result.shift
    np.testing.assert_array_equal(strip, current[260 - result.shift :])


def test_identical_capture_is_unchanged():
    image = _source(300)
    result = align_vertical(image, image.copy())
    assert result.status == "unchanged"
    assert result.shift is None


def test_periodic_grid_is_rejected_instead_of_picking_a_duplicate_page():
    image = np.full((300, 180, 3), 245, dtype=np.uint8)
    for y in range(0, 300, 20):
        image[y : y + 2, :, :] = 90
    # A 20-pixel translation is exactly the same periodic viewport.
    result = align_vertical(image[:240], image[20:260])
    assert not result.aligned
    assert result.status in {"unchanged", "ambiguous", "no_overlap"}


def test_random_different_capture_is_not_accepted():
    rng = np.random.default_rng(42)
    first = rng.integers(0, 255, (260, 180, 3), dtype=np.uint8)
    second = rng.integers(0, 255, first.shape, dtype=np.uint8)
    result = align_vertical(first, second)
    assert not result.aligned
    assert result.status in {"ambiguous", "scaled", "no_overlap"}


def test_ocr_agreement_disambiguates_a_valid_shift():
    source = _source()
    shift = 31
    previous, current = source[:260], source[shift : shift + 260]
    before_ocr = {
        "items": [
            {"text": "unique-one", "ymin": 90, "ymax": 100},
            {"text": "unique-two", "ymin": 180, "ymax": 190},
        ]
    }
    after_ocr = {
        "items": [
            {"text": "unique-one", "ymin": 59, "ymax": 69},
            {"text": "unique-two", "ymin": 149, "ymax": 159},
        ]
    }
    result = align_vertical(previous, current, ocr_previous=before_ocr, ocr_current=after_ocr)
    assert result.aligned
    assert result.ocr_shift == shift


def test_ocr_disagreement_rejects_pixel_coincidence():
    source = _source()
    shift = 31
    previous, current = source[:260], source[shift : shift + 260]
    before_ocr = {"items": [{"text": "only", "ymin": 90, "ymax": 100}]}
    after_ocr = {"items": [{"text": "only", "ymin": 10, "ymax": 20}]}
    result = align_vertical(previous, current, ocr_previous=before_ocr, ocr_current=after_ocr)
    assert not result.aligned
    assert result.status == "ambiguous"


def test_invalid_dimensions_are_reported():
    result = align_vertical(np.zeros((64, 64, 3), np.uint8), np.zeros((65, 64, 3), np.uint8))
    assert result.status == "invalid"




def test_vertical_zoom_change_is_rejected():
    source = _source()
    previous = source[:260]
    # A browser zoom changes row geometry, so no one-pixel vertical shift is
    # trusted even though both captures have the same outer dimensions.
    scaled = np.asarray(Image.fromarray(source).resize((180, int(source.shape[0] * 1.10))))
    current = scaled[37 : 37 + 260]
    result = align_vertical(previous, current)
    assert not result.aligned
    assert result.status in {"scaled", "ambiguous"}
