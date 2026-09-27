"""Pure Python helpers for paginating Excel rows.

The functions in this module deliberately know nothing about Excel COM.  Row
numbers are one-based and ranges are inclusive, which matches Excel's row
numbering and makes the returned spans directly usable by a screenshot worker.
"""

from __future__ import annotations

from dataclasses import dataclass
import math
import re
from collections.abc import Mapping, Sequence
from typing import Iterator


@dataclass(frozen=True, order=True)
class RowSpan:
    """An inclusive row range.

    ``RowSpan(3, 7)`` represents rows 3 through 7.  The class intentionally
    permits row zero so it can also be used by callers while building a plan;
    :func:`plan_page` validates that its requested range is positive.
    """

    start: int
    end: int

    def __post_init__(self) -> None:
        if not isinstance(self.start, int) or isinstance(self.start, bool):
            raise TypeError("RowSpan.start must be an integer")
        if not isinstance(self.end, int) or isinstance(self.end, bool):
            raise TypeError("RowSpan.end must be an integer")
        if self.start > self.end:
            raise ValueError("RowSpan.start cannot be greater than end")

    def __len__(self) -> int:
        """Return the number of rows in the inclusive span."""

        return self.end - self.start + 1

    def contains(self, row: int) -> bool:
        return self.start <= row <= self.end


def _as_height(value: object, row: int) -> float:
    """Convert an Excel row height to a non-negative finite float."""

    # Some COM wrappers expose an unset height as None.  It has no visible
    # extent for pagination purposes, just like a hidden row.
    if value is None:
        return 0.0
    try:
        result = float(value)
    except (TypeError, ValueError) as exc:
        raise TypeError(f"row {row} height must be a number") from exc
    if not math.isfinite(result) or result < 0:
        raise ValueError(f"row {row} height must be finite and non-negative")
    return result


def _height_getter(
    row_heights: Mapping[int, object] | Sequence[object],
    start: int,
    end: int,
):
    """Build a 1-based row-height lookup for the requested range.

    Mappings are keyed by Excel row number.  Sequences are normally treated as
    a 1-based worksheet slice (index ``row - 1``); a sequence whose length is
    exactly the requested range is also accepted as a local slice, which is
    convenient when paginating a preselected section.
    """

    if isinstance(row_heights, Mapping):
        missing = [row for row in range(start, end + 1) if row not in row_heights]
        if missing:
            raise ValueError(f"row heights missing row(s): {missing[:5]}")

        def get_height(row: int) -> float:
            return _as_height(row_heights[row], row)

        return get_height

    if isinstance(row_heights, (str, bytes, bytearray)):
        raise TypeError("row_heights must be a mapping or a numeric sequence")
    try:
        size = len(row_heights)
    except TypeError as exc:
        raise TypeError("row_heights must be a mapping or a numeric sequence") from exc

    local_slice = size < end and size == end - start + 1
    if not local_slice and size < end:
        raise ValueError(
            f"row_heights sequence has {size} item(s), but row {end} is required"
        )

    def get_height(row: int) -> float:
        index = row - start if local_slice else row - 1
        try:
            value = row_heights[index]
        except (IndexError, KeyError, TypeError) as exc:
            raise ValueError(f"row heights missing row {row}") from exc
        return _as_height(value, row)

    return get_height


def _normalise_merged_spans(
    merged_spans: Sequence[RowSpan] | None,
    start: int,
    end: int,
) -> list[RowSpan]:
    if merged_spans is None:
        return []
    normalised: list[RowSpan] = []
    for span in merged_spans:
        if not isinstance(span, RowSpan):
            try:
                span = RowSpan(span.start, span.end)  # type: ignore[attr-defined]
            except AttributeError as exc:
                raise TypeError("merged_spans must contain RowSpan values") from exc
        # A merge that intersects the target must be wholly inside it: a page
        # cannot represent the part of a merge outside the requested range.
        if span.end < start or span.start > end:
            continue
        if span.start < start or span.end > end:
            raise ValueError(
                f"合并单元格 {span.start}:{span.end} 超出目标行范围 {start}:{end}，无法完整截图"
            )
        normalised.append(span)

    normalised.sort(key=lambda item: (item.start, item.end))
    for previous, current in zip(normalised, normalised[1:]):
        if current.start <= previous.end:
            raise ValueError(
                f"合并单元格范围重叠: {previous.start}:{previous.end} 与 "
                f"{current.start}:{current.end}"
            )
    return normalised


def plan_page(
    row_heights: Mapping[int, object] | Sequence[object],
    start: int,
    end: int,
    max_height: float,
    merged_spans: Sequence[RowSpan] | None = None,
) -> RowSpan:
    """Plan one greedy page of complete rows.

    Rows are consumed from ``start`` through the returned ``end``.  Every row
    has a place in exactly one page when callers repeatedly invoke this
    function with ``next_start = previous.end + 1``.  A hidden row is a row
    whose height is zero; it is still consumed so it cannot be repeated on the
    next page.  Merged ranges are atomic and are moved wholly to the next page
    when they do not fit in the remaining viewport.

    A merged range whose total height exceeds ``max_height`` raises
    :class:`ValueError`, with a Chinese message suitable for displaying to an
    operator.  A non-merged row that is taller than the viewport is returned as
    a one-row page; cutting that row would violate the complete-row guarantee.
    """

    if not isinstance(start, int) or isinstance(start, bool):
        raise TypeError("start must be an integer")
    if not isinstance(end, int) or isinstance(end, bool):
        raise TypeError("end must be an integer")
    if start < 1 or end < 1:
        raise ValueError("start and end must be positive Excel row numbers")
    if start > end:
        raise ValueError("start cannot be greater than end")
    try:
        viewport = float(max_height)
    except (TypeError, ValueError) as exc:
        raise TypeError("max_height must be a positive number") from exc
    if not math.isfinite(viewport) or viewport <= 0:
        raise ValueError("max_height must be a finite positive number")

    get_height = _height_getter(row_heights, start, end)
    merges = _normalise_merged_spans(merged_spans, start, end)
    merge_at = {span.start: span for span in merges}

    current = start
    page_end: int | None = None
    used = 0.0
    epsilon = max(1e-9, viewport * 1e-12)

    while current <= end:
        # A previous page must never leave us inside a merge.  This check also
        # catches direct callers that ask to start halfway through a merge.
        containing = next(
            (span for span in merges if span.start < current <= span.end), None
        )
        if containing is not None:
            raise ValueError(
                f"目标起始行 {current} 位于合并单元格 {containing.start}:{containing.end} 内，"
                "无法截取完整单元格"
            )

        merge = merge_at.get(current)
        group_end = merge.end if merge is not None else current
        group_height = sum(get_height(row) for row in range(current, group_end + 1))

        if merge is not None and group_height > viewport + epsilon:
            raise ValueError(
                f"合并单元格 {merge.start}:{merge.end} 高度 {group_height:g} "
                f"超过截图高度 {viewport:g}，无法完整截图"
            )

        fits = page_end is None or used + group_height <= viewport + epsilon
        if not fits:
            # Keep the merge intact by ending the current page immediately
            # before it.  Hidden rows before the merge have already been
            # consumed by the loop and therefore cannot duplicate later.
            break

        page_end = group_end
        used += group_height
        current = group_end + 1

    # The first row is always consumed, including an oversized non-merged row.
    # Thus page_end is guaranteed to be set for a valid non-empty request.
    assert page_end is not None
    return RowSpan(start, page_end)


def plan_pages(
    row_heights: Mapping[int, object] | Sequence[object],
    start: int,
    end: int,
    max_height: float,
    merged_spans: Sequence[RowSpan] | None = None,
) -> Iterator[RowSpan]:
    """Yield contiguous, non-overlapping page spans covering ``start:end``."""

    next_start = start
    while next_start <= end:
        span = plan_page(row_heights, next_start, end, max_height, merged_spans)
        yield span
        if span.end < next_start:  # defensive guard against custom RowSpan types
            raise RuntimeError("pagination did not make progress")
        next_start = span.end + 1


_A1_RE = re.compile(r"^\$?([A-Za-z]{1,3})\$?([1-9][0-9]*)$")
_MAX_EXCEL_COLUMN = 16_384  # XFD
_MAX_EXCEL_ROW = 1_048_576


def column_name_to_index(name: str) -> int:
    """Convert an Excel column name (``A`` .. ``XFD``) to a 1-based index."""

    if not isinstance(name, str) or not re.fullmatch(r"[A-Za-z]{1,3}", name):
        raise ValueError(f"invalid Excel column name: {name!r}")
    index = 0
    for char in name.upper():
        index = index * 26 + (ord(char) - ord("A") + 1)
    if index > _MAX_EXCEL_COLUMN:
        raise ValueError(f"Excel column exceeds XFD: {name!r}")
    return index


def column_index_to_name(index: int) -> str:
    """Convert a 1-based Excel column index to its letters."""

    if not isinstance(index, int) or isinstance(index, bool):
        raise TypeError("column index must be an integer")
    if not 1 <= index <= _MAX_EXCEL_COLUMN:
        raise ValueError("Excel column index must be between 1 and 16384")
    result: list[str] = []
    value = index
    while value:
        value, remainder = divmod(value - 1, 26)
        result.append(chr(ord("A") + remainder))
    return "".join(reversed(result))


# Common aliases used by callers that prefer the shorter terminology.
column_to_index = column_name_to_index
index_to_column = column_index_to_name
column_to_number = column_name_to_index
number_to_column = column_index_to_name


def parse_a1_address(address: str) -> tuple[str, int]:
    """Parse a plain A1 cell address and return ``(column, row)``.

    Sheet prefixes, workbook references, formulas, ranges, and other external
    references are deliberately rejected.  Optional ``$`` markers are
    accepted and removed in the returned column name.
    """

    if not isinstance(address, str):
        raise TypeError("A1 address must be a string")
    match = _A1_RE.fullmatch(address.strip())
    if match is None:
        raise ValueError(
            f"invalid A1 address {address!r}; expected a single cell such as A1"
        )
    column, row_text = match.groups()
    column = column.upper()
    column_name_to_index(column)
    row = int(row_text)
    if row > _MAX_EXCEL_ROW:
        raise ValueError(f"Excel row exceeds 1048576: {row}")
    return column, row


def validate_a1_address(address: str) -> bool:
    """Return ``True`` for a valid local A1 cell address, otherwise ``False``."""

    try:
        parse_a1_address(address)
    except (TypeError, ValueError):
        return False
    return True


# Short aliases keep the helpers convenient for small integration scripts.
parse_a1 = parse_a1_address
is_valid_a1 = validate_a1_address


_INVALID_FILENAME_CHARS = re.compile(r'[<>:"/\\|?*\x00-\x1f]')
_RESERVED_WINDOWS_NAMES = {
    "CON",
    "PRN",
    "AUX",
    "NUL",
    *(f"COM{i}" for i in range(1, 10)),
    *(f"LPT{i}" for i in range(1, 10)),
}


def safe_filename(name: object, *, fallback: str = "sheet", max_length: int = 180) -> str:
    """Return a Windows-safe filename stem.

    Invalid device/path characters are replaced with underscores, surrounding
    spaces and dots are removed, and reserved device names receive a suffix.
    The result is a stem rather than a path; callers can append ``.png``.
    """

    if not isinstance(fallback, str) or not fallback:
        raise ValueError("fallback must be a non-empty string")
    if not isinstance(max_length, int) or max_length < 1:
        raise ValueError("max_length must be a positive integer")
    text = str(name) if name is not None else ""
    text = _INVALID_FILENAME_CHARS.sub("_", text).strip(" .")
    if not text:
        text = _INVALID_FILENAME_CHARS.sub("_", fallback).strip(" .") or "sheet"
    # Windows reserves the stem even when a caller supplied an extension,
    # e.g. ``CON.txt``.  Preserve that extension while making the stem safe.
    stem, separator, suffix = text.partition(".")
    if stem.upper() in _RESERVED_WINDOWS_NAMES:
        text = stem + "_file" + (separator + suffix if separator else "")
    # Avoid ending in a dot/space after truncation as Windows rejects it.
    text = text[:max_length].rstrip(" .")
    return text or "sheet"


__all__ = [
    "RowSpan",
    "plan_page",
    "plan_pages",
    "column_name_to_index",
    "column_index_to_name",
    "column_to_index",
    "index_to_column",
    "column_to_number",
    "number_to_column",
    "parse_a1_address",
    "validate_a1_address",
    "parse_a1",
    "is_valid_a1",
    "safe_filename",
]
