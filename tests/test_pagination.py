from __future__ import annotations

import pytest

from excelshot.pagination import (
    RowSpan,
    column_index_to_name,
    column_name_to_index,
    parse_a1_address,
    plan_page,
    plan_pages,
    safe_filename,
    validate_a1_address,
)


def test_plan_page_greedy_variable_heights_and_hidden_rows() -> None:
    heights = {1: 10, 2: 0, 3: 15, 4: 20, 5: 0, 6: 10}
    assert plan_page(heights, 1, 6, 25) == RowSpan(1, 3)
    # Row 4 cannot fit, but the following hidden row is consumed with it only
    # on the next page; this keeps every output range contiguous.
    assert list(plan_pages(heights, 1, 6, 25)) == [
        RowSpan(1, 3),
        RowSpan(4, 5),
        RowSpan(6, 6),
    ]


def test_merged_span_moves_to_next_page_without_being_cut() -> None:
    heights = {1: 10, 2: 10, 3: 10, 4: 10, 5: 10}
    merged = [RowSpan(3, 4)]
    assert list(plan_pages(heights, 1, 5, 25, merged)) == [
        RowSpan(1, 2),
        RowSpan(3, 4),
        RowSpan(5, 5),
    ]


def test_merged_group_taller_than_viewport_is_clear_error() -> None:
    with pytest.raises(ValueError, match="超过截图高度"):
        plan_page({1: 30, 2: 30}, 1, 2, 50, [RowSpan(1, 2)])


def test_all_hidden_rows_are_consumed() -> None:
    heights = [0, 0, 0, 0]
    assert list(plan_pages(heights, 1, 4, 1)) == [RowSpan(1, 4)]


def test_oversized_non_merged_row_remains_complete() -> None:
    assert plan_page({1: 100, 2: 10}, 1, 2, 20) == RowSpan(1, 1)


def test_no_overlaps_and_full_coverage_with_mapping() -> None:
    heights = {row: 7 for row in range(1, 21)}
    spans = list(plan_pages(heights, 1, 20, 20))
    assert spans[0].start == 1
    assert spans[-1].end == 20
    assert all(left.end + 1 == right.start for left, right in zip(spans, spans[1:]))


def test_invalid_ranges_and_overlapping_merges() -> None:
    with pytest.raises(ValueError):
        plan_page([10], 2, 1, 20)
    with pytest.raises(ValueError, match="重叠"):
        plan_page({1: 5, 2: 5, 3: 5}, 1, 3, 20, [RowSpan(1, 2), RowSpan(2, 3)])


def test_column_helpers_and_a1_validation_reject_external_references() -> None:
    assert column_name_to_index("xFd") == 16_384
    assert column_index_to_name(16_384) == "XFD"
    assert parse_a1_address("$b$12") == ("B", 12)
    assert validate_a1_address("A1")
    assert not validate_a1_address("Sheet1!A1")
    assert not validate_a1_address("[Book.xlsx]Sheet1!A1")
    assert not validate_a1_address("=A1")
    assert not validate_a1_address("A1:B2")


def test_safe_filename_is_windows_safe_and_bounded() -> None:
    assert safe_filename('report:Q4\\2026?.png') == "report_Q4_2026_.png"
    assert safe_filename("CON") == "CON_file"
    assert safe_filename("...") == "sheet"
    assert len(safe_filename("x" * 500, max_length=20)) <= 20
