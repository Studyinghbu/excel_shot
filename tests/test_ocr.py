from excelshot.ocr import OfflineOCR


def test_compare_adjacent_only_flags_boundary_text():
    ocr = OfflineOCR.__new__(OfflineOCR)
    previous = {"items": [{"text": "标题"}, {"text": "最后一行"}]}
    current = {"items": [{"text": "最后一行"}, {"text": "下一行"}]}
    result = ocr.compare_adjacent(previous, current)
    assert result["duplicate"] is True
    assert "相邻页" in result["warning"]


def test_compare_adjacent_does_not_flag_disjoint_text():
    ocr = OfflineOCR.__new__(OfflineOCR)
    result = ocr.compare_adjacent(
        {"items": [{"text": "A"}]}, {"items": [{"text": "B"}]} 
    )
    assert result == {"duplicate": False, "warning": ""}

