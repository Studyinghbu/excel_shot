"""Entry point for ExcelCellShot."""
from __future__ import annotations

import argparse

from excelshot.backend import CaptureOptions, capture


def main() -> int:
    parser = argparse.ArgumentParser(description="按完整单元格行分页截图 Excel 有效区域")
    parser.add_argument("path", nargs="?", help="Excel 文件路径；省略则打开图形界面")
    parser.add_argument("--sheet", default="", help="工作表名")
    parser.add_argument("--range", dest="address", default="", help="A1 或 A1:D200")
    parser.add_argument("--zoom", type=int, default=100)
    parser.add_argument("--delay", type=float, default=0.6)
    parser.add_argument("--wps", action="store_true", help="使用 WPS COM（默认 Microsoft Excel）")
    parser.add_argument("--self-test", action="store_true", help="检查内置 OCR 与打包资源后退出")
    args = parser.parse_args()
    if args.self_test:
        from excelshot.ocr import OfflineOCR
        OfflineOCR()
        print("ExcelCellShot self-test OK")
        return 0
    if not args.path:
        from excelshot.gui import run_gui
        return run_gui()
    try:
        capture(CaptureOptions(args.path, args.sheet, args.address,
                               "wps" if args.wps else "excel", args.zoom, args.delay))
        return 0
    except Exception as exc:
        print(f"失败：{exc}")
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
