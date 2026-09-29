"""Entry point for BrowserScreenShot.exe."""
from __future__ import annotations

import argparse

from excelshot.screen_io import enable_dpi_awareness


def main() -> int:
    enable_dpi_awareness()
    parser = argparse.ArgumentParser(description="通过桌面截图读取浏览器远程 Excel")
    parser.add_argument("--self-test", action="store_true", help="检查屏幕截图模式依赖后退出")
    args = parser.parse_args()
    if args.self_test:
        from excelshot.ocr import OfflineOCR
        OfflineOCR()
        print("BrowserScreenShot self-test OK")
        return 0
    from excelshot.screen_gui import run_gui
    run_gui()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
