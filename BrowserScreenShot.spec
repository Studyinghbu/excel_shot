# -*- mode: python ; coding: utf-8 -*-
"""PyInstaller configuration for the browser/remote-KVM screenshot tool."""

from PyInstaller.utils.hooks import collect_data_files, collect_submodules

# RapidOCR ships its ONNX models as package data. Keep them in the one-file
# executable just as in ExcelCellShot.spec.
datas = collect_data_files("rapidocr_onnxruntime", include_py_files=False)
hiddenimports = collect_submodules("rapidocr_onnxruntime") + [
    "onnxruntime",
    "onnxruntime.capi._pybind_state",
    "PIL.ImageGrab",
]

a = Analysis(
    ["browser_capture.py"],
    pathex=["."],
    binaries=[],
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=["torch", "tensorflow", "paddle", "matplotlib"],
    noarchive=False,
)

pyz = PYZ(a.pure)
exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name="BrowserScreenShot",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    console=False,
)
