# -*- mode: python ; coding: utf-8 -*-
from PyInstaller.utils.hooks import collect_data_files, collect_submodules

datas = collect_data_files('rapidocr_onnxruntime', include_py_files=False)
hiddenimports = collect_submodules('rapidocr_onnxruntime') + [
    'win32com', 'win32com.client', 'pythoncom', 'pywintypes', 'win32timezone'
]

a = Analysis(
    ['main.py'], pathex=['.'], binaries=[], datas=datas,
    hiddenimports=hiddenimports, hookspath=[], hooksconfig={}, runtime_hooks=[],
    excludes=['torch', 'tensorflow', 'paddle', 'matplotlib'], noarchive=False,
)
pyz = PYZ(a.pure)
exe = EXE(
    pyz, a.scripts, a.binaries, a.datas, [], name='ExcelCellShot',
    debug=False, bootloader_ignore_signals=False, strip=False, upx=False,
    console=False,
)
