$ErrorActionPreference = 'Stop'
Set-Location $PSScriptRoot
if (-not (Test-Path .venv\Scripts\python.exe)) { python -m venv .venv }
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
Remove-Item -Recurse -Force build,dist -ErrorAction SilentlyContinue
.\.venv\Scripts\pyinstaller.exe --clean --noconfirm ExcelCellShot.spec
New-Item -ItemType Directory -Force release | Out-Null
Remove-Item -Recurse -Force release\ExcelCellShot -ErrorAction SilentlyContinue
New-Item -ItemType Directory -Force release\ExcelCellShot | Out-Null
Copy-Item dist\ExcelCellShot.exe release\ExcelCellShot\ExcelCellShot.exe
Copy-Item README.md release\ExcelCellShot\README.md
Write-Host "完成：release\ExcelCellShot\ExcelCellShot.exe"
