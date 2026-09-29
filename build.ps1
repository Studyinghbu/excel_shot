[CmdletBinding()]
param(
    [ValidateSet("all", "browser", "excel")]
    [string]$Target = "all"
)

$ErrorActionPreference = "Stop"
Set-Location -LiteralPath $PSScriptRoot

$root = (Get-Location).Path
$python = Join-Path $root ".venv\Scripts\python.exe"
$pyinstaller = Join-Path $root ".venv\Scripts\pyinstaller.exe"

if (-not (Test-Path -LiteralPath $python -PathType Leaf)) {
    $systemPython = (Get-Command python -ErrorAction Stop).Source
    & $systemPython -m venv (Join-Path $root ".venv")
    if ($LASTEXITCODE -ne 0) {
        throw "创建 .venv 失败。"
    }
}

& $python -m pip install -r (Join-Path $root "requirements.txt")
if ($LASTEXITCODE -ne 0) {
    throw "安装 requirements.txt 失败。"
}

if (-not (Test-Path -LiteralPath $pyinstaller -PathType Leaf)) {
    throw "找不到 PyInstaller：$pyinstaller"
}

function Build-Target {
    param(
        [Parameter(Mandatory = $true)][string]$Name,
        [Parameter(Mandatory = $true)][string]$SpecName
    )

    $specPath = Join-Path $root $SpecName
    if (-not (Test-Path -LiteralPath $specPath -PathType Leaf)) {
        throw "找不到打包配置：$specPath"
    }

    # Each target has its own work/dist directory. We intentionally do not
    # remove build or dist recursively, so unrelated artifacts remain intact.
    $workPath = Join-Path $root ("build\" + $Name)
    $distPath = Join-Path $root ("dist\" + $Name)
    New-Item -ItemType Directory -Force -Path $workPath, $distPath | Out-Null

    Write-Host "正在构建 $Name ..."
    & $pyinstaller --clean --noconfirm --workpath $workPath --distpath $distPath $specPath
    if ($LASTEXITCODE -ne 0) {
        throw "PyInstaller 构建 $Name 失败。"
    }

    $builtExe = Join-Path $distPath ($Name + ".exe")
    if (-not (Test-Path -LiteralPath $builtExe -PathType Leaf)) {
        throw "构建完成但未找到 EXE：$builtExe"
    }

    $releaseDir = Join-Path $root ("release\" + $Name)
    New-Item -ItemType Directory -Force -Path $releaseDir | Out-Null
    Copy-Item -LiteralPath $builtExe -Destination (Join-Path $releaseDir ($Name + ".exe")) -Force
    Copy-Item -LiteralPath (Join-Path $root "README.md") -Destination (Join-Path $releaseDir "README.md") -Force

    if ($Name -eq "BrowserScreenShot") {
        # Keep a convenient copy next to the source tree as well as in the
        # distributable folder. Runtime output (shot) is created beside it.
        Copy-Item -LiteralPath $builtExe -Destination (Join-Path $root "BrowserScreenShot.exe") -Force

        $zipPath = Join-Path $root "release\BrowserScreenShot_Win10.zip"
        # This is one known file, never a recursive directory deletion.
        if (Test-Path -LiteralPath $zipPath -PathType Leaf) {
            Remove-Item -LiteralPath $zipPath -Force
        }
        Compress-Archive -Path (Join-Path $releaseDir "*") -DestinationPath $zipPath -CompressionLevel Optimal
        if (-not (Test-Path -LiteralPath $zipPath -PathType Leaf)) {
            throw "未生成 ZIP：$zipPath"
        }
    }
    elseif ($Name -eq "ExcelCellShot") {
        Copy-Item -LiteralPath $builtExe -Destination (Join-Path $root "ExcelCellShot.exe") -Force
    }

    Write-Host "完成：$releaseDir"
}

switch ($Target) {
    "all" {
        Build-Target -Name "ExcelCellShot" -SpecName "ExcelCellShot.spec"
        Build-Target -Name "BrowserScreenShot" -SpecName "BrowserScreenShot.spec"
    }
    "browser" {
        Build-Target -Name "BrowserScreenShot" -SpecName "BrowserScreenShot.spec"
    }
    "excel" {
        Build-Target -Name "ExcelCellShot" -SpecName "ExcelCellShot.spec"
    }
}
