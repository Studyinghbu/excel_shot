# ExcelCellShot / BrowserScreenShot

这个项目有两个 Windows 10/11 x64 程序：

- `ExcelCellShot.exe` 适合本机安装了 Microsoft Excel 的情况。它通过 Excel 接口读取工作表并按行导出单元格区域。
- `BrowserScreenShot.exe` 适合 Excel 只显示在浏览器、远程桌面或 KVM 画面中的情况。它不需要访问工作簿文件，也不需要本机安装 Excel，只读取屏幕像素。

浏览器模式不会访问网页 DOM、浏览器地址、剪贴板或网络接口。程序只会在你选定的矩形内截图、用本机 OCR 识别文字，并据此判断相邻画面的对齐位置；浏览器工具栏、远程桌面工具栏和任务栏不会被写入输出图。

## 浏览器模式使用方法

1. 运行 `BrowserScreenShot.exe`，再打开浏览器中的远程 KVM 画面，让表格稳定显示。
2. 按程序提示用鼠标框出**会滚动的单元格区域**。请从第一列到最后一列选择数据单元格，不要把冻结窗格、固定表头、行号列、工作表标签或滚动条包含进来。冻结表头不会随滚轮移动，放进选择区会影响页面对齐。
3. 确认首屏后开始采集。程序会截取当前区域，使用 OCR 读取文字坐标，再在同一位置向下滚动一个屏幕；下一张会用前后画面的重叠内容做保守对齐，尽量让上下页连接。
4. 采集到表格底部时按 `Esc` 停止。界面也提供最大滚动步数，用于防止远程页面没有底部标记时持续滚动；达到上限会自动结束。

每次运行会在 EXE 同级建立 `shot` 文件夹。里面保留单元格区域的逐页 PNG 和 `manifest.json`，包括 OCR 摘要、滚动步数、对齐结果和停止原因，方便检查远程画面或重新处理。截图文件名包含时间、页码和滚动序号；相同文字出现在不同单元格时不会因此删除截图。默认结果是逐页图片；是否提供拼接选项以当前程序界面为准，拼接不是采集成功的必要条件。

浏览器表格的行高可能随内容变化。程序按屏幕像素和 OCR 坐标测量重叠，再把已确认的重复区域裁掉，只保存首屏和每次滚动后新增的区域；`stitched.png` 会把这些区域顺序拼接。某一张图的边缘仍可能落在一行中间。它不能从远程画面可靠推断“有效区域”或真正的表格末尾；对于无限滚动、虚拟化表格、合并单元格或行高不断变化的页面，请用最大步数和首尾截图检查结果。

## 输出和隐私

OCR 模型和运行库已经打包在 EXE 中，Win10/Win11 x64 可离线运行。OCR 只在本机使用，不会上传截图。输出区域只包含你选择的单元格，不会主动保存浏览器窗口之外的画面；如果选择框本身包含敏感信息，它仍会出现在 PNG 中，请在分享 `shot` 前检查内容。

## 直接运行源码

在项目目录执行：

```powershell
.\.venv\Scripts\python.exe browser_capture.py
```

浏览器模式不依赖本机 Excel。旧的 Excel 文件模式仍可使用：

```powershell
.\.venv\Scripts\python.exe main.py "D:\data\表格.xlsx" --sheet "Sheet1" --range A1:H200
```

## 下载和打包

已打包文件位于：

```text
release\BrowserScreenShot\BrowserScreenShot.exe
release\BrowserScreenShot_Win10.zip
release\ExcelCellShot\ExcelCellShot.exe
```

解压 `BrowserScreenShot_Win10.zip` 后，直接运行其中的 `BrowserScreenShot.exe`。程序会在该 EXE 的同级创建 `shot`；不需要安装 Python、Excel 或 OCR 模型。首次运行时请给 Windows 桌面捕获和输入控制权限，并保持远程 KVM 页面可见。

需要从源码重新打包时，在项目目录运行：

```powershell
.\build.ps1                 # 同时构建两个程序
.\build.ps1 -Target browser # 只构建 BrowserScreenShot
.\build.ps1 -Target excel   # 只构建 ExcelCellShot
```

构建脚本使用 `requirements.txt` 中锁定的版本，并把 RapidOCR 的 ONNX 模型一并放入 EXE。构建过程中不会删除整个 `build` 或 `dist` 目录；每个目标使用自己的临时子目录，已有的其他构建产物会保留。
