# ExcelCellShot

这是一个 Windows 10/11 64 位独立程序：选择 Excel 文件后，它连接 Microsoft Excel，读取当前工作表的有效数据区，按完整的 worksheet 行分页并滚动到每一页，再把**单元格区域本身**导出为 PNG。上下页使用不重叠的行号，隐藏行也只会被消费一次；合并行不会被切开。每次运行会在 `ExcelCellShot.exe` 同级建立 `shot` 文件夹。

程序内置 RapidOCR 的中英文 ONNX 模型，OCR 在本机离线运行，只用于检查相邻截图的文字衔接和记录到 manifest，不会把表格上传网络，也不会因为相同文字出现在不同单元格而删除合法截图。PNG 文件名包含工作簿、工作表、页码和起止行号；同一页完全相同的导出图会跳过并记录在 manifest 中。

## 使用

1. 先关闭该工作簿的编辑对话框并确保 Excel 可以正常显示；打开 `ExcelCellShot.exe`。
2. 选择 `.xlsx`、`.xls`、`.xlsm` 或 `.xlsb`。工作表名和 A1 范围留空时，使用当前工作表的 `UsedRange`。
3. 点击“开始截图”。截图期间 Excel 保持可见，程序窗口最小化以免遮挡；完成后图片在 `shot`。

程序也支持命令行：

```powershell
ExcelCellShot.exe "D:\data\表格.xlsx" --sheet "Sheet1" --range A1:H200
```

## 打包

在本目录运行 `build.ps1`。它会安装锁定版本依赖，并用 PyInstaller 将 OCR 模型、ONNX Runtime 和 COM 依赖全部放进 `release\ExcelCellShot`。目标机器需要安装 Microsoft Excel（Office 32 位或 64 位均可，但 Python 打包程序应使用相同位数的系统 COM 注册）。
