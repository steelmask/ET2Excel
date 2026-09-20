# ET 转 Excel 小工具

图形界面工具：选择 WPS 表格 `.et` 文件，转换并另存为 `.xlsx` 或 `.xls`。

## 使用前准备

推荐安装 **WPS Office（含 WPS 表格）**，它会通过 WPS 的自动化接口转换，兼容性最好。随后在本目录执行：

```powershell
py -m pip install -r requirements.txt
py et_converter.py
```

若电脑没有 WPS，工具会自动尝试使用 LibreOffice；请先安装 LibreOffice，并让 `soffice.exe` 可被系统找到。

## 使用方法

1. 点击“选择文件”，选取 `.et` 文件。
2. 选择 `.xlsx` 或 `.xls`。
3. 点击“转换并保存”，指定输出位置。

`.xls` 是旧版 Excel 格式，有工作表行列数和格式能力限制；没有兼容旧系统的需求时，建议选择 `.xlsx`。
