# SpreadsheetExporter.Excel.NPOI
**NPOI** 的 Facade 專案，將 **CloudyWing.SpreadsheetExporter** API 產生的 Spreadsheet Info，透過 **NPOI** 匯出成 Excel。
建立 ExcelExporter 時，可用 `ExcelFormat` 控制產出格式，設定值如下：
* ExcelBinaryFileFormat：副檔名為 `xls`。
* OfficeOpenXmlDocument：副檔名為 `xlsx`。

## 相依套件
| 套件名稱 | 最低支援版本 |
|-|-|
| CloudyWing.SpreadsheetExporter | v1.0.0 |
| NOPI | v2.5.4 |

## 尚未支援支援功能
部分功能因為 NPOI 未實作，或是我不知道如何利用 NPOI 實作，所以並未支援，此類功能預設會 `throw new NotImplementedException()`，如果希望遇到未實作功能是忽略，而非拋出例外，請將 `IsClosedNotImplementedException` 設為 `true`。

目前尚未支援功能如下：
* ExcelBinaryFileFormat 格式(xls)：
    * 浮水印。
* OfficeOpenXmlDocument 格式(xlsx)：
    * 保護活頁簿。

## 參考文件
* [API 文件](../../docs/API/Excel/NPOI/CloudyWing.SpreadsheetExporter.Excel.NPOI.md)
