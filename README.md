# SpreadsheetExporter

**SpreadsheetExporter** 為 [SpreadsheetExport](https://github.com/CloudyWing/SpreadsheetExport) 的 .NET Standard 版本。

套件本身並無 Excel 匯出功能，而是針對 Excel 的匯出的部分，提供一個 Facade 專案，在使用上仍需要搭配 NPOI 或 EPPlus 之類的 Excel 套件使用。

此套件目的在於如果未來有需要更換 Excel 套件，只需要在繼承 `ExporterBase` 的 Exporter 來裡變更 Excel 套件即可，而不會動到操作介面的部分，因此目前雖然有提供一個具體實作的 `NPOI.Exporter`，但長期性的專案仍不建議直接使用，而是另寫一個 Exporter 來繼承它以處理邊界問題。

## Projects
 * [SpreadsheetExporter](./SpreadsheetExporter/)：建立匯出 Excel 所需要的資訊。
 * [SpreadsheetExporter.Excel.NPOI](./SpreadsheetExporter.Excel.NPOI/)： 是負責將 SpreadsheetExporter 的介面所建立的資訊，透過 NPOI 來匯出 Excel。
 * [SpreadsheetExporter.Excel.EPPlus](./SpreadsheetExporter.Excel.EPPlus/)： 是負責將 SpreadsheetExporter 的介面所建立的資訊，透過 EPPlus 來匯出 Excel。

## Templates
目前提供兩種基本的Templates分別如下：
* GridTemplate：使用上類似 HTML 的 table，提供 `CreateRow()`(可理解成 HTML tr Tag) 和 `CreateCell()` (可理解成 HTML td Tag)兩個 API，藉由在 `CreateCell()` 傳入 `rowSpan` 和 `columnSpan` 的參數來產生各儲存格的座標與合併資訊。
* ListTemplate：藉由設定 `DataSource`，以及設定各欄位的名稱以及對應 `DataSource` 的 Property Name 來產生表格格式的儲存格的座標與合併資訊。

由於此套件的重點在於給 Excel 套件的匯出功能做出邊界，在使用上未必比使用原套件方便，但如果你專案匯出的 Excel 格式相當固定，可自行定義 Template，內部利用 `GridTemplate` 和 `ListTemplate`來處理。


## License
This project is MIT [licensed](./LICENSE.md).