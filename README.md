# SpreadsheetExporter

**SpreadsheetExporter** 為 [SpreadsheetExport](https://github.com/CloudyWing/SpreadsheetExport) 的 .NET Standard 版本。

套件本身並無 Excel 匯出功能，而是針對 Excel 的匯出的部分，提供一個 Facade 專案，在使用上仍需要搭配 NPOI 或 EPPlus 之類的 Excel 套件使用。

此套件目的在於如果未來有需要更換 Excel 套件，只需要在繼承 `ExporterBase` 的 Exporter 來裡變更 Excel 套件即可，而不會動到操作介面的部分，因此目前雖然有提供一個具體實作的 `NPOI.Exporter`，但長期性的專案仍不建議直接使用，而是另寫一個 Exporter 來繼承它以處理邊界問題。

## Projects
 * [SpreadsheetExporter](./SpreadsheetExporter/)：建立匯出 Excel 所需要的資訊。
 * [SpreadsheetExporter.Excel.NPOI](./SpreadsheetExporter.Excel.NPOI/)： 是負責將 SpreadsheetExporter 的介面所建立的資訊，透過 NPOI 來匯出 Excel。
 * [SpreadsheetExporter.Excel.EPPlus](./SpreadsheetExporter.Excel.EPPlus/)： 是負責將 SpreadsheetExporter 的介面所建立的資訊，透過 EPPlus 來匯出 Excel。

## 參考文件
* [入門教學](./docs/Guide/%E5%85%A5%E9%96%80%E6%95%99%E5%AD%B8.md)。
* [API 文件](./docs/API/index.md)。

## License
This project is MIT [licensed](./LICENSE.md).