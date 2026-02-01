# SpreadsheetExporter

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](./LICENSE.md)

**SpreadsheetExporter** 為 [SpreadsheetExport](https://github.com/CloudyWing/SpreadsheetExport) 的 .NET Standard 版本。

本套件專注於 **建立匯出 Excel 所需的資料結構與邏輯**，並不直接依賴特定的 Excel 函式庫（如 NPOI 或 EPPlus）。而是透過 Facade 模式，將資料準備與 Excel 產生這兩個責任分離。

這意味著：

- **解耦**：您的業務邏輯與操作介面不需要綁死在 NPOI 或 EPPlus 上。
- **彈性**：未來若需更換 Excel 底層套件（例如從 NPOI 遷移至 EPPlus），只需更換 Exporter 實作，完全無需修改報表產生的邏輯。
- **維護性**：長期專案建議繼承 `ExporterBase` 實作自己的 Exporter 來處理特定邊界情況，而非直接依賴套件提供的預設實作。

## Projects

本解決方案包含以下專案：

- **[SpreadsheetExporter](./SpreadsheetExporter/)**
  核心專案，負責定義介面與建立匯出所需的資料結構（Templates、Schemas 等）。

- **[SpreadsheetExporter.Excel.NPOI](./SpreadsheetExporter.Excel.NPOI/)**
  實作專案，使用 **NPOI** 將 `SpreadsheetExporter` 定義的結構轉換為 Excel 檔案。

- **[SpreadsheetExporter.Excel.EPPlus](./SpreadsheetExporter.Excel.EPPlus/)**
  實作專案，使用 **EPPlus** 將 `SpreadsheetExporter` 定義的結構轉換為 Excel 檔案。

> [!WARNING]
> **關於 .NET 6+ 支援**
>
> 由於 NPOI 與 EPPlus 在升級至 .NET 6+ 後可能面臨 `System.Drawing.Common` 的圖形跨平台相容性問題，且本解決方案中的 `SpreadsheetExporter.Excel.NPOI` 與 `SpreadsheetExporter.Excel.EPPlus` 僅作為 **參考範本** 存在。
>
> 因此，這兩個實作套件 **不會提供支援 .NET 6 以上的版本**。建議您參考這些範本的程式碼，自行繼承 `ExporterBase` 實作符合您專案需求的 Exporter。

## Installation

透過 NuGet Package Manager 安裝核心與對應的實作套件：

### .NET CLI

```bash
dotnet add package CloudyWing.SpreadsheetExporter
# 選擇您需要的實作：
dotnet add package CloudyWing.SpreadsheetExporter.Excel.NPOI
# 或
dotnet add package CloudyWing.SpreadsheetExporter.Excel.EPPlus
```

## Documentation

完整文檔請參考：<https://cloudywing.github.io/SpreadsheetExporter/>

或查看以下主要文章：

- **[入門指南](./docs/articles/getting-started.md)**
- **[Exporters（匯出器）](./docs/articles/exporters.md)**
- **[Sheeters（工作表）](./docs/articles/sheeters.md)**
- **[Templates（範本）](./docs/articles/templates.md)**
- **[自訂樣式](./docs/articles/customization.md)**

## License

This project is MIT [licensed](./LICENSE.md).
