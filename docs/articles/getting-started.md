# 入門指南

本指南將協助您在專案中設定 SpreadsheetExporter，並建立第一個 Excel 匯出功能。

## 安裝

### 安裝 NuGet 套件

1. 安裝核心函式庫：

   ```powershell
   dotnet add package CloudyWing.SpreadsheetExporter
   ```

2. 選擇並安裝 Excel 實作套件：

   **選項 1：NPOI**（Apache POI for .NET）

   ```powershell
   dotnet add package CloudyWing.SpreadsheetExporter.Excel.NPOI
   dotnet add package NPOI
   ```

   **選項 2：EPPlus**

   ```powershell
   dotnet add package CloudyWing.SpreadsheetExporter.Excel.EPPlus
   dotnet add package EPPlus
   ```

> [!TIP]
> NPOI 和 EPPlus 皆受支援，請依據專案需求和授權考量選擇合適的套件。

## 設定

在應用程式啟動時（例如 `Program.cs`）加入以下設定：

```csharp
using CloudyWing.SpreadsheetExporter;
using CloudyWing.SpreadsheetExporter.Excel.NPOI;

SpreadsheetManager.SetExporter(() => new ExcelExporter());
```

> [!IMPORTANT]
> 請務必使用 `SpreadsheetManager.CreateExporter()` 建立 Exporter 實例，而非直接使用 `new ExcelExporter()`。這樣可以輕鬆切換 Excel 實作套件，並套用全域設定。

## 第一個匯出範例

以下是建立 Excel 檔案的簡單範例：

```csharp
// 建立 Exporter 實例
ExporterBase exporter = SpreadsheetManager.CreateExporter();

// 建立 Sheeter（代表一個工作表）
Sheeter sheeter = exporter.CreateSheeter();

// 建立 Template 並加入內容
GridTemplate template = new GridTemplate();
template.CreateRow();
template.CreateCell("Hello, SpreadsheetExporter!");

// 將 Template 加入 Sheeter
sheeter.AddTemplate(template);

// 匯出為檔案
exporter.ExportFile($@"C:\Sample{exporter.FileNameExtension}");
```

## 下一步

現在您已經完成 SpreadsheetExporter 的設定，接下來可以深入了解核心概念：

- **[Exporters](exporters.md)** - 了解活頁簿層級的設定
- **[Sheeters](sheeters.md)** - 理解工作表管理
- **[Templates](templates.md)** - 探索不同的資料結構化方式
- **[Customization](customization.md)** - 自訂儲存格樣式與格式
