# 入門指南

本指南協助您在專案中設定 SpreadsheetExporter，並建立第一個 Excel 匯出功能。

## 安裝

安裝核心函式庫：

```powershell
dotnet add package CloudyWing.SpreadsheetExporter
```

安裝 Excel 實作套件（ClosedXML）：

```powershell
dotnet add package CloudyWing.SpreadsheetExporter.Renderer.ClosedXML
```

## 設定

SpreadsheetExporter 目前有兩種啟動方式：

- 非 DI 專案：使用 `SpreadsheetManager.SetRenderer(...)`，適合 Console、腳本、單機工具與快速驗證。
- 已使用 DI 的專案：在組合根註冊 `ISpreadsheetRenderer`，再由應用程式流程建立 `SpreadsheetDocument`。

### 非 DI 專案

在應用程式啟動時（例如 `Program.cs`）登錄 renderer：

```csharp
using CloudyWing.SpreadsheetExporter;
using CloudyWing.SpreadsheetExporter.Renderer.ClosedXML;

SpreadsheetManager.SetRenderer(static () => new ExcelRenderer());
```

> [!IMPORTANT]
> 請務必透過 `SpreadsheetManager.CreateDocument()` 建立文件實例，而非直接使用 `new SpreadsheetDocument(new ExcelRenderer())`。統一透過 `SpreadsheetManager` 管理，可確保全域設定（如 `DefaultCellStyles`）正確套用。

### DI 專案

若你的應用程式本來就使用 `Microsoft.Extensions.DependencyInjection`，可以直接在組合根註冊 renderer：

```csharp
using CloudyWing.SpreadsheetExporter;
using CloudyWing.SpreadsheetExporter.Renderer.ClosedXML;
using Microsoft.Extensions.DependencyInjection;

ServiceCollection services = new();
services.AddSingleton<ISpreadsheetRenderer>(_ => new ExcelRenderer());

ServiceProvider serviceProvider = services.BuildServiceProvider();
ISpreadsheetRenderer renderer = serviceProvider.GetRequiredService<ISpreadsheetRenderer>();

SpreadsheetDocument document = new SpreadsheetDocument(renderer);
```

若需要初始化預設值，也可以在組合根先設定：

```csharp
SpreadsheetManager.DefaultCellStyles = new CellStyleConfiguration {
    HeaderStyle = SpreadsheetManager.DefaultCellStyles.HeaderStyle with {
        HasBorder = true
    }
};
```

### 自訂 renderer 的生命週期選擇

DI 專案中，`ISpreadsheetRenderer` 的生命週期取決於自訂 renderer 是否保存狀態，以及是否依賴其他 scoped 服務。

- `Singleton`：適合無狀態 renderer。renderer 不保存跨呼叫狀態，也不依賴 scoped 服務時，可優先使用。
- `Scoped`：適合 renderer 需要使用目前 request、租戶、語系或其他 scoped 資訊。
- `Transient`：適合 renderer 內部會累積暫存狀態，或包裝的物件不適合重用。

例如下列情境應考慮 `Scoped` 或 `Transient`，不要直接使用 `Singleton`：

- renderer 依賴 scoped 的 tenant provider、使用者資訊或資料存取服務。
- renderer 內部持有可變欄位，且每次匯出都會改寫。
- renderer 包裝的第三方物件不保證 thread-safe。

內建 `ExcelRenderer` 本身偏向無狀態，通常可使用 `Singleton`。若自訂 renderer 的 thread-safe 或重用行為不明確，建議先從 `Transient` 或 `Scoped` 開始，再依實際需求調整。

## 第一個匯出範例

```csharp
// 建立文件
SpreadsheetDocument document = SpreadsheetManager.CreateDocument();

// 建立工作表
SheetDefinition sheet = document.CreateSheet("工作表1");

// 建立 GridTemplate 並加入內容
GridTemplate template = new GridTemplate();
template.CreateRow();
template.CreateCell("Hello, SpreadsheetExporter!");
sheet.AddTemplate(template);

// 匯出至檔案
document.ExportFile($@"C:\Sample{document.FileNameExtension}");
```

## 下一步

- **[SpreadsheetDocument](spreadsheet-document.md)** - 了解活頁簿層級的設定與匯出
- **[SheetDefinition](sheet-definition.md)** - 理解工作表管理
- **[Templates](templates.md)** - 探索不同的資料結構化方式
- **[自訂樣式](customization.md)** - 自訂儲存格樣式與格式
