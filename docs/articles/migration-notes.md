# Migration Notes

本文整理 `SpreadsheetExporter` 各版本之間的升級重點，方便後續持續追加新的遷移內容。

## v2 → v3

### 核心命名調整

- `Sheeter` 改為 `SheetDefinition`
- `SheeterContext` 改為 `SheetLayout`
- `TemplateContext` 改為 `TemplateLayout`
- `GetContext()` 改為 `GetLayout()`
- `ITemplate` 改為 `ISheetTemplate`
- `ISpreadsheetExporter` / `ExporterBase` 改為 `ISpreadsheetRenderer`

### Renderer 與套件調整

v3 將 renderer 抽離成 `ISpreadsheetRenderer`，並以 `SpreadsheetExporter.Renderer.ClosedXML` 作為目前的 `.xlsx` 輸出實作。

- 舊的 EPPlus / NPOI solution 專案已移除
- 請改用 `CloudyWing.SpreadsheetExporter.Renderer.ClosedXML`
- 使用入口改為 `ExcelRenderer`

若是非 DI 專案，可透過 `SpreadsheetManager.SetRenderer(...)` 完成初始化：

```csharp
using CloudyWing.SpreadsheetExporter;
using CloudyWing.SpreadsheetExporter.Renderer.ClosedXML;

SpreadsheetManager.SetRenderer(static () => new ExcelRenderer());
SpreadsheetDocument document = SpreadsheetManager.CreateDocument();
```

若是 DI 專案，則可直接註冊 `ISpreadsheetRenderer`，並在使用處建立 `SpreadsheetDocument`：

```csharp
using CloudyWing.SpreadsheetExporter;
using CloudyWing.SpreadsheetExporter.Renderer.ClosedXML;
using Microsoft.Extensions.DependencyInjection;

ServiceCollection services = new();
services.AddSingleton<ISpreadsheetRenderer>(_ => new ExcelRenderer());

ServiceProvider serviceProvider = services.BuildServiceProvider();
ISpreadsheetRenderer renderer = serviceProvider.GetRequiredService<ISpreadsheetRenderer>();

SpreadsheetDocument document = new(renderer);
```

### 工作表層級設定調整

`FreezePanes` 與 `IsAutoFilterEnabled` 在 v3 屬於 `SheetDefinition` 的設定，不再掛在 `RecordSetTemplate`。

```csharp
SheetDefinition sheet = document.CreateSheet("Orders");
sheet
    .SetFreezePanes(0, 1)
    .SetAutoFilterEnabled();
```

### RecordSet 階層欄位調整

`RecordSetColumnCollection<T>` 移除了 `AddChildToLastColumn(...)` 系列多載，改為先取得最後一個欄位，再對其 `ChildColumns` 操作。

```csharp
template.Columns
    .Add("Order Info")
    .GetLastColumn().ChildColumns
        .Add("Order ID", static x => x.OrderId)
        .Add("Customer", static x => x.Customer);
```

### Cell Generator 型別調整

`Cell` 上的 generator 屬性型別已從 `Func<int, int, T>` 改為 `CellGenerator<T>`。

多數直接寫 lambda 的情境通常不需要改寫，但若有明確宣告舊 delegate 型別，請一併更新。

## 相關文件

- [入門指南](getting-started.md)
- [SpreadsheetDocument（文件）](spreadsheet-document.md)
- [SheetDefinition（工作表）](sheet-definition.md)
- [Templates（範本）](templates.md)
