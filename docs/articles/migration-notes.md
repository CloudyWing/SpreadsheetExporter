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

## v3.1 → v3.2

本節整理既有程式碼升級後可能觀察到的行為變更。新增的 API 請參考各主題文件。

### 自動篩選的預設範圍

工作表有多個 templates，且其中包含 `RecordSetTemplate<T>` 或 `DataTableTemplate` 時，`SetAutoFilterEnabled()` 的篩選範圍為第一個 `RecordSetTemplate<T>` 或 `DataTableTemplate` 的範圍（表頭到最後一筆資料），不再使用工作表已使用的範圍。只有單一 template 的工作表維持原行為。需要指定其他範圍時，使用 `SetAutoFilterRange(...)`，詳見 [SheetDefinition](sheet-definition.md)。

### 工作表名稱

- `CreateSheet` 比對重複名稱時不分大小寫。已有 `Sales` 時再建立 `SALES`，會得到 `SALES(1)`。
- `GetLayoutDiagnostics()` 與 `ValidateLayout()` 新增 `SE-LAYOUT-005`（名稱不符合 Excel 規則）與 `SE-LAYOUT-006`（名稱重複）。呼叫這兩個 API 的既有程式碼，若工作表名稱違反規則，升級後會收到診斷或 `SpreadsheetLayoutException`。

### 巢狀欄位路徑

`RecordSetTemplate<T>` 的欄位以巢狀屬性運算式或字串路徑取值時，路徑中途的物件為 `null`，欄位值為 `null`。先前會拋出 `ArgumentException`。`FromJson(...)` 建立的 `RecordSet` 不在此變更範圍，`FieldKey` 路徑中途為 `null` 時仍會拋出 `ArgumentException`。

### 清單資料驗證

- 清單項目含雙引號時會自動跳脫，先前會產生無效的公式。
- 清單來源（含前後雙引號）超過 255 個字元時，ClosedXML renderer 會拋出 `ArgumentOutOfRangeException`。

## 相關文件

- [入門指南](getting-started.md)
- [SpreadsheetDocument（文件）](spreadsheet-document.md)
- [SheetDefinition（工作表）](sheet-definition.md)
- [Templates（範本）](templates.md)
