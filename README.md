# SpreadsheetExporter

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](./LICENSE.md)

此專案只是我為懶得計算 Excel 座標和處理合併儲存格開發出來的套件，只有在我比較常使用 Excel 那段時間會進行維護，且可能一加功能就是破壞性變更的大改版，不建議其他人使用。

## 專案內容

`SpreadsheetExporter` 採用了「**先描述活頁簿，再交給 Renderer 輸出**」的架構。核心專案負責建立 `SpreadsheetDocument`、`SheetDefinition` 與各種模板；實際 `.xlsx` 產出則由 `ClosedXML` Renderer 負責。

## Solution 內容

- **[`src/SpreadsheetExporter`](./src/SpreadsheetExporter/)**
  核心模型與模板：`SpreadsheetDocument`、`SheetDefinition`、`GridTemplate`、`RecordSetTemplate<T>` 等。
- **[`src/SpreadsheetExporter.Renderer.ClosedXML`](./src/SpreadsheetExporter.Renderer.ClosedXML/)**
  `ISpreadsheetRenderer` 的 ClosedXML 實作，輸出 `.xlsx` 檔案。
- **[`samples/SpreadsheetExporter.Sample`](./samples/SpreadsheetExporter.Sample/)**
  展示 v3 Fluent API 與 `FromJson(...)` 的基本使用方式。
- **[`benchmarks/SpreadsheetExporter.Benchmarks`](./benchmarks/SpreadsheetExporter.Benchmarks/)**
  BenchmarkDotNet 基準，評估模板建立與 ClosedXML 匯出成本。

## Installation

### .NET CLI

```bash
dotnet add package CloudyWing.SpreadsheetExporter
dotnet add package CloudyWing.SpreadsheetExporter.Renderer.ClosedXML
```

## Quick Start

若是 Console、腳本或沒有導入 DI 的專案，建議直接用 `SpreadsheetManager.SetRenderer(...)` 作為最短路徑：

```csharp
using CloudyWing.SpreadsheetExporter;
using CloudyWing.SpreadsheetExporter.Renderer.ClosedXML;
using CloudyWing.SpreadsheetExporter.Templates;
using CloudyWing.SpreadsheetExporter.Templates.RecordSet;

SpreadsheetManager.SetRenderer(static () => new ExcelRenderer());

SpreadsheetDocument document = SpreadsheetManager.CreateDocument();
SheetDefinition sheet = document.CreateSheet("Orders", defaultRowHeight: 20);

RecordSetTemplate<Order> template = new(GetOrders()) {
    HeaderHeight = 22,
    RecordHeight = 20
};

template.Columns
    .Add("Order ID", order => order.OrderId)
    .Add("Customer", order => order.Customer)
    .Add(
        "Amount",
        order => order.Amount,
        fieldStyleGenerator: _ => SpreadsheetManager.DefaultCellStyles.FieldStyle with {
            HorizontalAlignment = HorizontalAlignment.Right,
            DataFormat = "#,##0.00"
        }
    );

sheet.AddTemplate(template);
sheet
    .SetFreezePanes(0, 1)
    .SetAutoFilterEnabled();

document.ExportFile("orders.xlsx");
```

若你偏好以設定檔描述報表，也可以使用：

```csharp
SpreadsheetManager.SetRenderer(static () => new ExcelRenderer());
SpreadsheetDocument document = SpreadsheetDocument.FromJson(json);
byte[] bytes = document.Export();
```

若你的專案已經使用 DI，也可以在組合根先完成初始化，再建立文件：

```csharp
using CloudyWing.SpreadsheetExporter;
using CloudyWing.SpreadsheetExporter.Renderer.ClosedXML;
using Microsoft.Extensions.DependencyInjection;

ServiceCollection services = new();
services.AddSingleton<ISpreadsheetRenderer>(_ => new ExcelRenderer());
services.AddSingleton(_ => {
    SpreadsheetManager.DefaultCellStyles = new CellStyleConfiguration {
        FieldStyle = SpreadsheetManager.DefaultCellStyles.FieldStyle with {
            DataFormat = "#,##0.00"
        }
    };

    return 0;
});

ServiceProvider serviceProvider = services.BuildServiceProvider();
ISpreadsheetRenderer renderer = serviceProvider.GetRequiredService<ISpreadsheetRenderer>();

SpreadsheetDocument document = new(renderer);
```

## Sample 與 Benchmark

### 執行 Sample

```bash
dotnet run --project .\samples\SpreadsheetExporter.Sample
```

Sample 會在執行輸出目錄下的 `artifacts` 目錄建立：

- `sales-report.xlsx`
- `sales-report-json.xlsx`

### 執行 Benchmarks

```bash
dotnet run -c Release --project .\benchmarks\SpreadsheetExporter.Benchmarks -- --filter *
```

目前 benchmark 聚焦於兩類工作負載：

- `TemplateBenchmarks`：量測模板建立與 JSON 解析成本。
- `ExporterBenchmarks`：量測 ClosedXML 匯出基本版與含樣式版活頁簿的成本。

## Documentation

完整文件請參考：<https://cloudywing.github.io/SpreadsheetExporter/>

主要文件入口：

- [入門指南](./docs/articles/getting-started.md)
- [Migration Notes](./docs/articles/migration-notes.md)
- [SpreadsheetDocument](./docs/articles/spreadsheet-document.md)
- [SheetDefinition](./docs/articles/sheet-definition.md)
- [Templates](./docs/articles/templates.md)
- [自訂樣式](./docs/articles/customization.md)

## License

This project is MIT [licensed](./LICENSE.md).
