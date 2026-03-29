# SpreadsheetDocument

`SpreadsheetDocument` 代表一份試算表文件，負責管理工作表定義並協調最終渲染輸出。

## 建立文件

若是非 DI 專案，建議使用 `SpreadsheetManager` 作為啟動入口：

透過 `SpreadsheetManager.CreateDocument()` 建立文件：

```csharp
SpreadsheetManager.SetRenderer(static () => new ExcelRenderer());

SpreadsheetDocument document = SpreadsheetManager.CreateDocument();
```

若你的專案已使用 DI，也可直接傳入 renderer 實例建立：

```csharp
SpreadsheetDocument document = new SpreadsheetDocument(new ExcelRenderer());
```

例如在 DI 組合根解析 renderer 後建立：

```csharp
ISpreadsheetRenderer renderer = serviceProvider.GetRequiredService<ISpreadsheetRenderer>();
SpreadsheetDocument document = new SpreadsheetDocument(renderer);
```

## 文件層級設定

| 屬性 | 說明 |
| --- | --- |
| `DefaultFont` | 套用至整份文件的預設字型；`null` 代表使用 renderer 預設值 |
| `DefaultSheetName` | 自動產生工作表名稱時使用的前綴（預設值：「工作表」） |
| `ContentType` | 匯出格式的 MIME 類型，來自 renderer |
| `FileNameExtension` | 匯出格式的副檔名（含前置點），來自 renderer |

```csharp
SpreadsheetDocument document = SpreadsheetManager.CreateDocument();
document.DefaultFont = new CellFont("Calibri", 11);
document.DefaultSheetName = "Sheet";
```

## 工作表管理

### 建立工作表

```csharp
// 建立指定名稱的工作表
SheetDefinition sheet = document.CreateSheet("銷售報表");

// 建立指定名稱與預設列高的工作表
SheetDefinition sheet = document.CreateSheet("銷售報表", defaultRowHeight: 20);

// 自動命名（依序產生「工作表1」、「工作表2」等）
SheetDefinition sheet = document.CreateSheet();
```

若傳入重複名稱，會自動加上序號後綴（例如 `銷售報表(1)`）。

### 存取工作表

```csharp
// 取得最後一個工作表（若無則自動建立）
SheetDefinition last = document.LastSheet;

// 依零基索引取得
SheetDefinition first = document.GetSheet(0);

// 安全取得
if (document.TryGetSheet(0, out SheetDefinition sheet)) {
    // ...
}
```

## 匯出

### 匯出至位元組陣列

適用於網頁回應或記憶體內處理：

```csharp
public IActionResult Download() {
    SpreadsheetDocument document = SpreadsheetManager.CreateDocument();
    SheetDefinition sheet = document.CreateSheet("報表");
    sheet.AddTemplate(BuildTemplate());

    return File(
        document.Export(),
        document.ContentType,
        $"report{document.FileNameExtension}"
    );
}
```

### 匯出至檔案

```csharp
// 若檔案存在則覆寫（預設）
document.ExportFile($@"C:\Reports\output{document.FileNameExtension}");

// 若檔案存在則拋出 IOException
document.ExportFile(
    $@"C:\Reports\output{document.FileNameExtension}",
    SpreadsheetFileMode.CreateNew
);
```

**SpreadsheetFileMode 選項：**

- `Create` - 建立新檔案，若存在則覆寫（預設）
- `CreateNew` - 建立新檔案，若存在則拋出例外

## 從 JSON 建立文件

`SpreadsheetDocument.FromJson` 可從 JSON 字串直接建立完整文件，適合將報表結構外部化配置：

```csharp
string json = File.ReadAllText("report-template.json");
SpreadsheetDocument document = SpreadsheetDocument.FromJson(json);
document.ExportFile($"output{document.FileNameExtension}");
```

JSON 格式為工作表陣列，目前支援 `Grid` 與 `RecordSet` 兩種 template 類型：

```json
[
  {
    "SheetName": "Report",
    "DefaultRowHeight": 20,
    "ColumnWidths": [
      { "Index": 0, "Width": 18 },
      { "Index": 1, "Width": 24 }
    ],
    "Templates": [
      {
        "Type": "Grid",
        "Rows": [
          {
            "Height": 24,
            "Cells": [
              { "Value": "標題", "ColumnSpan": 2, "Style": { "HorizontalAlignment": "Center" } }
            ]
          }
        ]
      }
    ]
  }
]
```

### 工作表層級支援欄位

| 欄位 | 型別 | 說明 |
| --- | --- | --- |
| `SheetName` | `string` | 工作表名稱；省略時會自動命名 |
| `DefaultRowHeight` | `number` | 工作表預設列高 |
| `Password` | `string` | 工作表保護密碼 |
| `FreezePanes` | `object` | 凍結窗格設定，需同時提供 `Row` 與 `Column` |
| `IsAutoFilterEnabled` | `boolean` | 啟用自動篩選 |
| `ColumnWidths` | `object` / `array` | 欄寬設定，可用索引物件或 `{ Index, Width }` 陣列 |
| `PageSettings` | `object` | 列印設定，目前支援 `PageOrientation` 與 `PaperSize` |
| `Metadata` | `object` | 附加自訂資料，供 renderer 或擴充邏輯讀取 |
| `Templates` | `array` | 工作表上的 template 定義 |

`FreezePanes` 範例：

```json
{
  "FreezePanes": {
    "Row": 1,
    "Column": 0
  }
}
```

`ColumnWidths` 可用兩種格式：

```json
{
  "ColumnWidths": {
    "0": 18,
    "1": 24
  }
}
```

```json
{
  "ColumnWidths": [
    { "Index": 0, "Width": 18 },
    { "Index": 1, "Width": 24 }
  ]
}
```

### Grid template 支援欄位

| 欄位 | 型別 | 說明 |
| --- | --- | --- |
| `Type` | `string` | 固定為 `Grid` |
| `Rows` | `array` | 列定義 |
| `Rows[].Height` | `number` | 該列列高 |
| `Rows[].Cells` | `array` | 儲存格定義 |
| `Rows[].Cells[].Value` | `any` | 固定值 |
| `Rows[].Cells[].Formula` | `string` | 公式字串 |
| `Rows[].Cells[].ColumnSpan` | `number` | 欄合併數，預設 `1` |
| `Rows[].Cells[].RowSpan` | `number` | 列合併數，預設 `1` |
| `Rows[].Cells[].Style` | `object` | 儲存格樣式 |

`Grid` 的 `Value` 與 `Formula` 只能擇一設定。

### RecordSet template 支援欄位

| 欄位 | 型別 | 說明 |
| --- | --- | --- |
| `Type` | `string` | 固定為 `RecordSet` |
| `HeaderHeight` | `number` | 標題列高度 |
| `RecordHeight` | `number` | 資料列高度 |
| `Columns` | `array` | 欄位定義 |
| `Records` | `array` | 資料列；每筆記錄會轉成 `Dictionary<string, object?>` |

`Columns[]` 支援欄位：

| 欄位 | 型別 | 說明 |
| --- | --- | --- |
| `HeaderText` | `string` | 欄位標題 |
| `HeaderStyle` | `object` | 標題樣式 |
| `FieldStyle` | `object` | 資料儲存格樣式 |
| `FieldKey` | `string` | 從 `Records[]` 取值的欄位名稱 |
| `Value` | `any` | 固定值 |
| `Formula` | `string` | 固定公式 |
| `Children` | `array` | 子欄位，建立多層標題 |

`RecordSet` 欄位的 `FieldKey`、`Value`、`Formula` 只能擇一設定。

`FieldKey` 支援巢狀物件路徑，例如：

```json
{
  "HeaderText": "City",
  "FieldKey": "Customer.Address.City"
}
```

### Style 與 Font 支援欄位

`Grid` 的 `Style`，以及 `RecordSet` 的 `HeaderStyle`、`FieldStyle`，目前支援以下欄位：

| 欄位 | 型別 | 說明 |
| --- | --- | --- |
| `HorizontalAlignment` | `string` / `number` | 水平對齊 |
| `VerticalAlignment` | `string` / `number` | 垂直對齊 |
| `HasBorder` | `boolean` | 是否套用框線 |
| `WrapText` | `boolean` | 是否自動換行 |
| `BackgroundColor` | `string` / `object` | 背景色 |
| `DataFormat` | `string` / `null` | Excel 格式字串 |
| `IsLocked` | `boolean` | 是否鎖定儲存格 |
| `Font` | `object` | 字型設定 |

`Font` 支援欄位：

| 欄位 | 型別 | 說明 |
| --- | --- | --- |
| `Name` | `string` / `null` | 字型名稱 |
| `Size` | `number` | 字型大小 |
| `Color` | `string` / `object` | 字型顏色 |
| `Style` | `string` / `number` / `array` | 字型樣式，可用 `Bold`、`Italic`、`Underline`、`Strikeout` 或其數值 |

色彩格式支援：

- 命名色彩，例如 `Red`
- HTML 色碼，例如 `#FFAA00`
- ARGB / RGB 物件，例如 `{ "R": 255, "G": 170, "B": 0 }`

列舉值支援名稱或數值，且字串比對不分大小寫。

### 完整範例

```json
[
  {
    "SheetName": "Orders",
    "DefaultRowHeight": 20,
    "Password": "sheet-pass",
    "FreezePanes": { "Row": 1, "Column": 0 },
    "IsAutoFilterEnabled": true,
    "ColumnWidths": {
      "0": 12,
      "1": 24,
      "2": 18
    },
    "PageSettings": {
      "PageOrientation": "Landscape",
      "PaperSize": "A4"
    },
    "Metadata": {
      "reportType": "monthly",
      "generatedBy": "system"
    },
    "Templates": [
      {
        "Type": "Grid",
        "Rows": [
          {
            "Height": 28,
            "Cells": [
              {
                "Value": "Order Summary",
                "ColumnSpan": 3,
                "Style": {
                  "HorizontalAlignment": "Center",
                  "HasBorder": true,
                  "BackgroundColor": "#1F4E79",
                  "Font": {
                    "Name": "Calibri",
                    "Size": 14,
                    "Color": "White",
                    "Style": ["Bold"]
                  }
                }
              }
            ]
          }
        ]
      },
      {
        "Type": "RecordSet",
        "HeaderHeight": 22,
        "RecordHeight": 20,
        "Columns": [
          { "HeaderText": "Order ID", "FieldKey": "OrderId" },
          {
            "HeaderText": "Customer",
            "FieldKey": "Customer.Name"
          },
          {
            "HeaderText": "Amount",
            "FieldKey": "Amount",
            "FieldStyle": {
              "HorizontalAlignment": "Right",
              "DataFormat": "#,##0.00"
            }
          }
        ],
        "Records": [
          {
            "OrderId": 1001,
            "Customer": {
              "Name": "Northwind"
            },
            "Amount": 1250.40
          },
          {
            "OrderId": 1002,
            "Customer": {
              "Name": "Adventure Works"
            },
            "Amount": 980.00
          }
        ]
      }
    ]
  }
]
```

### 目前未支援的 JSON 功能

- `DataTableTemplate`。
- `MergedTemplate`。
- Grid / RecordSet 的 `DataValidation`。
- 需要以程式碼動態計算的 generator 邏輯。
- 自訂 template 型別，除非先自行註冊對應的 JSON parser。

### 註冊自訂 JSON parser

`SpreadsheetDocument.FromJson(...)` 會透過 `JsonTemplateRegistry` 依 `Type` 欄位尋找對應的 parser。

內建已註冊的型別有：

- `Grid`
- `RecordSet`

若要支援自訂 template，可實作 `ITemplateJsonParser`，再於呼叫 `FromJson(...)` 前註冊：

```csharp
using System.Text.Json;
using CloudyWing.SpreadsheetExporter;
using CloudyWing.SpreadsheetExporter.Templates;
using CloudyWing.SpreadsheetExporter.Templates.Json;

public class ReportHeaderTemplateJsonParser : ITemplateJsonParser {
    public string TypeName => "ReportHeader";

    public ISheetTemplate Parse(JsonElement element) {
        string title = element.GetProperty("Title").GetString()!;
        int columnSpan = element.GetProperty("ColumnSpan").GetInt32();
        return new ReportHeaderTemplate(title, columnSpan);
    }
}

JsonTemplateRegistry.Register(new ReportHeaderTemplateJsonParser());

SpreadsheetManager.SetRenderer(static () => new ExcelRenderer());
SpreadsheetDocument document = SpreadsheetDocument.FromJson(json);
```

對應 JSON：

```json
{
  "Type": "ReportHeader",
  "Title": "Q1 Sales Report",
  "ColumnSpan": 4
}
```

若需要完全重設註冊內容，可先呼叫 `JsonTemplateRegistry.Clear()`，再視需要呼叫 `JsonTemplateRegistry.RegisterBuiltins()` 與自訂 parser 註冊。

## 自訂 renderer 行為

`ExcelRenderer` 提供 `OnSheetCreated` 虛擬方法，可在工作表建立後透過 ClosedXML API 進行額外設定：

```csharp
using ClosedXML.Excel;
using CloudyWing.SpreadsheetExporter;
using CloudyWing.SpreadsheetExporter.Renderer.ClosedXML;

public class CustomExcelRenderer : ExcelRenderer {
    protected override void OnSheetCreated(IXLWorksheet worksheet, SheetLayout context) {
        // 透過 ClosedXML API 進行自訂設定
        worksheet.SheetView.ShowGridLines = false;
    }
}

SpreadsheetManager.SetRenderer(static () => new CustomExcelRenderer());
```

## 相關主題

- [入門指南](getting-started.md) - 從零開始設定 SpreadsheetExporter
- [SheetDefinition](sheet-definition.md) - 工作表管理與設定
- [Templates](templates.md) - 探索不同的資料結構化方式
- [Migration Notes](migration-notes.md) - 檢視不同版本之間的 JSON 與 API 調整
- [自訂樣式](customization.md) - 透過 SpreadsheetManager 設定全域儲存格樣式
