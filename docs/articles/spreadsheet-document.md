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

JSON 格式為工作表陣列，每個工作表可包含 `Grid` 和 `RecordSet` 類型的 templates：

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
- [自訂樣式](customization.md) - 透過 SpreadsheetManager 設定全域儲存格樣式
