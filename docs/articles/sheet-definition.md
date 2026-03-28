# SheetDefinition

`SheetDefinition` 代表文件中的單一工作表，包含 templates、欄寬設定、頁面設定與保護密碼。

## 基本設定

### 工作表名稱

`CreateSheet` 傳入的名稱即為工作表顯示名稱，建立後仍可修改：

```csharp
SheetDefinition sheet = document.CreateSheet("銷售報表");
sheet.SheetName = "Q1 銷售報表";
```

### 預設列高

```csharp
SheetDefinition sheet = document.CreateSheet("報表", defaultRowHeight: 20);
// 或建立後設定
sheet.DefaultRowHeight = 20;
```

### 工作表保護

為工作表設定密碼保護（保護後使用者無法編輯未解鎖的儲存格）：

```csharp
SheetDefinition sheet = document.CreateSheet("報表");
sheet.Password = "your-password";
```

## 欄位設定

使用 `SetColumnWidth` 設定欄寬（索引從 0 開始）。方法支援 Fluent 鏈式呼叫：

```csharp
sheet
    .SetColumnWidth(0, 18d)                            // 設定指定寬度（字元單位）
    .SetColumnWidth(1, Constants.HiddenColumn)         // 隱藏欄位
    .SetColumnWidth(2, Constants.AutoFitColumnWidth);  // 自動調整欄寬
```

## 頁面設定

透過 `PageSettings` 設定列印相關選項：

```csharp
sheet.PageSettings.PageOrientation = PageOrientation.Landscape;
sheet.PageSettings.PaperSize = PaperSize.A4;
```

**PageOrientation 選項：**

- `Portrait` - 直向（預設）
- `Landscape` - 橫向

## 使用 Templates

Templates 會依加入順序**垂直堆疊**。若 `template1` 佔用 3 列，`template2` 將從第 4 列開始。

```csharp
// 加入單一 template
sheet.AddTemplate(template1);

// 加入多個 templates（params 語法）
sheet.AddTemplates(template2, template3);

// 加入集合
sheet.AddTemplates(templateList);

// Fluent 鏈式呼叫
sheet
    .AddTemplate(headerTemplate)
    .AddTemplates(dataTemplate, footerTemplate);
```

> [!NOTE]
> 同一份文件可建立多個工作表，每個工作表各自維護獨立的 templates 與設定。

## Metadata

`Metadata` 字典可附加任意資訊，供自訂 renderer 使用（例如在 `OnSheetCreated` 中讀取）：

```csharp
sheet.Metadata["reportType"] = "quarterly";
sheet.Metadata["generatedBy"] = "system";
```

在自訂 renderer 中讀取：

```csharp
using ClosedXML.Excel;
using CloudyWing.SpreadsheetExporter;
using CloudyWing.SpreadsheetExporter.Renderer.ClosedXML;

public class CustomExcelRenderer : ExcelRenderer {
    protected override void OnSheetCreated(IXLWorksheet worksheet, SheetLayout context) {
        if (context.Metadata.TryGetValue("reportType", out object reportType)) {
            worksheet.Tab.Color = reportType as string == "quarterly"
                ? XLColor.Blue
                : XLColor.Gray;
        }
    }
}
```

## 用擴充方法封裝慣用設定

若某些工作表設定在專案內會重複出現，可以用擴充方法包起來，而不是讓呼叫端直接操作 `Metadata` 鍵名或重複設定流程。

```csharp
public static class SheetDefinitionExtensions {
    public static SheetDefinition SetReportType(this SheetDefinition sheet, string reportType) {
        sheet.Metadata["reportType"] = reportType;
        return sheet;
    }

    public static SheetDefinition ConfigureQuarterlyReport(this SheetDefinition sheet) {
        return sheet
            .SetReportType("quarterly")
            .SetFreezePanes(0, 1)
            .SetAutoFilterEnabled();
    }
}
```

使用時可保持呼叫端簡潔：

```csharp
document
    .CreateSheet("Q1")
    .ConfigureQuarterlyReport()
    .AddTemplate(template);
```

這類擴充方法適合包裝：

- 專案內固定會重複的 `Metadata` 鍵值。
- 常見的工作表初始化流程。
- 特定 renderer 需要的慣例設定。

## 相關主題

- [入門指南](getting-started.md) - 了解 SheetDefinition 在整體匯出流程中的角色
- [SpreadsheetDocument](spreadsheet-document.md) - 學習如何透過文件建立工作表
- [Templates](templates.md) - 探索可加入工作表的各種範本類型
- [自訂樣式](customization.md) - 設定工作表中儲存格的預設樣式
