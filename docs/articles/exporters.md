# Exporters

`ISpreadsheetExporter` 介面代表一個活頁簿（Workbook），提供活頁簿層級的設定方法以及工作表管理功能。

## 設定

您可以透過 `SpreadsheetManager.SetExporter()` 全域設定 Exporter：

```csharp
SpreadsheetManager.SetExporter(() => {
    ExcelExporter exporter = new ExcelExporter();
    exporter.Password = "Your Workbook Password";
    exporter.DefaultBasicSheetName = "工作表";
    
    return exporter;
});
```

### 可用屬性

| 屬性 | 說明 |
| --- -| --- |
| `Password` | 設定活頁簿密碼保護 |
| `DefaultBasicSheetName` | 工作表名稱預設前綴（預設值：「工作表」）。自動產生的工作表名稱格式為 `{前綴}{編號}` |
| `ContentType` | 匯出格式的 MIME 類型（例如：`application/vnd.openxmlformats-officedocument.spreadsheetml.sheet`） |
| `FileNameExtension` | 匯出格式的副檔名（例如：`.xlsx`） |

### 事件

您可以透過事件掛勾匯出生命週期：

```csharp
SpreadsheetManager.SetExporter(() => {
    ExcelExporter exporter = new ExcelExporter();
    
    exporter.SpreadsheetExportingEvent += (sender, e) => {
        // 在匯出開始前觸發
        Console.WriteLine("開始匯出...");
    };
    
    exporter.SpreadsheetExportedEvent += (sender, e) => {
        // 在匯出完成後觸發
        Console.WriteLine("匯出完成！");
    };
    
    exporter.SheetCreatedEvent += (sender, e) => {
        // 在建立工作表時觸發
        // 可存取底層的工作表物件（例如 NPOI 的 ISheet 或 EPPlus 的 ExcelWorksheet）
        var sheetObject = e.SheetObject;
        var context = e.SheeterContext;
    };
    
    return exporter;
});
```

## 使用 Sheeters

### 建立 Sheeters

```csharp
ISpreadsheetExporter exporter = SpreadsheetManager.CreateExporter();

// 建立指定名稱的工作表
Sheeter sheeter = exporter.CreateSheeter("測試工作表");

// 建立自動命名的工作表（例如「工作表1」）
Sheeter sheeter2 = exporter.CreateSheeter();
```

### 取得 Sheeters

```csharp
// 取得最後一個 Sheeter（若無則建立一個）
Sheeter lastSheeter = exporter.LastSheeter;

// 依索引取得 Sheeter
Sheeter firstSheeter = exporter.GetSheeter(0);
```

## 匯出

### 匯出至位元組陣列

當不需要儲存到磁碟時（例如網頁下載），使用 `Export()`：

```csharp
public IActionResult Download() {
    ISpreadsheetExporter exporter = SpreadsheetManager.CreateExporter();
    exporter.CreateSheeter();
    
    return File(
        exporter.Export(), 
        exporter.ContentType, 
        $"Spreadsheet{exporter.FileNameExtension}"
    );
}
```

### 匯出至檔案

使用 `ExportFile()` 儲存到磁碟：

```csharp
ISpreadsheetExporter exporter = SpreadsheetManager.CreateExporter();
exporter.CreateSheeter();

// 若檔案存在則覆寫
exporter.ExportFile($@"C:\Sample{exporter.FileNameExtension}");

// 若檔案存在則拋出例外
exporter.ExportFile(
    $@"C:\Sample{exporter.FileNameExtension}", 
    SpreadsheetFileMode.CreateNew
);
```

**SpreadsheetFileMode 選項：**

- `Create` - 建立新檔案，若存在則覆寫（預設）
- `CreateNew` - 建立新檔案，若存在則拋出例外

## 相關主題

- [入門指南](getting-started.md) - 從零開始設定 SpreadsheetExporter
- [Sheeters](sheeters.md) - 學習如何為 Exporter 建立與管理工作表
- [Templates](templates.md) - 了解如何在 Sheeter 中使用各種範本類型
- [自訂樣式](customization.md) - 透過 SpreadsheetManager 設定全域儲存格樣式
