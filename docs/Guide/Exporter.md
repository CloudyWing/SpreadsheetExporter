# Exporter

## 設置
用來設置與 Workbook 相關的資訊，提供以下設置。

* Password：設定 Workbook 的密碼。
* DefaultBasicSheetName：設定新 Sheet 的預設基本名稱，建立 Sheeter 時，若未指定 Sheet 名稱，產生的預設名稱為 `DefaultBasicSheetName{Index}`，`{Index}` 為從 1 開始的流水號，預設值為 `工作表`。
* SpreadsheetExportingEvent：設定產出 Spreadsheet 前觸發的事件，Debug 寫 Log 用。
* SpreadsheetExportedEvent：設定產出 Spreadsheet 後觸發的事件，Debug 寫 Log 用。
* SheetCreatedEvent：設定當 Exporter 建立 Sheet 完成時會觸發的事件，SheetCreatedEventArgs 包含以下參數：  
  * SheetObject：Exporter 將其設定為所建立 Sheet 的類別，例如 NPOI 的 `ISheet` 或 EPPlus 的 `ExcelSheet`，以便使用該類別支援的功能。
  * SheeterContext：用於建立 Sheet 的相關資訊。

可以搭配 `SpreadsheetManager.SetExporter()` 設定 Exporter 的初始值。
```csharp
SpreadsheetManager.SetExporter(() => {
    ExcelExporter exporter = new ExcelExporter();
    exporter.Password = "Your Workbook Password";
    exporter.DefaultBasicSheetName = "New Sheet";

    return exporter;
});
```

## 使用說明
### Spreadsheet 資訊
不同 Spreadsheet 的 Content-Type 和 副檔名可能會不一樣，可使用 `ContentType` 和 `FileNameExtension` 取得當前 Exporter 的資訊。
```csharp
ExporterBase exporter = SpreadsheetManager.CreateExporter();
string contentType = exporter.ContentType;
string fileNameExtension = exporter.FileNameExtension;
```

### 建立或取得 Sheeter
```csharp
ExporterBase exporter = SpreadsheetManager.CreateExporter();
// 後續產出 Spreadsheet，會建立一個名為 Test 的 Sheet
Sheeter sheeter = exporter.CreateSheeter("Test");
// 未傳入參數，則 Sheet 名稱會是 DefaultBasicSheetName{Index}，e.g.「工作表1」
Sheeter sheeter2 = exporter.CreateSheeter();

// 取得最後一個 Sheeter，若還沒建立，則建立一個
Sheeter sheeter3 = exporter.LastSheeter;

// 取得第一個 Sheeter
Sheeter sheeter4 = exporter.GetSheeter(0);
```

### 產出 Spreadsheet
#### `Export()`
使用於不需要實際產出檔案在主機上的場合，如 Web 產出檔案給使用者下載。
```csharp
public IActionResult Download() {
    ExporterBase exporter = SpreadsheetManager.CreateExporter();
    exporter.CreateSheeter();

    return File(exporter.Export(), exporter.ContentType, $"Spreadsheet{exporter.FileNameExtension}");
}
```

#### `ExportFile()`
產出檔案在主機上。
```csharp
ExporterBase exporter = SpreadsheetManager.CreateExporter();
exporter.CreateSheeter();
exporter.ExportFile($@"C:\Sample{exporter.FileNameExtension}");

// 如果檔案已存在是拋出例外，而非覆蓋檔案，則使用 SpreadsheetFileMode.CreateNew
exporter.ExportFile($@"C:\Sample{exporter.FileNameExtension}", SpreadsheetFileMode.CreateNew);
```
* ExportFile(string path, SpreadsheetFileMode fileMode)：產出 Spreadsheet 檔案，SpreadsheetFileMode 用來指定若檔案存在時是否覆蓋檔案，可選擇以下值：
    * Create：建立新檔案，若檔案已存在覆蓋舊檔案，預設值。
    * CreateNew：建立新檔案，若檔案已存在則拋出異常。

## 其他教學
* [入門教學](./%E5%85%A5%E9%96%80%E6%95%99%E5%AD%B8.md)
* [Sheeter](./Sheeter.md)
* [Template](./Template.md)
* [設定預設的儲存格樣式](./%E8%A8%AD%E5%AE%9A%E9%A0%90%E8%A8%AD%E7%9A%84%E5%84%B2%E5%AD%98%E6%A0%BC%E6%A8%A3%E5%BC%8F.md)
