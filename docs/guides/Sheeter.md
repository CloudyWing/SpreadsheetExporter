# Sheeter

用來建立後續產生 Spreadsheet Sheet 的所需資訊。

## 使用教學

### 更改 Sheet 名稱
建立 Sheeter 以後，仍可變更 Sheet 名稱
```csharp
Sheeter sheeter = exporter.CreateSheeter();
sheeter.SheetName = "New Sheet1";
```

### 設定鎖定儲存格的密碼
```csharp
Sheeter sheeter = exporter.CreateSheeter();
sheeter.Password = "Your Sheet Password";
```
可用擴充方法簡化。
```csharp
public static class ExporterBaseExtensions {
    public static Sheeter CreateSheeterWithPassword(this ExporterBase exporter, string sheetName = null) {
        Sheeter sheeter = exporter.CreateSheeter(sheetName);
        sheeter.Password = "Your Sheet Password";
        return sheeter;
    }
}
```

### 設定欄寬
使用 `SetColumnWidth()` 設定指定欄的寬度，欄 Index 從 0 開始計算。
```csharp
Sheeter sheeter = exporter.CreateSheeter();
sheeter.SetColumnWidth(0, 10d);

// 可使用 HiddenColumn 設定隱藏欄寬
sheeter.SetColumnWidth(1, Constants.HiddenColumn);

// 可使用 AutoFitColumnWidth 設定自動調整欄寬
sheeter.SetColumnWidth(2, Constants.AutoFitColumnWidth);
```

### 加入 `ITemplate`
如果 Sheeter 加入多個 `ITemplate`，會以列為單位往下疊加內容，舉例來說 template1 佔 3 列，template2 的內容會從第 4 列開始產生。
```csharp
sheeter.AddTemplate(template1);
sheeter.AddTemplate(template2);

// 可使用 AddTemplates 一次加入多個 ITemplate
sheeter.AddTemplates(template3, template4);
```

## 其他教學
* [入門教學](./%E5%85%A5%E9%96%80%E6%95%99%E5%AD%B8.md)
* [Exporter](./Exporter.md)
* [Template](./Template.md)
* [設定預設的儲存格樣式](./%E8%A8%AD%E5%AE%9A%E9%A0%90%E8%A8%AD%E7%9A%84%E5%84%B2%E5%AD%98%E6%A0%BC%E6%A8%A3%E5%BC%8F.md)
