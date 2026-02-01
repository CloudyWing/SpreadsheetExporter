# Sheeters

`Sheeter` 代表一個工作表，包含產生工作表內容所需的設定與 Templates。

## 基本設定

### 變更工作表名稱

建立 Sheeter 後仍可修改工作表名稱：

```csharp
Sheeter sheeter = exporter.CreateSheeter();
sheeter.SheetName = "新工作表1";
```

### 設定工作表密碼

為工作表設定密碼保護：

```csharp
Sheeter sheeter = exporter.CreateSheeter();
sheeter.Password = "Your Sheet Password";
```

您可以透過擴充方法簡化此模式：

```csharp
public static class ExporterBaseExtensions {
    public static Sheeter CreateSheeterWithPassword(
        this ExporterBase exporter, 
        string sheetName = null
    ) {
        Sheeter sheeter = exporter.CreateSheeter(sheetName);
        sheeter.Password = "Your Sheet Password";
        return sheeter;
    }
}
```

## 欄位設定

### 設定欄寬

使用 `SetColumnWidth()` 設定欄寬（欄位索引從 0 開始）：

```csharp
Sheeter sheeter = exporter.CreateSheeter();

// 設定特定寬度
sheeter.SetColumnWidth(0, 10d);

// 隱藏欄位
sheeter.SetColumnWidth(1, Constants.HiddenColumn);

// 自動調整欄寬
sheeter.SetColumnWidth(2, Constants.AutoFitColumnWidth);
```

## 使用 Templates

### 加入 Templates

Templates 會垂直堆疊。若 `template1` 佔用 3 列，`template2` 將從第 4 列開始。

```csharp
sheeter.AddTemplate(template1);
sheeter.AddTemplate(template2);

// 一次加入多個 Templates
sheeter.AddTemplates(template3, template4);
```

> [!NOTE]
> Templates 定義工作表的結構與內容。請參閱 [Templates](templates.md) 指南了解更多資訊。

## 相關主題

- [入門指南](getting-started.md) - 了解 Sheeter 在整體匯出流程中的角色
- [Exporters](exporters.md) - 學習如何透過 Exporter 建立 Sheeter
- [Templates](templates.md) - 探索可加入 Sheeter 的各種範本類型
- [自訂樣式](customization.md) - 設定 Sheeter 中儲存格的預設樣式
