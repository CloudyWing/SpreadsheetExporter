# Templates

Templates 定義工作表儲存格的結構與內容。SpreadsheetExporter 提供多種 Template 類型以滿足不同使用情境。

## 本文內容

- [GridTemplate](#gridtemplate) - 手動逐格配置
- [DataTableTemplate](#datatabletemplate) - 基於 DataTable 的匯出
- [RecordSetTemplate](#recordsettemplate) - 強型別集合
- [MergedTemplate](#mergedtemplate) - 合併多個 Templates
- [自訂 Templates](#自訂-templates) - 建立您自己的 Template

## GridTemplate

`GridTemplate` 提供精細的儲存格配置控制，類似於 HTML 的 `<table>`、`<tr>` 與 `<td>`。支援方法鏈（Method Chaining）以簡化程式碼。

### 建立列

```csharp
GridTemplate template = new GridTemplate();

// 預設列高
template.CreateRow();

// 指定列高
template.CreateRow(20d);

// 隱藏列
template.CreateRow(Constants.HiddenRow);

// 自動調整列高
template.CreateRow(Constants.AutoFitRowHeight);
```

### 建立儲存格

```csharp
GridTemplate template = new GridTemplate();
template.CreateRow();

// 簡單儲存格
template.CreateCell("Value1_1");

// 合併儲存格（RowSpan=3, ColumnSpan=2）
template.CreateCell("Value1_2", 3, 2);

// 自訂樣式的儲存格
template.CreateCell("Value1_3", 1, 1, new CellStyle());

// 使用方法鏈與公式
template.CreateRow()
    .CreateCell("Value2_1")
    .CreateCell((cell, row) => $"{cell} + {row}"); // 公式（索引從 0 開始）

/*
輸出：
| -------- | -------- | -------- | -------- | -------- | 
| Value1_1 | Value1_2                       | Value1_3 |
| -------- |                                | -------- |
| Value2_1 |                                | =1 + 1   |
| -------- | -------- | -------- | -------- | -------- |
*/
```

## DataTableTemplate

直接將 `System.Data.DataTable` 匯出至 Excel，並自動對應欄位。

```csharp
// 建立 DataTable
System.Data.DataTable dataTable = new System.Data.DataTable();
dataTable.Columns.Add("Name", typeof(string));
dataTable.Columns.Add("Age", typeof(int));

dataTable.Rows.Add("John", 30);
dataTable.Rows.Add("Mary", 25);

// 建立 Template
DataTableTemplate template = new DataTableTemplate(dataTable);

// 設定列高
template.HeaderHeight = 25;
template.RecordHeight = 20;

/*
輸出：
| ---- | --- |
| Name | Age |
| ---- | --- |
| John | 30  |
| ---- | --- |
| Mary | 25  |
| ---- | --- |
*/
```

## RecordSetTemplate

最強大的 Template，適用於強型別資料集合。提供完整的欄位設定、資料轉換與樣式控制。

### 基本用法

```csharp
public class Record {
    public int Id { get; set; }
    public string Name { get; set; }
}

List<Record> source = new List<Record> {
    new Record { Id = 0, Name = "Marry" },
    new Record { Id = 1, Name = "Terry" }
};

RecordSetTemplate<Record> template = new RecordSetTemplate<Record>(source);
template.Columns.Add("編號", x => x.Id);
template.Columns.Add("姓名", x => x.Name);

/*
輸出：
| ---- | ---- | 
| 編號 | 姓名 |
| ---- | ---- |
| 0    | Marry |
| ---- | ---- |
| 1    | Terry |
| ---- | ---- |
*/
```

### 資料轉換

在匯出過程中轉換值：

```csharp
// 轉換屬性值
template.Columns.Add("大寫姓名", x => x.Name, x => x.UseValue(y => y.Value.ToUpper()));

// 合併多個屬性
template.Columns.Add("合併資料", x => x.UseValue(y => y.Record.Id + y.Record.Name));

/*
輸出：
| -------- | -------- | 
| 大寫姓名 | 合併資料 |
| -------- | -------- |
| MARRY    | 0Marry   |
| -------- | -------- |
| TERRY    | 1Terry   |
| -------- | -------- |
*/
```

### 自訂儲存格樣式

套用條件樣式：

```csharp
CellStyleConfiguration cellStyles = SpreadsheetManager.DefaultCellStyles;

template.Columns.Add(
    "狀態", 
    x => x.Id,
    // 標題樣式：紅色背景
    cellStyles.HeaderStyle with {
        BackgroundColor = Color.Red
    },
    // 條件式欄位樣式
    x => x.Value == 0
        ? cellStyles.FieldStyle with { BackgroundColor = Color.Blue }
        : cellStyles.FieldStyle with { BackgroundColor = Color.Yellow }
);
```

### 多層標題

建立分組欄位標題：

```csharp
template.Columns.Add("群組1")
    .AddChildToLast("編號", x => x.Id)
    .AddChildToLast("姓名", x => x.Name)
    .Add("群組2")
    .AddChildToLast("子群組1");

DataColumnCollection<Record> lastChildColumns = template.Columns.Last().ChildColumns;
lastChildColumns.AddChildToLast("編號", x => x.Id)
    .AddChildToLast("姓名", x => x.Name);

template.Columns.AddChildToLast("子群組2");
lastChildColumns = template.Columns.Last().ChildColumns
    .AddChildToLast("編號", x => x.Id)
    .AddChildToLast("姓名", x => x.Name);

/*
輸出：
| ---- | ---- | ---- | ---- | ---- | ---- |
|              |          群組2              |
|    群組1     | -------- | -------- |
|              | 子群組1  | 子群組2  |
| ---- | ---- | ---- | ---- | ---- | ---- |
| 編號 | 姓名 | 編號 | 姓名 | 編號 | 姓名 |
| ---- | ---- | ---- | ---- | ---- | ---- |
*/
```

### 公式

在儲存格中使用 Excel 公式：

```csharp
template.Columns.Add("公式", x => x.UseFormula(y => $"{y.CellIndex} + {y.RowIndex}"));

/*
輸出（第一筆資料在索引 1）：
| ---- |
| 公式 |
| ---- |
| =0+1 |
| ---- |
| =0+2 |
| ---- |
*/
```

### 凍結窗格與自動篩選

```csharp
RecordSetTemplate<Record> template = new RecordSetTemplate<Record>(source);
// ... 加入欄位 ...

// 凍結標題列（依據標題層數自動決定）
template.IsFreezeHeader = true;

// 啟用自動篩選（包含標題與資料範圍）
template.IsAutoFilterEnabled = true;

// 設定列高
template.HeaderHeight = 30;
template.RecordHeight = 25;
```

## 資料驗證

為儲存格加入資料驗證規則以限制輸入。

### 在 RecordSetTemplate 中使用

```csharp
template.Columns.Add("年齡", x => x.Age, provider => provider.UseDataValidation(x => new DataValidation {
    ValidationType = DataValidationType.Integer,
    Operator = DataValidationOperator.Between,
    Value1 = 18,
    Value2 = 65,
    ErrorTitle = "輸入錯誤",
    ErrorMessage = "年齡必須在 18 到 65 歲之間",
    IsErrorAlertShown = true
}));
```

### 在 GridTemplate 中使用

```csharp
GridTemplate template = new GridTemplate();
template.CreateRow();
template.CreateCell(cell => {
    cell.ValueGenerator = (c, r) => "請選擇";
    cell.DataValidationGenerator = (c, r) => new DataValidation {
        ValidationType = DataValidationType.List,
        ListItems = new[] { "選項 A", "選項 B", "選項 C" },
        IsDropdownShown = true
    };
});
```

## MergedTemplate

合併多個 Templates。適用於建立複雜版面配置。

```csharp
MergedTemplate merged = new MergedTemplate();
merged.AddTemplate(headerTemplate);
merged.AddTemplate(dataTemplate);
merged.AddTemplate(footerTemplate);

sheeter.AddTemplate(merged);
```

## 自訂 Templates

實作 `ITemplate` 介面以建立可重複使用的自訂 Template：

```csharp
public class ReportInfoTemplate : ITemplate {
    private readonly GridTemplate gridTemplate = new GridTemplate();

    public ReportInfoTemplate(string title, string user, int colSpan) {
        CellStyleConfiguration cellStyles = SpreadsheetManager.DefaultCellStyles;
        CellStyle titleStyle = cellStyles.CellStyle with {
            HorizontalAlignment = HorizontalAlignment.Center,
            Font = cellStyles.CellStyle.Font with { Size = 14 }
        };

        // 標題列
        gridTemplate.CreateRow();
        gridTemplate.CreateCell(title, colSpan, cellStyle: titleStyle);

        // 資訊列
        int leftColSpan = colSpan / 2;
        gridTemplate.CreateRow();
        gridTemplate.CreateCell($"使用者：{user}", leftColSpan, cellStyle: cellStyles.CellStyle);
        gridTemplate.CreateCell(
            $"產生時間：{DateTime.Now:yyyy-MM-dd HH:mm:ss}", 
            colSpan - leftColSpan, 
            cellStyle: cellStyles.CellStyle
        );
    }

    public TemplateContext GetContext() {
        return gridTemplate.GetContext();
    }
}
```

## 相關主題

- [入門指南](getting-started.md) - 了解如何在專案中開始使用 Templates
- [Exporters](exporters.md) - 學習 Template 與 Exporter 的整體協作流程
- [Sheeters](sheeters.md) - 了解如何將 Templates 加入 Sheeter
- [自訂樣式](customization.md) - 為 Template 中的儲存格套用自訂樣式
