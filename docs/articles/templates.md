# Templates

Templates 定義工作表儲存格的結構與內容。SpreadsheetExporter 提供四種內建 template 類型，並支援自訂擴充。

## 本文內容

- [GridTemplate](#gridtemplate) - 手動逐格配置
- [DataTableTemplate](#datatabletemplate) - 基於 DataTable 的匯出
- [RecordSetTemplate](#recordsettemplate) - 強型別集合
- [MergedTemplate](#mergedtemplate) - 合併多個 templates
- [自訂 Templates](#自訂-templates) - 建立可重複使用的自訂 template

若你是透過 `SpreadsheetDocument.FromJson(...)` 建立文件，目前 JSON 只支援 `Grid` 與 `RecordSet` 兩種 template 型別。完整 JSON 欄位請參考 [SpreadsheetDocument](spreadsheet-document.md) 的 JSON 章節。

## GridTemplate

`GridTemplate` 提供精細的儲存格配置控制，類似 HTML 的 `<table>`、`<tr>` 與 `<td>`。支援方法鏈（Method Chaining）。

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

// 簡單值儲存格
template.CreateCell("Hello");

// 合併儲存格（columnSpan=2, rowSpan=2）
template.CreateCell("Merged", columnSpan: 2, rowSpan: 2);

// 自訂樣式
CellStyle style = new CellStyle(HorizontalAlignment: HorizontalAlignment.Center, HasBorder: true);
template.CreateCell("Styled", cellStyle: style);
```

### 公式儲存格

```csharp
// 傳入 (cellIndex, rowIndex) 的 lambda
template.CreateRow()
    .CreateCell((cellIndex, rowIndex) => $"=SUM(A{rowIndex + 1},1)");
```

### 完整自訂儲存格

使用 `Action<Cell>` 多載可同時設定值、公式與資料驗證：

```csharp
template.CreateRow();
template.CreateCell(cell => {
    cell.ValueGenerator = (_, _) => "請選擇";
    cell.DataValidationGenerator = (_, _) => new DataValidation {
        ValidationType = DataValidationType.List,
        ListItems = ["選項 A", "選項 B", "選項 C"],
        IsDropdownShown = true
    };
});
```

## DataTableTemplate

直接將 `System.Data.DataTable` 匯出至試算表，自動對應欄位名稱為標題列。

```csharp
System.Data.DataTable dataTable = new System.Data.DataTable();
dataTable.Columns.Add("Name", typeof(string));
dataTable.Columns.Add("Age", typeof(int));
dataTable.Rows.Add("John", 30);
dataTable.Rows.Add("Mary", 25);

DataTableTemplate template = new DataTableTemplate(dataTable) {
    HeaderHeight = 25,
    RecordHeight = 20
};

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

適用於強型別資料集合，支援欄位綁定、資料轉換、條件樣式、多層標題與資料驗證。

### 基本用法

```csharp
public class Order {
    public int Id { get; set; }
    public string Customer { get; set; }
    public decimal Amount { get; set; }
}

List<Order> orders = [
    new Order { Id = 1, Customer = "Northwind", Amount = 1250.40m },
    new Order { Id = 2, Customer = "Adventure Works", Amount = 980.00m }
];

RecordSetTemplate<Order> template = new RecordSetTemplate<Order>(orders);
template.Columns.Add<int>("編號", order => order.Id);
template.Columns.Add<string>("客戶", order => order.Customer);
template.Columns.Add<decimal>("金額", order => order.Amount);
```

### 資料轉換

使用 `configureGenerators` 參數自訂值或公式產生方式：

```csharp
// 轉換值
template.Columns.Add<string>(
    "大寫客戶",
    order => order.Customer,
    configureGenerators: config => config.UseValue(ctx => ctx.Value.ToUpper())
);

// 不綁定欄位，完全自訂值
template.Columns.Add(
    "合併資料",
    configureGenerators: config => config.UseValue(
        ctx => $"{ctx.Record.Id} - {ctx.Record.Customer}"
    )
);
```

### 公式

```csharp
// context.RowIndex 為資料列的零基索引（標題列不計入）
template.Columns.Add(
    "含稅金額",
    configureGenerators: config => config.UseFormula(
        context => $"C{context.RowIndex + 2}*1.05"
    )
);
```

### 條件樣式

```csharp
CellStyle currencyStyle = new CellStyle(
    HorizontalAlignment: HorizontalAlignment.Right,
    DataFormat: "#,##0.00"
);

template.Columns.Add<decimal>(
    "金額",
    order => order.Amount,
    fieldStyleGenerator: ctx => ctx.Value >= 1000
        ? currencyStyle with { BackgroundColor = Color.LightGreen }
        : currencyStyle
);
```

### 多層標題

```csharp
template.Columns
    .Add<int>("編號", order => order.Id)
    .Add("客戶資訊")
    .GetLastColumn().ChildColumns
        .Add("客戶", order => order.Customer)
        .Add("地區", order => order.Region);

/*
輸出標題：
| ---- | ------------ |
|      |   客戶資訊   |
| 編號 | 客戶 | 地區 |
*/
```

### 搭配工作表設定凍結窗格與自動篩選

```csharp
RecordSetTemplate<Order> template = new RecordSetTemplate<Order>(orders);

SheetDefinition sheet = document.CreateSheet("Orders");
sheet
    .AddTemplate(template)
    .SetFreezePanes(0, 1)
    .SetAutoFilterEnabled();
```

### 資料驗證

```csharp
template.Columns.Add<string>(
    "地區",
    order => order.Region,
    configureGenerators: config => config.UseDataValidation(_ => new DataValidation {
        ValidationType = DataValidationType.List,
        ListItems = ["North", "South", "East", "West"],
        ErrorTitle = "無效地區",
        ErrorMessage = "請選擇有效的銷售地區。",
        IsErrorAlertShown = true
    })
);
```

## MergedTemplate

合併多個 templates，垂直堆疊輸出。適合將標題、資料表與頁尾組合成單一 template 傳入工作表。

```csharp
MergedTemplate merged = new MergedTemplate(headerTemplate, dataTemplate, footerTemplate);
sheet.AddTemplate(merged);
```

## 自訂 Templates

實作 `ISheetTemplate` 介面可建立可重複使用的自訂 template：

```csharp
public class ReportHeaderTemplate : ISheetTemplate {
    private readonly GridTemplate gridTemplate = new GridTemplate();

    public ReportHeaderTemplate(string title, int columnSpan) {
        CellStyle titleStyle = SpreadsheetManager.DefaultCellStyles.GridCellStyle with {
            HorizontalAlignment = HorizontalAlignment.Center,
            Font = SpreadsheetManager.DefaultCellStyles.GridCellStyle.Font with { Size = 14 }
        };

        gridTemplate.CreateRow(28);
        gridTemplate.CreateCell(title, columnSpan: columnSpan, cellStyle: titleStyle);

        gridTemplate.CreateRow();
        gridTemplate.CreateCell(
            $"產生時間：{DateTimeOffset.Now:yyyy-MM-dd HH:mm:ss zzz}",
            columnSpan: columnSpan
        );
    }

    public int ColumnSpan => gridTemplate.ColumnSpan;

    public int RowSpan => gridTemplate.RowSpan;

    public TemplateLayout GetLayout() => gridTemplate.GetLayout();
}
```

使用方式：

```csharp
sheet.AddTemplate(new ReportHeaderTemplate("Q1 銷售報表", columnSpan: 4));
sheet.AddTemplate(dataTemplate);
```

## 相關主題

- [入門指南](getting-started.md) - 了解如何在專案中開始使用 templates
- [SpreadsheetDocument](spreadsheet-document.md) - 了解 template 與文件的整體協作流程
- [SheetDefinition](sheet-definition.md) - 了解如何將 templates 加入工作表
- [自訂樣式](customization.md) - 為 template 中的儲存格套用自訂樣式
