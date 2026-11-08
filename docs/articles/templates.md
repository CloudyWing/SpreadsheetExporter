# Templates

Templates 定義工作表儲存格的結構與內容。SpreadsheetExporter 提供四種內建 template 類型，並支援自訂擴充。

## 本文內容

- [GridTemplate](#gridtemplate) - 手動逐格配置
- [DataTableTemplate](#datatabletemplate) - 基於 DataTable 的匯出
- [RecordSetTemplate](#recordsettemplate) - 強型別集合
- [MergedTemplate](#mergedtemplate) - 合併多個 templates
- [自訂 Templates](#自訂-templates) - 建立可重複使用的自訂 template

若你是透過 `SpreadsheetDocument.FromJson(...)` 建立文件，目前 JSON 支援 `Grid`、`RecordSet`、`DataTable` 與 `Merged` 四種 template 型別。完整 JSON 欄位請參考 [SpreadsheetDocument](spreadsheet-document.md) 的 JSON 章節。

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

### 資料驗證

`DataTableColumn.FieldDataValidationGenerator` 可依資料值回傳驗證規則：

```csharp
DataTableTemplate template = new DataTableTemplate(dataTable);
template.Columns[1].FieldDataValidationGenerator = value => new DataValidation {
    ValidationType = DataValidationType.Integer,
    Operator = DataValidationOperator.Between,
    Value1 = 1,
    Value2 = 120
};
```

### 依整列資料設定樣式

`FieldStyleGenerator` 只能取得該欄的值。需要依同一列其他欄位決定樣式時，改用 `FieldStyleContextGenerator`，它會收到 `DataRowContext`：

```csharp
CellStyle overdueStyle = SpreadsheetManager.DefaultCellStyles.FieldStyle with {
    BackgroundColor = Color.MistyRose
};

template.Columns[0].FieldStyleContextGenerator = context =>
    context.DataRow["Status"] as string == "Overdue"
        ? overdueStyle
        : SpreadsheetManager.DefaultCellStyles.FieldStyle;
```

| 屬性 | 說明 |
| --- | --- |
| `DataRow` | 該儲存格所在的 `DataRow` |
| `DataRowIndex` | 在 `DataTable.Rows` 中的零基索引 |
| `ColumnName` | 綁定的欄位名稱 |
| `Value` | 儲存格值，`DBNull` 會轉為 `null` |
| `CellIndex`、`RowIndex` | 儲存格的零基欄、列索引（與其他 cell generator 相同） |

同時設定兩者時以 `FieldStyleContextGenerator` 為準。

### 更換 DataTable

設定 `DataTable` 屬性時，若新資料表包含 `Columns` 目前引用的所有欄位，`Columns` 的設定（標題文字、generator 等）會保留；若缺少任一欄位，`Columns` 會依新資料表的欄位重新建立，原本的設定會被清除。

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

欄位可以指向巢狀屬性。假設 `Order` 具有 `Shipping` 物件屬性，可使用 `order => order.Shipping.City`，或字串路徑 `"Shipping.City"`。路徑中途的物件為 `null` 時，欄位值為 `null`，不會拋出例外。

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
// context.RowIndex 為儲存格在工作表中的零基列索引（包含標題列與上方的 templates），加 1 即為 Excel 列號
template.Columns.Add(
    "含稅金額",
    configureGenerators: config => config.UseFormula(
        context => $"C{context.RowIndex + 1}*1.05"
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

### 欄寬

`Add` 提供帶 `width` 參數的多載，可直接指定欄寬（字元單位），也可使用 `Constants.HiddenColumn` 隱藏欄位或 `Constants.AutoFitColumnWidth` 自動調整：

```csharp
template.Columns
    .Add("編號", order => order.Id, 8)
    .Add("客戶", order => order.Customer, 30)
    .Add<decimal>("金額", "Amount", Constants.AutoFitColumnWidth);

// 需要同時指定樣式時，可直接設定欄位的 Width 屬性
template.Columns.GetLastColumn().Width = 16;
```

只有最底層的欄位會套用欄寬。同一欄若也以 `SheetDefinition.SetColumnWidth` 設定，以工作表設定為準；同一工作表有多個 `RecordSetTemplate<T>` 指定同一欄時，以先加入的 template 為準。

欄寬只對直接加入工作表的 `RecordSetTemplate<T>` 生效，放進 `MergedTemplate` 的 `RecordSetTemplate<T>` 不會套用，需改用 `SheetDefinition.SetColumnWidth`。欄寬為負值（`Constants.AutoFitColumnWidth` 除外）時，會拋出 `ArgumentOutOfRangeException`。

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

清單驗證的 `ListItems` 會以逗號串接成 Excel 字串常值。項目中的雙引號會自動跳脫；Excel 無法跳脫逗號，含逗號的項目會被拆成多個選項。ClosedXML renderer 的清單來源（含前後雙引號）上限為 255 個字元，超過時拋出 `ArgumentOutOfRangeException`。

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
