# Template

用來建立儲存格內容資訊的範本。

## GridTemplate
使用 `CreateRow()` 和 `CreateCell()` 來建立儲存格資訊，可以用 HTML 的 `tr` 和 `td` 來理解會比較懂此 Template 的概念，可使用 Method Chaining 來簡化操作。

使用 `CreateRow()` 建立新列，並指定列高。
```csharp
GridTemplate template = new GridTemplate();
// 預設列高 16.5d
template.CreateRow();

// 可設定列高
template.CreateRow(20d);

// 使用 HiddenRow 產生隱藏列
template.CreateRow(Constants.HiddenRow);

// 使用 AutoFiteRowHeight 產生自動產生列高的列
template.CreateRow(Constants.AutoFiteRowHeight);
```

使用 `CreateCell()` 建立儲存格資訊，
```csharp
GridTemplate template = new GridTemplate();
template.CreateRow();
template.CreateCell("Value1_1");
// 建立 ColumnSpan = 2、RowSpan = 3，的儲存格
template.CreateCell("Value1_2", 3, 2);
// 可傳入型別為 CellStyle 的參數來設定儲存格樣式
template.CreateCell("Value1_3", 1, 1, new CellStyle());

// Method Chaining
template.CreateRow()
    .CreateCell("Value2_1") // 如果 ColumnSpan 傳入 2，會造成座標重疊而發生 Exception
    .CreateCell((cell, row) => $"{cell} + {row}"); // 可設定公式，cell 和 row，從 0 開始計算

/*
產出
| -------- | -------- | -------- | -------- | -------- | 
| Value1_1 | Value1_2                       | Value1_3 |
| -------- |                                | -------- |
| Value2_1 |                                | =4 + 1   |
| -------- | -------- | -------- | -------- | -------- |
*/
```

## RecordSetTemplate
使用設定 Data Source 和 Data Columns 來建立儲存格資訊，可使用 Method Chaining 來簡化操作。

```csharp
public class Record {
    public int Id { get; set; }
    public string Name { get; set; }
}

List<Record> source = new List<Record> {
    new Record { Id = 0, Name = "Marry" },
    new Record { Id = 1, Name = "Terry" }
};

RecordSetTemplate<Record> template = new RecordSetTemplate<Record(source);
template.Columns.Add("Index", x => x.Id);
template.Columns.Add("Name", x => x.Name);

/*
產出
| ------ | ----- | 
| Index  | Name  |
| ------ | ----- |
| 0      | Marry |
| ------ | ----- |
| 1      | Terry |
| ------ | ----- |
*/
```

### 資料再修正
```csharp
// 將 Name 轉為大寫
template.Columns.Add("Uppercase Name", x => x.Name, x => x.UseValue(y => y.Value.ToUpper()));

// 不指定欄位，選擇將 Record 的兩個資料合併
template.Columns.Add("Merge Data", x => x.UseValue(y => y.Record.Id + y.Record.Name));

/*
產出
| -------------- | ----------- | 
| Uppercase Name | Merge Data  |
| -------------- | ----------- |
| MARY           | 0Marry      |
| -------------- | ----------- |
| TERRY          | 1Terry      |
| -------------- | ----------- |
*/
```

### 設定儲存格樣式
```csharp
CellStyleConfiguration cellStyles = SpreadsheetManager.DefaultCellStyles;
// 標題背景色是紅色，Id 為 0 的資料背景色是藍色，其他是黃色
template.Columns.Add(
    "樣式", x => x.Id,
    cellStyles.HeaderStyle.CloneAndSetBackgroundColor(Color.Red),
    x => x.Value == 0
        ? cellStyles.FieldStyle.CloneAndSetBackgroundColor(Color.Blue)
        : cellStyles.FieldStyle.CloneAndSetBackgroundColor(Color.Yellow)
);
```

### 產生子標題列
```csharp
// Method Chaining
template.Columns.Add("Group1")
    .AddChildToLast("Index", x => x.Id)
    .AddChildToLast("Name", x => x.Name)
    .Add("Group2")
    .AddChildToLast("Child Group1");
DataColumnCollection<Record> lastChildColumns = template.Columns.Last().ChildColumns;
lastChildColumns.AddChildToLast("Index", x => x.Id)
    .AddChildToLast("Name", x => x.Name);
template.Columns.AddChildToLast("Child Group2");
lastChildColumns = template.Columns.Last().ChildColumns
    .AddChildToLast("Index", x => x.Id)
    .AddChildToLast("Name", x => x.Name);

/*
產出
| ----- | ---- | ----- | ---- | ----- | ---- |
|              |           Group2            |
|    Group1    | ------------ | ------------ |
| 		       | Child Group1 | Child Group2 |
| ----- | ---- | ----- | ---- | ----- | ---- |
| Index | Name | Index | Name | Index | Name |
| ----- | ---- | ----- | ---- | ----- | ---- |
*/
```

### 使用公式
```csharp
template.Columns.Add("Formula", x => x.UseFormula(y => $"{y.CellIndex} + {y.RowIndex}"));

/*
產出
第一列標題(Row Index = 0)，所以第一筆資料 Cell Index = 0, Row Index = 1
| ------- |
| Formula |
| ------  |
| =0 + 1  |
| ------- |
| =0 + 2  |
| ------- |
*/
```

## MergedTemplate
可合併多個 Template，自訂複雜的 Template 時，可用來在內部合併多個 Template 使 用。

## 自訂 Template
內部使用 GridTemplate 來建立可產生報表資訊列的 Template。
```csharp
public class ReportInfoTemplate : ITemplate {
    private readonly GridTemplate gridTemplate = new GridTemplate();

    public ReportInfoTemplate(string title, string user, int colSpan) {
        CellStyleConfiguration cellStyles = SpreadsheetManager.DefaultCellStyles;
        CellStyle titleStyle = cellStyles.CellStyle.CloneAndSetHorizontalAlignment(HorizontalAlignment.Center)
            .CloneAndSetFont(SpreadsheetManager.DefaultCellStyles.CellStyle.Font.CloneAndSetSize(14));

        gridTemplate.CreateRow();
        gridTemplate.CreateCell(title, 8, cellStyle: titleStyle);

        int leftColSpan = colSpan / 2;
        CellStyle rightCellStyle = cellStyles.CellStyle.CloneAndSetHorizontalAlignment(HorizontalAlignment.Left);
        CellStyle leftCellStyle = cellStyles.CellStyle.CloneAndSetHorizontalAlignment(HorizontalAlignment.Left);

        gridTemplate.CreateRow();
        gridTemplate.CreateCell("列印者：" + user, leftColSpan, cellStyle: rightCellStyle);
        gridTemplate.CreateCell("列印時間：" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"), colSpan - leftColSpan, cellStyle: leftCellStyle);
    }

    public TemplateContext GetContext() {
        return gridTemplate.GetContext();
    }
}
```

## 其他教學
* [入門教學](./%E5%85%A5%E9%96%80%E6%95%99%E5%AD%B8.md)
* [Exporter](./Exporter.md)
* [Sheeter](./Sheeter.md)
* [設定預設的儲存格樣式](./%E8%A8%AD%E5%AE%9A%E9%A0%90%E8%A8%AD%E7%9A%84%E5%84%B2%E5%AD%98%E6%A0%BC%E6%A8%A3%E5%BC%8F.md)
