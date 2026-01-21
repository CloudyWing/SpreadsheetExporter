#### [CloudyWing\.SpreadsheetExporter](index.md 'index')
### [CloudyWing\.SpreadsheetExporter\.Templates\.DataTable](CloudyWing.SpreadsheetExporter.Templates.DataTable.md 'CloudyWing\.SpreadsheetExporter\.Templates\.DataTable')

## DataTableColumn Class

The DataTable column\.

```csharp
public class DataTableColumn
```

Inheritance [System\.Object](https://learn.microsoft.com/en-us/dotnet/api/system.object 'System\.Object') &#129106; DataTableColumn
### Properties

<a name='CloudyWing.SpreadsheetExporter.Templates.DataTable.DataTableColumn.ColumnName'></a>

## DataTableColumn\.ColumnName Property

Gets or sets the name of the column in the DataTable\.

```csharp
public string ColumnName { get; set; }
```

#### Property Value
[System\.String](https://learn.microsoft.com/en-us/dotnet/api/system.string 'System\.String')  
The name of the column\.

<a name='CloudyWing.SpreadsheetExporter.Templates.DataTable.DataTableColumn.FieldFormulaGenerator'></a>

## DataTableColumn\.FieldFormulaGenerator Property

Gets or sets the field formula generator\.

```csharp
public System.Func<object,string> FieldFormulaGenerator { get; set; }
```

#### Property Value
[System\.Func&lt;](https://learn.microsoft.com/en-us/dotnet/api/system.func-2 'System\.Func\`2')[System\.Object](https://learn.microsoft.com/en-us/dotnet/api/system.object 'System\.Object')[,](https://learn.microsoft.com/en-us/dotnet/api/system.func-2 'System\.Func\`2')[System\.String](https://learn.microsoft.com/en-us/dotnet/api/system.string 'System\.String')[&gt;](https://learn.microsoft.com/en-us/dotnet/api/system.func-2 'System\.Func\`2')  
The field formula generator\.

<a name='CloudyWing.SpreadsheetExporter.Templates.DataTable.DataTableColumn.FieldStyleGenerator'></a>

## DataTableColumn\.FieldStyleGenerator Property

Gets or sets the field style generator\.

```csharp
public System.Func<object,CloudyWing.SpreadsheetExporter.CellStyle> FieldStyleGenerator { get; set; }
```

#### Property Value
[System\.Func&lt;](https://learn.microsoft.com/en-us/dotnet/api/system.func-2 'System\.Func\`2')[System\.Object](https://learn.microsoft.com/en-us/dotnet/api/system.object 'System\.Object')[,](https://learn.microsoft.com/en-us/dotnet/api/system.func-2 'System\.Func\`2')[CellStyle](CloudyWing.SpreadsheetExporter.CellStyle.md 'CloudyWing\.SpreadsheetExporter\.CellStyle')[&gt;](https://learn.microsoft.com/en-us/dotnet/api/system.func-2 'System\.Func\`2')  
The field style generator\.

<a name='CloudyWing.SpreadsheetExporter.Templates.DataTable.DataTableColumn.FieldValueGenerator'></a>

## DataTableColumn\.FieldValueGenerator Property

Gets or sets the field value generator\.

```csharp
public System.Func<object,object> FieldValueGenerator { get; set; }
```

#### Property Value
[System\.Func&lt;](https://learn.microsoft.com/en-us/dotnet/api/system.func-2 'System\.Func\`2')[System\.Object](https://learn.microsoft.com/en-us/dotnet/api/system.object 'System\.Object')[,](https://learn.microsoft.com/en-us/dotnet/api/system.func-2 'System\.Func\`2')[System\.Object](https://learn.microsoft.com/en-us/dotnet/api/system.object 'System\.Object')[&gt;](https://learn.microsoft.com/en-us/dotnet/api/system.func-2 'System\.Func\`2')  
The field value generator\.

<a name='CloudyWing.SpreadsheetExporter.Templates.DataTable.DataTableColumn.HeaderStyle'></a>

## DataTableColumn\.HeaderStyle Property

Gets or sets the header style\.

```csharp
public CloudyWing.SpreadsheetExporter.CellStyle HeaderStyle { get; set; }
```

#### Property Value
[CellStyle](CloudyWing.SpreadsheetExporter.CellStyle.md 'CloudyWing\.SpreadsheetExporter\.CellStyle')  
The header style\.

<a name='CloudyWing.SpreadsheetExporter.Templates.DataTable.DataTableColumn.HeaderText'></a>

## DataTableColumn\.HeaderText Property

Gets or sets the header text\.

```csharp
public string HeaderText { get; set; }
```

#### Property Value
[System\.String](https://learn.microsoft.com/en-us/dotnet/api/system.string 'System\.String')  
The header text\.