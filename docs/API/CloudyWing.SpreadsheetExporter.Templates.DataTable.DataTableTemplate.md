#### [CloudyWing\.SpreadsheetExporter](index.md 'index')
### [CloudyWing\.SpreadsheetExporter\.Templates\.DataTable](CloudyWing.SpreadsheetExporter.Templates.DataTable.md 'CloudyWing\.SpreadsheetExporter\.Templates\.DataTable')

## DataTableTemplate Class

The DataTable template\. Create cell information using DataTable as data source\.

```csharp
public class DataTableTemplate : CloudyWing.SpreadsheetExporter.Templates.ITemplate
```

Inheritance [System\.Object](https://learn.microsoft.com/en-us/dotnet/api/system.object 'System\.Object') &#129106; DataTableTemplate

Implements [ITemplate](CloudyWing.SpreadsheetExporter.Templates.ITemplate.md 'CloudyWing\.SpreadsheetExporter\.Templates\.ITemplate')

#### Exceptions

[System\.ArgumentNullException](https://learn.microsoft.com/en-us/dotnet/api/system.argumentnullexception 'System\.ArgumentNullException')  
dataTable

### See Also
- [ITemplate](CloudyWing.SpreadsheetExporter.Templates.ITemplate.md 'CloudyWing\.SpreadsheetExporter\.Templates\.ITemplate')
### Constructors

<a name='CloudyWing.SpreadsheetExporter.Templates.DataTable.DataTableTemplate.DataTableTemplate(System.Data.DataTable)'></a>

## DataTableTemplate\(DataTable\) Constructor

The DataTable template\. Create cell information using DataTable as data source\.

```csharp
public DataTableTemplate(System.Data.DataTable dataTable);
```
#### Parameters

<a name='CloudyWing.SpreadsheetExporter.Templates.DataTable.DataTableTemplate.DataTableTemplate(System.Data.DataTable).dataTable'></a>

`dataTable` [System\.Data\.DataTable](https://learn.microsoft.com/en-us/dotnet/api/system.data.datatable 'System\.Data\.DataTable')

The data table\.

#### Exceptions

[System\.ArgumentNullException](https://learn.microsoft.com/en-us/dotnet/api/system.argumentnullexception 'System\.ArgumentNullException')  
dataTable

### See Also
- [ITemplate](CloudyWing.SpreadsheetExporter.Templates.ITemplate.md 'CloudyWing\.SpreadsheetExporter\.Templates\.ITemplate')
### Properties

<a name='CloudyWing.SpreadsheetExporter.Templates.DataTable.DataTableTemplate.Columns'></a>

## DataTableTemplate\.Columns Property

Gets the columns\.

```csharp
public CloudyWing.SpreadsheetExporter.Templates.DataTable.DataTableColumnCollection Columns { get; }
```

#### Property Value
[DataTableColumnCollection](CloudyWing.SpreadsheetExporter.Templates.DataTable.DataTableColumnCollection.md 'CloudyWing\.SpreadsheetExporter\.Templates\.DataTable\.DataTableColumnCollection')  
The columns\.

<a name='CloudyWing.SpreadsheetExporter.Templates.DataTable.DataTableTemplate.ColumnSpan'></a>

## DataTableTemplate\.ColumnSpan Property

Gets the column span\.

```csharp
public int ColumnSpan { get; }
```

#### Property Value
[System\.Int32](https://learn.microsoft.com/en-us/dotnet/api/system.int32 'System\.Int32')  
The column span\.

<a name='CloudyWing.SpreadsheetExporter.Templates.DataTable.DataTableTemplate.DataTable'></a>

## DataTableTemplate\.DataTable Property

Gets or sets the data table\.

```csharp
public System.Data.DataTable DataTable { get; set; }
```

#### Property Value
[System\.Data\.DataTable](https://learn.microsoft.com/en-us/dotnet/api/system.data.datatable 'System\.Data\.DataTable')  
The data table\.

<a name='CloudyWing.SpreadsheetExporter.Templates.DataTable.DataTableTemplate.HeaderHeight'></a>

## DataTableTemplate\.HeaderHeight Property

Gets or sets the height of the header\.

```csharp
public System.Nullable<double> HeaderHeight { get; set; }
```

#### Property Value
[System\.Nullable&lt;](https://learn.microsoft.com/en-us/dotnet/api/system.nullable-1 'System\.Nullable\`1')[System\.Double](https://learn.microsoft.com/en-us/dotnet/api/system.double 'System\.Double')[&gt;](https://learn.microsoft.com/en-us/dotnet/api/system.nullable-1 'System\.Nullable\`1')  
The height of the header\.

<a name='CloudyWing.SpreadsheetExporter.Templates.DataTable.DataTableTemplate.RecordHeight'></a>

## DataTableTemplate\.RecordHeight Property

Gets or sets the height of the record\.

```csharp
public System.Nullable<double> RecordHeight { get; set; }
```

#### Property Value
[System\.Nullable&lt;](https://learn.microsoft.com/en-us/dotnet/api/system.nullable-1 'System\.Nullable\`1')[System\.Double](https://learn.microsoft.com/en-us/dotnet/api/system.double 'System\.Double')[&gt;](https://learn.microsoft.com/en-us/dotnet/api/system.nullable-1 'System\.Nullable\`1')  
The height of the record\.

<a name='CloudyWing.SpreadsheetExporter.Templates.DataTable.DataTableTemplate.RowSpan'></a>

## DataTableTemplate\.RowSpan Property

Gets the row span\.

```csharp
public int RowSpan { get; }
```

#### Property Value
[System\.Int32](https://learn.microsoft.com/en-us/dotnet/api/system.int32 'System\.Int32')  
The row span\.
### Methods

<a name='CloudyWing.SpreadsheetExporter.Templates.DataTable.DataTableTemplate.GetContext()'></a>

## DataTableTemplate\.GetContext\(\) Method

Gets the context\.

```csharp
public CloudyWing.SpreadsheetExporter.Templates.TemplateContext GetContext();
```

Implements [GetContext\(\)](CloudyWing.SpreadsheetExporter.Templates.ITemplate.md#CloudyWing.SpreadsheetExporter.Templates.ITemplate.GetContext() 'CloudyWing\.SpreadsheetExporter\.Templates\.ITemplate\.GetContext\(\)')

#### Returns
[TemplateContext](CloudyWing.SpreadsheetExporter.Templates.TemplateContext.md 'CloudyWing\.SpreadsheetExporter\.Templates\.TemplateContext')  
The template context\.