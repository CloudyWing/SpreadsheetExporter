### [CloudyWing.SpreadsheetExporter.Excel.NPOI](CloudyWing.SpreadsheetExporter.Excel.NPOI.md 'CloudyWing.SpreadsheetExporter.Excel.NPOI')

## ExcelExporter Class

The excel exporter, using npoi.

```csharp
public sealed class ExcelExporter : CloudyWing.SpreadsheetExporter.ExporterBase
```

Inheritance [System.Object](https://docs.microsoft.com/en-us/dotnet/api/System.Object 'System.Object') &#129106; [CloudyWing.SpreadsheetExporter.ExporterBase](https://docs.microsoft.com/en-us/dotnet/api/CloudyWing.SpreadsheetExporter.ExporterBase 'CloudyWing.SpreadsheetExporter.ExporterBase') &#129106; ExcelExporter

### See Also
- [CloudyWing.SpreadsheetExporter.ExporterBase](https://docs.microsoft.com/en-us/dotnet/api/CloudyWing.SpreadsheetExporter.ExporterBase 'CloudyWing.SpreadsheetExporter.ExporterBase')
### Constructors

<a name='CloudyWing.SpreadsheetExporter.Excel.NPOI.ExcelExporter.ExcelExporter(CloudyWing.SpreadsheetExporter.Excel.NPOI.ExcelFormat)'></a>

## ExcelExporter(ExcelFormat) Constructor

The excel exporter, using npoi.

```csharp
public ExcelExporter(CloudyWing.SpreadsheetExporter.Excel.NPOI.ExcelFormat excelFormat=CloudyWing.SpreadsheetExporter.Excel.NPOI.ExcelFormat.OfficeOpenXmlDocument);
```
#### Parameters

<a name='CloudyWing.SpreadsheetExporter.Excel.NPOI.ExcelExporter.ExcelExporter(CloudyWing.SpreadsheetExporter.Excel.NPOI.ExcelFormat).excelFormat'></a>

`excelFormat` [ExcelFormat](CloudyWing.SpreadsheetExporter.Excel.NPOI.ExcelFormat.md 'CloudyWing.SpreadsheetExporter.Excel.NPOI.ExcelFormat')

The excel format.

### See Also
- [CloudyWing.SpreadsheetExporter.ExporterBase](https://docs.microsoft.com/en-us/dotnet/api/CloudyWing.SpreadsheetExporter.ExporterBase 'CloudyWing.SpreadsheetExporter.ExporterBase')
### Properties

<a name='CloudyWing.SpreadsheetExporter.Excel.NPOI.ExcelExporter.ContentType'></a>

## ExcelExporter.ContentType Property

Gets the content-type.

```csharp
public override string ContentType { get; }
```

#### Property Value
[System.String](https://docs.microsoft.com/en-us/dotnet/api/System.String 'System.String')  
The content-type.

<a name='CloudyWing.SpreadsheetExporter.Excel.NPOI.ExcelExporter.ExcelFormat'></a>

## ExcelExporter.ExcelFormat Property

Gets or sets the excel format.

```csharp
public CloudyWing.SpreadsheetExporter.Excel.NPOI.ExcelFormat ExcelFormat { get; set; }
```

#### Property Value
[ExcelFormat](CloudyWing.SpreadsheetExporter.Excel.NPOI.ExcelFormat.md 'CloudyWing.SpreadsheetExporter.Excel.NPOI.ExcelFormat')  
The excel format.

<a name='CloudyWing.SpreadsheetExporter.Excel.NPOI.ExcelExporter.FileNameExtension'></a>

## ExcelExporter.FileNameExtension Property

Gets the file name extension.

```csharp
public override string FileNameExtension { get; }
```

#### Property Value
[System.String](https://docs.microsoft.com/en-us/dotnet/api/System.String 'System.String')  
The file name extension.

<a name='CloudyWing.SpreadsheetExporter.Excel.NPOI.ExcelExporter.IsClosedNotImplementedException'></a>

## ExcelExporter.IsClosedNotImplementedException Property

Gets or sets a value indicating whether this instance is closed not implemented exception.

```csharp
public bool IsClosedNotImplementedException { get; set; }
```

#### Property Value
[System.Boolean](https://docs.microsoft.com/en-us/dotnet/api/System.Boolean 'System.Boolean')  
`true` if this instance is closed not implemented exception; otherwise, `false`.

<a name='CloudyWing.SpreadsheetExporter.Excel.NPOI.ExcelExporter.IsOfficeOpenXmlDocument'></a>

## ExcelExporter.IsOfficeOpenXmlDocument Property

Gets a value indicating whether this instance is office open XML document.

```csharp
public bool IsOfficeOpenXmlDocument { get; }
```

#### Property Value
[System.Boolean](https://docs.microsoft.com/en-us/dotnet/api/System.Boolean 'System.Boolean')  
`true` if this instance is office open XML document; otherwise, `false`.
### Methods

<a name='CloudyWing.SpreadsheetExporter.Excel.NPOI.ExcelExporter.ExecuteExport(System.Collections.Generic.IEnumerable_CloudyWing.SpreadsheetExporter.SheeterContext_)'></a>

## ExcelExporter.ExecuteExport(IEnumerable<SheeterContext>) Method

Executes the export.

```csharp
protected override byte[] ExecuteExport(System.Collections.Generic.IEnumerable<CloudyWing.SpreadsheetExporter.SheeterContext> contexts);
```
#### Parameters

<a name='CloudyWing.SpreadsheetExporter.Excel.NPOI.ExcelExporter.ExecuteExport(System.Collections.Generic.IEnumerable_CloudyWing.SpreadsheetExporter.SheeterContext_).contexts'></a>

`contexts` [System.Collections.Generic.IEnumerable&lt;](https://docs.microsoft.com/en-us/dotnet/api/System.Collections.Generic.IEnumerable-1 'System.Collections.Generic.IEnumerable`1')[CloudyWing.SpreadsheetExporter.SheeterContext](https://docs.microsoft.com/en-us/dotnet/api/CloudyWing.SpreadsheetExporter.SheeterContext 'CloudyWing.SpreadsheetExporter.SheeterContext')[&gt;](https://docs.microsoft.com/en-us/dotnet/api/System.Collections.Generic.IEnumerable-1 'System.Collections.Generic.IEnumerable`1')

The contexts.

#### Returns
[System.Byte](https://docs.microsoft.com/en-us/dotnet/api/System.Byte 'System.Byte')[[]](https://docs.microsoft.com/en-us/dotnet/api/System.Array 'System.Array')  
The exported bytes of the file.