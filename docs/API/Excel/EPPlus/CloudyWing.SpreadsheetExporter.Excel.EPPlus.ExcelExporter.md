### [CloudyWing.SpreadsheetExporter.Excel.EPPlus](CloudyWing.SpreadsheetExporter.Excel.EPPlus.md 'CloudyWing.SpreadsheetExporter.Excel.EPPlus')

## ExcelExporter Class

The excel exporter, using epplus.

```csharp
public class ExcelExporter : CloudyWing.SpreadsheetExporter.ExporterBase
```

Inheritance [System.Object](https://docs.microsoft.com/en-us/dotnet/api/System.Object 'System.Object') &#129106; [CloudyWing.SpreadsheetExporter.ExporterBase](https://docs.microsoft.com/en-us/dotnet/api/CloudyWing.SpreadsheetExporter.ExporterBase 'CloudyWing.SpreadsheetExporter.ExporterBase') &#129106; ExcelExporter

### See Also
- [CloudyWing.SpreadsheetExporter.ExporterBase](https://docs.microsoft.com/en-us/dotnet/api/CloudyWing.SpreadsheetExporter.ExporterBase 'CloudyWing.SpreadsheetExporter.ExporterBase')
### Properties

<a name='CloudyWing.SpreadsheetExporter.Excel.EPPlus.ExcelExporter.ContentType'></a>

## ExcelExporter.ContentType Property

Gets the content-type.

```csharp
public override string ContentType { get; }
```

Implements [ContentType](https://docs.microsoft.com/en-us/dotnet/api/CloudyWing.SpreadsheetExporter.ISpreadsheetExporter.ContentType 'CloudyWing.SpreadsheetExporter.ISpreadsheetExporter.ContentType')

#### Property Value
[System.String](https://docs.microsoft.com/en-us/dotnet/api/System.String 'System.String')  
The content-type.

<a name='CloudyWing.SpreadsheetExporter.Excel.EPPlus.ExcelExporter.FileNameExtension'></a>

## ExcelExporter.FileNameExtension Property

Gets the file name extension.

```csharp
public override string FileNameExtension { get; }
```

Implements [FileNameExtension](https://docs.microsoft.com/en-us/dotnet/api/CloudyWing.SpreadsheetExporter.ISpreadsheetExporter.FileNameExtension 'CloudyWing.SpreadsheetExporter.ISpreadsheetExporter.FileNameExtension')

#### Property Value
[System.String](https://docs.microsoft.com/en-us/dotnet/api/System.String 'System.String')  
The file name extension.
### Methods

<a name='CloudyWing.SpreadsheetExporter.Excel.EPPlus.ExcelExporter.ExecuteExport(System.Collections.Generic.IEnumerable_CloudyWing.SpreadsheetExporter.SheeterContext_)'></a>

## ExcelExporter.ExecuteExport(IEnumerable<SheeterContext>) Method

Executes the export.

```csharp
protected override byte[] ExecuteExport(System.Collections.Generic.IEnumerable<CloudyWing.SpreadsheetExporter.SheeterContext> contexts);
```
#### Parameters

<a name='CloudyWing.SpreadsheetExporter.Excel.EPPlus.ExcelExporter.ExecuteExport(System.Collections.Generic.IEnumerable_CloudyWing.SpreadsheetExporter.SheeterContext_).contexts'></a>

`contexts` [System.Collections.Generic.IEnumerable&lt;](https://docs.microsoft.com/en-us/dotnet/api/System.Collections.Generic.IEnumerable-1 'System.Collections.Generic.IEnumerable`1')[CloudyWing.SpreadsheetExporter.SheeterContext](https://docs.microsoft.com/en-us/dotnet/api/CloudyWing.SpreadsheetExporter.SheeterContext 'CloudyWing.SpreadsheetExporter.SheeterContext')[&gt;](https://docs.microsoft.com/en-us/dotnet/api/System.Collections.Generic.IEnumerable-1 'System.Collections.Generic.IEnumerable`1')

The contexts.

#### Returns
[System.Byte](https://docs.microsoft.com/en-us/dotnet/api/System.Byte 'System.Byte')[[]](https://docs.microsoft.com/en-us/dotnet/api/System.Array 'System.Array')  
The exported bytes of the file.