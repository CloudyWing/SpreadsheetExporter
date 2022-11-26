#### [CloudyWing.SpreadsheetExporter](index.md 'index')
### [CloudyWing.SpreadsheetExporter](CloudyWing.SpreadsheetExporter.md 'CloudyWing.SpreadsheetExporter')

## ExporterBase Class

The exporter base.

```csharp
public abstract class ExporterBase
```

Inheritance [System.Object](https://docs.microsoft.com/en-us/dotnet/api/System.Object 'System.Object') &#129106; ExporterBase
### Properties

<a name='CloudyWing.SpreadsheetExporter.ExporterBase.ContentType'></a>

## ExporterBase.ContentType Property

Gets the content-type.

```csharp
public abstract string ContentType { get; }
```

#### Property Value
[System.String](https://docs.microsoft.com/en-us/dotnet/api/System.String 'System.String')  
The content-type.

<a name='CloudyWing.SpreadsheetExporter.ExporterBase.DefaultBasicSheetName'></a>

## ExporterBase.DefaultBasicSheetName Property

Gets or sets the default basic name of the sheet.

```csharp
public string DefaultBasicSheetName { get; set; }
```

#### Property Value
[System.String](https://docs.microsoft.com/en-us/dotnet/api/System.String 'System.String')  
The default basic name of the sheet.

<a name='CloudyWing.SpreadsheetExporter.ExporterBase.FileNameExtension'></a>

## ExporterBase.FileNameExtension Property

Gets the file name extension.

```csharp
public abstract string FileNameExtension { get; }
```

#### Property Value
[System.String](https://docs.microsoft.com/en-us/dotnet/api/System.String 'System.String')  
The file name extension.

<a name='CloudyWing.SpreadsheetExporter.ExporterBase.HasPassword'></a>

## ExporterBase.HasPassword Property

Gets a value indicating whether this instance has password.

```csharp
public bool HasPassword { get; }
```

#### Property Value
[System.Boolean](https://docs.microsoft.com/en-us/dotnet/api/System.Boolean 'System.Boolean')  
`true` if this instance has password; otherwise, `false`.

<a name='CloudyWing.SpreadsheetExporter.ExporterBase.LastSheeter'></a>

## ExporterBase.LastSheeter Property

Gets the last sheeter. When there is no sheeter, create a sheeter.

```csharp
public CloudyWing.SpreadsheetExporter.Sheeter LastSheeter { get; }
```

#### Property Value
[Sheeter](CloudyWing.SpreadsheetExporter.Sheeter.md 'CloudyWing.SpreadsheetExporter.Sheeter')  
The last sheeter.

<a name='CloudyWing.SpreadsheetExporter.ExporterBase.Password'></a>

## ExporterBase.Password Property

Gets or sets the workbook protection password.

```csharp
public string Password { get; set; }
```

#### Property Value
[System.String](https://docs.microsoft.com/en-us/dotnet/api/System.String 'System.String')  
The password.
### Methods

<a name='CloudyWing.SpreadsheetExporter.ExporterBase.CreateSheeter(string)'></a>

## ExporterBase.CreateSheeter(string) Method

Creates the sheeter.

```csharp
public CloudyWing.SpreadsheetExporter.Sheeter CreateSheeter(string sheetName=null);
```
#### Parameters

<a name='CloudyWing.SpreadsheetExporter.ExporterBase.CreateSheeter(string).sheetName'></a>

`sheetName` [System.String](https://docs.microsoft.com/en-us/dotnet/api/System.String 'System.String')

Name of the sheet.

#### Returns
[Sheeter](CloudyWing.SpreadsheetExporter.Sheeter.md 'CloudyWing.SpreadsheetExporter.Sheeter')  
The sheeter.

<a name='CloudyWing.SpreadsheetExporter.ExporterBase.ExecuteExport(System.Collections.Generic.IEnumerable_CloudyWing.SpreadsheetExporter.SheeterContext_)'></a>

## ExporterBase.ExecuteExport(IEnumerable<SheeterContext>) Method

Executes the export.

```csharp
protected abstract byte[] ExecuteExport(System.Collections.Generic.IEnumerable<CloudyWing.SpreadsheetExporter.SheeterContext> contexts);
```
#### Parameters

<a name='CloudyWing.SpreadsheetExporter.ExporterBase.ExecuteExport(System.Collections.Generic.IEnumerable_CloudyWing.SpreadsheetExporter.SheeterContext_).contexts'></a>

`contexts` [System.Collections.Generic.IEnumerable&lt;](https://docs.microsoft.com/en-us/dotnet/api/System.Collections.Generic.IEnumerable-1 'System.Collections.Generic.IEnumerable`1')[SheeterContext](CloudyWing.SpreadsheetExporter.SheeterContext.md 'CloudyWing.SpreadsheetExporter.SheeterContext')[&gt;](https://docs.microsoft.com/en-us/dotnet/api/System.Collections.Generic.IEnumerable-1 'System.Collections.Generic.IEnumerable`1')

The contexts.

#### Returns
[System.Byte](https://docs.microsoft.com/en-us/dotnet/api/System.Byte 'System.Byte')[[]](https://docs.microsoft.com/en-us/dotnet/api/System.Array 'System.Array')  
The exported bytes of the file.

<a name='CloudyWing.SpreadsheetExporter.ExporterBase.Export()'></a>

## ExporterBase.Export() Method

Exports bytes of spreadsheet.

```csharp
public byte[] Export();
```

#### Returns
[System.Byte](https://docs.microsoft.com/en-us/dotnet/api/System.Byte 'System.Byte')[[]](https://docs.microsoft.com/en-us/dotnet/api/System.Array 'System.Array')  
The bytes of spreadsheet.

#### Exceptions

[SheeterNotFoundException](CloudyWing.SpreadsheetExporter.Exceptions.SheeterNotFoundException.md 'CloudyWing.SpreadsheetExporter.Exceptions.SheeterNotFoundException')  
Sheeter[0] is null.

<a name='CloudyWing.SpreadsheetExporter.ExporterBase.ExportFile(string,CloudyWing.SpreadsheetExporter.SpreadsheetFileMode)'></a>

## ExporterBase.ExportFile(string, SpreadsheetFileMode) Method

Exports the spreadsheet file.

```csharp
public void ExportFile(string path, CloudyWing.SpreadsheetExporter.SpreadsheetFileMode fileMode=CloudyWing.SpreadsheetExporter.SpreadsheetFileMode.Create);
```
#### Parameters

<a name='CloudyWing.SpreadsheetExporter.ExporterBase.ExportFile(string,CloudyWing.SpreadsheetExporter.SpreadsheetFileMode).path'></a>

`path` [System.String](https://docs.microsoft.com/en-us/dotnet/api/System.String 'System.String')

The path.

<a name='CloudyWing.SpreadsheetExporter.ExporterBase.ExportFile(string,CloudyWing.SpreadsheetExporter.SpreadsheetFileMode).fileMode'></a>

`fileMode` [SpreadsheetFileMode](CloudyWing.SpreadsheetExporter.SpreadsheetFileMode.md 'CloudyWing.SpreadsheetExporter.SpreadsheetFileMode')

The file mode.

<a name='CloudyWing.SpreadsheetExporter.ExporterBase.GetSheeter(int)'></a>

## ExporterBase.GetSheeter(int) Method

Gets the sheeter.

```csharp
public CloudyWing.SpreadsheetExporter.Sheeter GetSheeter(int index);
```
#### Parameters

<a name='CloudyWing.SpreadsheetExporter.ExporterBase.GetSheeter(int).index'></a>

`index` [System.Int32](https://docs.microsoft.com/en-us/dotnet/api/System.Int32 'System.Int32')

The index.

#### Returns
[Sheeter](CloudyWing.SpreadsheetExporter.Sheeter.md 'CloudyWing.SpreadsheetExporter.Sheeter')  
The sheeter.

<a name='CloudyWing.SpreadsheetExporter.ExporterBase.OnSpreadsheetExported(CloudyWing.SpreadsheetExporter.SpreadsheetExportedEventArgs)'></a>

## ExporterBase.OnSpreadsheetExported(SpreadsheetExportedEventArgs) Method

Raises the [SpreadsheetExported](https://docs.microsoft.com/en-us/dotnet/api/SpreadsheetExported 'SpreadsheetExported') event.

```csharp
protected virtual void OnSpreadsheetExported(CloudyWing.SpreadsheetExporter.SpreadsheetExportedEventArgs args);
```
#### Parameters

<a name='CloudyWing.SpreadsheetExporter.ExporterBase.OnSpreadsheetExported(CloudyWing.SpreadsheetExporter.SpreadsheetExportedEventArgs).args'></a>

`args` [SpreadsheetExportedEventArgs](CloudyWing.SpreadsheetExporter.SpreadsheetExportedEventArgs.md 'CloudyWing.SpreadsheetExporter.SpreadsheetExportedEventArgs')

The [SpreadsheetExportedEventArgs](CloudyWing.SpreadsheetExporter.SpreadsheetExportedEventArgs.md 'CloudyWing.SpreadsheetExporter.SpreadsheetExportedEventArgs') instance containing the event data.

<a name='CloudyWing.SpreadsheetExporter.ExporterBase.OnSpreadsheetExporting(CloudyWing.SpreadsheetExporter.SpreadsheetExportingEventArgs)'></a>

## ExporterBase.OnSpreadsheetExporting(SpreadsheetExportingEventArgs) Method

Raises the [SpreadsheetExporting](https://docs.microsoft.com/en-us/dotnet/api/SpreadsheetExporting 'SpreadsheetExporting') event.

```csharp
protected virtual void OnSpreadsheetExporting(CloudyWing.SpreadsheetExporter.SpreadsheetExportingEventArgs args);
```
#### Parameters

<a name='CloudyWing.SpreadsheetExporter.ExporterBase.OnSpreadsheetExporting(CloudyWing.SpreadsheetExporter.SpreadsheetExportingEventArgs).args'></a>

`args` [SpreadsheetExportingEventArgs](CloudyWing.SpreadsheetExporter.SpreadsheetExportingEventArgs.md 'CloudyWing.SpreadsheetExporter.SpreadsheetExportingEventArgs')

The [SpreadsheetExportingEventArgs](CloudyWing.SpreadsheetExporter.SpreadsheetExportingEventArgs.md 'CloudyWing.SpreadsheetExporter.SpreadsheetExportingEventArgs') instance containing the event data.
### Events

<a name='CloudyWing.SpreadsheetExporter.ExporterBase.SpreadsheetExportedEvent'></a>

## ExporterBase.SpreadsheetExportedEvent Event

Occurs when [spreadsheet exported event].

```csharp
public event EventHandler<SpreadsheetExportedEventArgs> SpreadsheetExportedEvent;
```

#### Event Type
[System.EventHandler&lt;](https://docs.microsoft.com/en-us/dotnet/api/System.EventHandler-1 'System.EventHandler`1')[SpreadsheetExportedEventArgs](CloudyWing.SpreadsheetExporter.SpreadsheetExportedEventArgs.md 'CloudyWing.SpreadsheetExporter.SpreadsheetExportedEventArgs')[&gt;](https://docs.microsoft.com/en-us/dotnet/api/System.EventHandler-1 'System.EventHandler`1')

<a name='CloudyWing.SpreadsheetExporter.ExporterBase.SpreadsheetExportingEvent'></a>

## ExporterBase.SpreadsheetExportingEvent Event

Occurs when [spreadsheet exporting event].

```csharp
public event EventHandler<SpreadsheetExportingEventArgs> SpreadsheetExportingEvent;
```

#### Event Type
[System.EventHandler&lt;](https://docs.microsoft.com/en-us/dotnet/api/System.EventHandler-1 'System.EventHandler`1')[SpreadsheetExportingEventArgs](CloudyWing.SpreadsheetExporter.SpreadsheetExportingEventArgs.md 'CloudyWing.SpreadsheetExporter.SpreadsheetExportingEventArgs')[&gt;](https://docs.microsoft.com/en-us/dotnet/api/System.EventHandler-1 'System.EventHandler`1')