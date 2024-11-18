#### [CloudyWing.SpreadsheetExporter](index.md 'index')
### [CloudyWing.SpreadsheetExporter](CloudyWing.SpreadsheetExporter.md 'CloudyWing.SpreadsheetExporter')

## ISpreadsheetExporter Interface

Interface for spreadsheet exporters.

```csharp
public interface ISpreadsheetExporter
```

Derived  
&#8627; [ExporterBase](CloudyWing.SpreadsheetExporter.ExporterBase.md 'CloudyWing.SpreadsheetExporter.ExporterBase')
### Properties

<a name='CloudyWing.SpreadsheetExporter.ISpreadsheetExporter.ContentType'></a>

## ISpreadsheetExporter.ContentType Property

Gets the content-type.

```csharp
string ContentType { get; }
```

#### Property Value
[System.String](https://docs.microsoft.com/en-us/dotnet/api/System.String 'System.String')  
The content-type.

<a name='CloudyWing.SpreadsheetExporter.ISpreadsheetExporter.DefaultBasicSheetName'></a>

## ISpreadsheetExporter.DefaultBasicSheetName Property

Gets or sets the default basic name of the sheet.

```csharp
string DefaultBasicSheetName { get; set; }
```

#### Property Value
[System.String](https://docs.microsoft.com/en-us/dotnet/api/System.String 'System.String')  
The default basic name of the sheet.

<a name='CloudyWing.SpreadsheetExporter.ISpreadsheetExporter.DefaultFont'></a>

## ISpreadsheetExporter.DefaultFont Property

Gets or sets the default font.

```csharp
System.Nullable<CloudyWing.SpreadsheetExporter.CellFont> DefaultFont { get; set; }
```

#### Property Value
[System.Nullable&lt;](https://docs.microsoft.com/en-us/dotnet/api/System.Nullable-1 'System.Nullable`1')[CellFont](CloudyWing.SpreadsheetExporter.CellFont.md 'CloudyWing.SpreadsheetExporter.CellFont')[&gt;](https://docs.microsoft.com/en-us/dotnet/api/System.Nullable-1 'System.Nullable`1')  
The default font.

<a name='CloudyWing.SpreadsheetExporter.ISpreadsheetExporter.FileNameExtension'></a>

## ISpreadsheetExporter.FileNameExtension Property

Gets the file name extension.

```csharp
string FileNameExtension { get; }
```

#### Property Value
[System.String](https://docs.microsoft.com/en-us/dotnet/api/System.String 'System.String')  
The file name extension.

<a name='CloudyWing.SpreadsheetExporter.ISpreadsheetExporter.HasPassword'></a>

## ISpreadsheetExporter.HasPassword Property

Gets a value indicating whether this instance has password.

```csharp
bool HasPassword { get; }
```

#### Property Value
[System.Boolean](https://docs.microsoft.com/en-us/dotnet/api/System.Boolean 'System.Boolean')  
`true` if this instance has password; otherwise, `false`.

<a name='CloudyWing.SpreadsheetExporter.ISpreadsheetExporter.LastSheeter'></a>

## ISpreadsheetExporter.LastSheeter Property

Gets the last sheeter. When there is no sheeter, create a sheeter.

```csharp
CloudyWing.SpreadsheetExporter.Sheeter LastSheeter { get; }
```

#### Property Value
[Sheeter](CloudyWing.SpreadsheetExporter.Sheeter.md 'CloudyWing.SpreadsheetExporter.Sheeter')  
The last sheeter.

<a name='CloudyWing.SpreadsheetExporter.ISpreadsheetExporter.Password'></a>

## ISpreadsheetExporter.Password Property

Gets or sets the workbook protection password.

```csharp
string Password { get; set; }
```

#### Property Value
[System.String](https://docs.microsoft.com/en-us/dotnet/api/System.String 'System.String')  
The password.
### Methods

<a name='CloudyWing.SpreadsheetExporter.ISpreadsheetExporter.CreateSheeter(string,System.Nullable_double_)'></a>

## ISpreadsheetExporter.CreateSheeter(string, Nullable<double>) Method

Creates the sheeter.

```csharp
CloudyWing.SpreadsheetExporter.Sheeter CreateSheeter(string sheetName=null, System.Nullable<double> defaultRowHeight=null);
```
#### Parameters

<a name='CloudyWing.SpreadsheetExporter.ISpreadsheetExporter.CreateSheeter(string,System.Nullable_double_).sheetName'></a>

`sheetName` [System.String](https://docs.microsoft.com/en-us/dotnet/api/System.String 'System.String')

Name of the sheet.

<a name='CloudyWing.SpreadsheetExporter.ISpreadsheetExporter.CreateSheeter(string,System.Nullable_double_).defaultRowHeight'></a>

`defaultRowHeight` [System.Nullable&lt;](https://docs.microsoft.com/en-us/dotnet/api/System.Nullable-1 'System.Nullable`1')[System.Double](https://docs.microsoft.com/en-us/dotnet/api/System.Double 'System.Double')[&gt;](https://docs.microsoft.com/en-us/dotnet/api/System.Nullable-1 'System.Nullable`1')

Default height of the row.

#### Returns
[Sheeter](CloudyWing.SpreadsheetExporter.Sheeter.md 'CloudyWing.SpreadsheetExporter.Sheeter')  
The sheeter.

<a name='CloudyWing.SpreadsheetExporter.ISpreadsheetExporter.Export()'></a>

## ISpreadsheetExporter.Export() Method

Exports bytes of spreadsheet.

```csharp
byte[] Export();
```

#### Returns
[System.Byte](https://docs.microsoft.com/en-us/dotnet/api/System.Byte 'System.Byte')[[]](https://docs.microsoft.com/en-us/dotnet/api/System.Array 'System.Array')  
The bytes of spreadsheet.

<a name='CloudyWing.SpreadsheetExporter.ISpreadsheetExporter.ExportFile(string,CloudyWing.SpreadsheetExporter.SpreadsheetFileMode)'></a>

## ISpreadsheetExporter.ExportFile(string, SpreadsheetFileMode) Method

Exports the spreadsheet file.

```csharp
void ExportFile(string path, CloudyWing.SpreadsheetExporter.SpreadsheetFileMode fileMode=CloudyWing.SpreadsheetExporter.SpreadsheetFileMode.Create);
```
#### Parameters

<a name='CloudyWing.SpreadsheetExporter.ISpreadsheetExporter.ExportFile(string,CloudyWing.SpreadsheetExporter.SpreadsheetFileMode).path'></a>

`path` [System.String](https://docs.microsoft.com/en-us/dotnet/api/System.String 'System.String')

The path.

<a name='CloudyWing.SpreadsheetExporter.ISpreadsheetExporter.ExportFile(string,CloudyWing.SpreadsheetExporter.SpreadsheetFileMode).fileMode'></a>

`fileMode` [SpreadsheetFileMode](CloudyWing.SpreadsheetExporter.SpreadsheetFileMode.md 'CloudyWing.SpreadsheetExporter.SpreadsheetFileMode')

The file mode.

<a name='CloudyWing.SpreadsheetExporter.ISpreadsheetExporter.GetSheeter(int)'></a>

## ISpreadsheetExporter.GetSheeter(int) Method

Gets the sheeter.

```csharp
CloudyWing.SpreadsheetExporter.Sheeter GetSheeter(int index);
```
#### Parameters

<a name='CloudyWing.SpreadsheetExporter.ISpreadsheetExporter.GetSheeter(int).index'></a>

`index` [System.Int32](https://docs.microsoft.com/en-us/dotnet/api/System.Int32 'System.Int32')

The index.

#### Returns
[Sheeter](CloudyWing.SpreadsheetExporter.Sheeter.md 'CloudyWing.SpreadsheetExporter.Sheeter')  
The sheeter.
### Events

<a name='CloudyWing.SpreadsheetExporter.ISpreadsheetExporter.SheetCreatedEvent'></a>

## ISpreadsheetExporter.SheetCreatedEvent Event

Occurs when [sheet created event].

```csharp
event EventHandler<SheetCreatedEventArgs> SheetCreatedEvent;
```

#### Event Type
[System.EventHandler&lt;](https://docs.microsoft.com/en-us/dotnet/api/System.EventHandler-1 'System.EventHandler`1')[SheetCreatedEventArgs](CloudyWing.SpreadsheetExporter.SheetCreatedEventArgs.md 'CloudyWing.SpreadsheetExporter.SheetCreatedEventArgs')[&gt;](https://docs.microsoft.com/en-us/dotnet/api/System.EventHandler-1 'System.EventHandler`1')

<a name='CloudyWing.SpreadsheetExporter.ISpreadsheetExporter.SpreadsheetExportedEvent'></a>

## ISpreadsheetExporter.SpreadsheetExportedEvent Event

Occurs when [spreadsheet exported event].

```csharp
event EventHandler<SpreadsheetExportedEventArgs> SpreadsheetExportedEvent;
```

#### Event Type
[System.EventHandler&lt;](https://docs.microsoft.com/en-us/dotnet/api/System.EventHandler-1 'System.EventHandler`1')[SpreadsheetExportedEventArgs](CloudyWing.SpreadsheetExporter.SpreadsheetExportedEventArgs.md 'CloudyWing.SpreadsheetExporter.SpreadsheetExportedEventArgs')[&gt;](https://docs.microsoft.com/en-us/dotnet/api/System.EventHandler-1 'System.EventHandler`1')

<a name='CloudyWing.SpreadsheetExporter.ISpreadsheetExporter.SpreadsheetExportingEvent'></a>

## ISpreadsheetExporter.SpreadsheetExportingEvent Event

Occurs when [spreadsheet exporting event].

```csharp
event EventHandler<SpreadsheetExportingEventArgs> SpreadsheetExportingEvent;
```

#### Event Type
[System.EventHandler&lt;](https://docs.microsoft.com/en-us/dotnet/api/System.EventHandler-1 'System.EventHandler`1')[SpreadsheetExportingEventArgs](CloudyWing.SpreadsheetExporter.SpreadsheetExportingEventArgs.md 'CloudyWing.SpreadsheetExporter.SpreadsheetExportingEventArgs')[&gt;](https://docs.microsoft.com/en-us/dotnet/api/System.EventHandler-1 'System.EventHandler`1')