#### [CloudyWing\.SpreadsheetExporter](index.md 'index')
### [CloudyWing\.SpreadsheetExporter](CloudyWing.SpreadsheetExporter.md 'CloudyWing\.SpreadsheetExporter')

## ExporterBase Class

The exporter base\.

```csharp
public abstract class ExporterBase : CloudyWing.SpreadsheetExporter.ISpreadsheetExporter
```

Inheritance [System\.Object](https://learn.microsoft.com/en-us/dotnet/api/system.object 'System\.Object') &#129106; ExporterBase

Implements [ISpreadsheetExporter](CloudyWing.SpreadsheetExporter.ISpreadsheetExporter.md 'CloudyWing\.SpreadsheetExporter\.ISpreadsheetExporter')
### Properties

<a name='CloudyWing.SpreadsheetExporter.ExporterBase.ContentType'></a>

## ExporterBase\.ContentType Property

Gets the content\-type\.

```csharp
public abstract string ContentType { get; }
```

Implements [ContentType](CloudyWing.SpreadsheetExporter.ISpreadsheetExporter.md#CloudyWing.SpreadsheetExporter.ISpreadsheetExporter.ContentType 'CloudyWing\.SpreadsheetExporter\.ISpreadsheetExporter\.ContentType')

#### Property Value
[System\.String](https://learn.microsoft.com/en-us/dotnet/api/system.string 'System\.String')  
The content\-type\.

<a name='CloudyWing.SpreadsheetExporter.ExporterBase.DefaultBasicSheetName'></a>

## ExporterBase\.DefaultBasicSheetName Property

Gets or sets the default basic name of the sheet\.

```csharp
public string DefaultBasicSheetName { get; set; }
```

Implements [DefaultBasicSheetName](CloudyWing.SpreadsheetExporter.ISpreadsheetExporter.md#CloudyWing.SpreadsheetExporter.ISpreadsheetExporter.DefaultBasicSheetName 'CloudyWing\.SpreadsheetExporter\.ISpreadsheetExporter\.DefaultBasicSheetName')

#### Property Value
[System\.String](https://learn.microsoft.com/en-us/dotnet/api/system.string 'System\.String')  
The default basic name of the sheet\.

<a name='CloudyWing.SpreadsheetExporter.ExporterBase.DefaultFont'></a>

## ExporterBase\.DefaultFont Property

Gets or sets the default font\.

```csharp
public System.Nullable<CloudyWing.SpreadsheetExporter.CellFont> DefaultFont { get; set; }
```

Implements [DefaultFont](CloudyWing.SpreadsheetExporter.ISpreadsheetExporter.md#CloudyWing.SpreadsheetExporter.ISpreadsheetExporter.DefaultFont 'CloudyWing\.SpreadsheetExporter\.ISpreadsheetExporter\.DefaultFont')

#### Property Value
[System\.Nullable&lt;](https://learn.microsoft.com/en-us/dotnet/api/system.nullable-1 'System\.Nullable\`1')[CellFont](CloudyWing.SpreadsheetExporter.CellFont.md 'CloudyWing\.SpreadsheetExporter\.CellFont')[&gt;](https://learn.microsoft.com/en-us/dotnet/api/system.nullable-1 'System\.Nullable\`1')  
The default font\.

<a name='CloudyWing.SpreadsheetExporter.ExporterBase.FileNameExtension'></a>

## ExporterBase\.FileNameExtension Property

Gets the file name extension\.

```csharp
public abstract string FileNameExtension { get; }
```

Implements [FileNameExtension](CloudyWing.SpreadsheetExporter.ISpreadsheetExporter.md#CloudyWing.SpreadsheetExporter.ISpreadsheetExporter.FileNameExtension 'CloudyWing\.SpreadsheetExporter\.ISpreadsheetExporter\.FileNameExtension')

#### Property Value
[System\.String](https://learn.microsoft.com/en-us/dotnet/api/system.string 'System\.String')  
The file name extension\.

<a name='CloudyWing.SpreadsheetExporter.ExporterBase.HasPassword'></a>

## ExporterBase\.HasPassword Property

Gets a value indicating whether this instance has password\.

```csharp
public bool HasPassword { get; }
```

Implements [HasPassword](CloudyWing.SpreadsheetExporter.ISpreadsheetExporter.md#CloudyWing.SpreadsheetExporter.ISpreadsheetExporter.HasPassword 'CloudyWing\.SpreadsheetExporter\.ISpreadsheetExporter\.HasPassword')

#### Property Value
[System\.Boolean](https://learn.microsoft.com/en-us/dotnet/api/system.boolean 'System\.Boolean')  
`true` if this instance has password; otherwise, `false`\.

<a name='CloudyWing.SpreadsheetExporter.ExporterBase.LastSheeter'></a>

## ExporterBase\.LastSheeter Property

Gets the last sheeter\. When there is no sheeter, create a sheeter\.

```csharp
public CloudyWing.SpreadsheetExporter.Sheeter LastSheeter { get; }
```

Implements [LastSheeter](CloudyWing.SpreadsheetExporter.ISpreadsheetExporter.md#CloudyWing.SpreadsheetExporter.ISpreadsheetExporter.LastSheeter 'CloudyWing\.SpreadsheetExporter\.ISpreadsheetExporter\.LastSheeter')

#### Property Value
[Sheeter](CloudyWing.SpreadsheetExporter.Sheeter.md 'CloudyWing\.SpreadsheetExporter\.Sheeter')  
The last sheeter\.

<a name='CloudyWing.SpreadsheetExporter.ExporterBase.Password'></a>

## ExporterBase\.Password Property

Gets or sets the workbook protection password\.

```csharp
public string Password { get; set; }
```

Implements [Password](CloudyWing.SpreadsheetExporter.ISpreadsheetExporter.md#CloudyWing.SpreadsheetExporter.ISpreadsheetExporter.Password 'CloudyWing\.SpreadsheetExporter\.ISpreadsheetExporter\.Password')

#### Property Value
[System\.String](https://learn.microsoft.com/en-us/dotnet/api/system.string 'System\.String')  
The password\.
### Methods

<a name='CloudyWing.SpreadsheetExporter.ExporterBase.CreateSheeter(string,System.Nullable_double_)'></a>

## ExporterBase\.CreateSheeter\(string, Nullable\<double\>\) Method

Creates the sheeter\.

```csharp
public CloudyWing.SpreadsheetExporter.Sheeter CreateSheeter(string sheetName=null, System.Nullable<double> defaultRowHeight=null);
```
#### Parameters

<a name='CloudyWing.SpreadsheetExporter.ExporterBase.CreateSheeter(string,System.Nullable_double_).sheetName'></a>

`sheetName` [System\.String](https://learn.microsoft.com/en-us/dotnet/api/system.string 'System\.String')

Name of the sheet\.

<a name='CloudyWing.SpreadsheetExporter.ExporterBase.CreateSheeter(string,System.Nullable_double_).defaultRowHeight'></a>

`defaultRowHeight` [System\.Nullable&lt;](https://learn.microsoft.com/en-us/dotnet/api/system.nullable-1 'System\.Nullable\`1')[System\.Double](https://learn.microsoft.com/en-us/dotnet/api/system.double 'System\.Double')[&gt;](https://learn.microsoft.com/en-us/dotnet/api/system.nullable-1 'System\.Nullable\`1')

Default height of the row\.

Implements [CreateSheeter\(string, Nullable&lt;double&gt;\)](CloudyWing.SpreadsheetExporter.ISpreadsheetExporter.md#CloudyWing.SpreadsheetExporter.ISpreadsheetExporter.CreateSheeter(string,System.Nullable_double_) 'CloudyWing\.SpreadsheetExporter\.ISpreadsheetExporter\.CreateSheeter\(string, System\.Nullable\<double\>\)')

#### Returns
[Sheeter](CloudyWing.SpreadsheetExporter.Sheeter.md 'CloudyWing\.SpreadsheetExporter\.Sheeter')  
The sheeter\.

<a name='CloudyWing.SpreadsheetExporter.ExporterBase.ExecuteExport(System.Collections.Generic.IEnumerable_CloudyWing.SpreadsheetExporter.SheeterContext_)'></a>

## ExporterBase\.ExecuteExport\(IEnumerable\<SheeterContext\>\) Method

Executes the export\.

```csharp
protected abstract byte[] ExecuteExport(System.Collections.Generic.IEnumerable<CloudyWing.SpreadsheetExporter.SheeterContext> contexts);
```
#### Parameters

<a name='CloudyWing.SpreadsheetExporter.ExporterBase.ExecuteExport(System.Collections.Generic.IEnumerable_CloudyWing.SpreadsheetExporter.SheeterContext_).contexts'></a>

`contexts` [System\.Collections\.Generic\.IEnumerable&lt;](https://learn.microsoft.com/en-us/dotnet/api/system.collections.generic.ienumerable-1 'System\.Collections\.Generic\.IEnumerable\`1')[SheeterContext](CloudyWing.SpreadsheetExporter.SheeterContext.md 'CloudyWing\.SpreadsheetExporter\.SheeterContext')[&gt;](https://learn.microsoft.com/en-us/dotnet/api/system.collections.generic.ienumerable-1 'System\.Collections\.Generic\.IEnumerable\`1')

The contexts\.

#### Returns
[System\.Byte](https://learn.microsoft.com/en-us/dotnet/api/system.byte 'System\.Byte')[\[\]](https://learn.microsoft.com/en-us/dotnet/api/system.array 'System\.Array')  
The exported bytes of the file\.

<a name='CloudyWing.SpreadsheetExporter.ExporterBase.Export()'></a>

## ExporterBase\.Export\(\) Method

Exports bytes of spreadsheet\.

```csharp
public byte[] Export();
```

Implements [Export\(\)](CloudyWing.SpreadsheetExporter.ISpreadsheetExporter.md#CloudyWing.SpreadsheetExporter.ISpreadsheetExporter.Export() 'CloudyWing\.SpreadsheetExporter\.ISpreadsheetExporter\.Export\(\)')

#### Returns
[System\.Byte](https://learn.microsoft.com/en-us/dotnet/api/system.byte 'System\.Byte')[\[\]](https://learn.microsoft.com/en-us/dotnet/api/system.array 'System\.Array')  
The bytes of spreadsheet\.

#### Exceptions

[SheeterNotFoundException](CloudyWing.SpreadsheetExporter.Exceptions.SheeterNotFoundException.md 'CloudyWing\.SpreadsheetExporter\.Exceptions\.SheeterNotFoundException')  
Sheeter\[0\] is null\.

<a name='CloudyWing.SpreadsheetExporter.ExporterBase.ExportFile(string,CloudyWing.SpreadsheetExporter.SpreadsheetFileMode)'></a>

## ExporterBase\.ExportFile\(string, SpreadsheetFileMode\) Method

Exports the spreadsheet file\.

```csharp
public void ExportFile(string path, CloudyWing.SpreadsheetExporter.SpreadsheetFileMode fileMode=CloudyWing.SpreadsheetExporter.SpreadsheetFileMode.Create);
```
#### Parameters

<a name='CloudyWing.SpreadsheetExporter.ExporterBase.ExportFile(string,CloudyWing.SpreadsheetExporter.SpreadsheetFileMode).path'></a>

`path` [System\.String](https://learn.microsoft.com/en-us/dotnet/api/system.string 'System\.String')

The path\.

<a name='CloudyWing.SpreadsheetExporter.ExporterBase.ExportFile(string,CloudyWing.SpreadsheetExporter.SpreadsheetFileMode).fileMode'></a>

`fileMode` [SpreadsheetFileMode](CloudyWing.SpreadsheetExporter.SpreadsheetFileMode.md 'CloudyWing\.SpreadsheetExporter\.SpreadsheetFileMode')

The file mode\.

Implements [ExportFile\(string, SpreadsheetFileMode\)](CloudyWing.SpreadsheetExporter.ISpreadsheetExporter.md#CloudyWing.SpreadsheetExporter.ISpreadsheetExporter.ExportFile(string,CloudyWing.SpreadsheetExporter.SpreadsheetFileMode) 'CloudyWing\.SpreadsheetExporter\.ISpreadsheetExporter\.ExportFile\(string, CloudyWing\.SpreadsheetExporter\.SpreadsheetFileMode\)')

#### Exceptions

[SheeterNotFoundException](CloudyWing.SpreadsheetExporter.Exceptions.SheeterNotFoundException.md 'CloudyWing\.SpreadsheetExporter\.Exceptions\.SheeterNotFoundException')  
No sheeters have been created\.

[System\.IO\.IOException](https://learn.microsoft.com/en-us/dotnet/api/system.io.ioexception 'System\.IO\.IOException')  
An I/O error occurred while creating the file\.

[System\.UnauthorizedAccessException](https://learn.microsoft.com/en-us/dotnet/api/system.unauthorizedaccessexception 'System\.UnauthorizedAccessException')  
The caller does not have the required permission\.

<a name='CloudyWing.SpreadsheetExporter.ExporterBase.GetSheeter(int)'></a>

## ExporterBase\.GetSheeter\(int\) Method

Gets the sheeter\.

```csharp
public CloudyWing.SpreadsheetExporter.Sheeter GetSheeter(int index);
```
#### Parameters

<a name='CloudyWing.SpreadsheetExporter.ExporterBase.GetSheeter(int).index'></a>

`index` [System\.Int32](https://learn.microsoft.com/en-us/dotnet/api/system.int32 'System\.Int32')

The index\.

Implements [GetSheeter\(int\)](CloudyWing.SpreadsheetExporter.ISpreadsheetExporter.md#CloudyWing.SpreadsheetExporter.ISpreadsheetExporter.GetSheeter(int) 'CloudyWing\.SpreadsheetExporter\.ISpreadsheetExporter\.GetSheeter\(int\)')

#### Returns
[Sheeter](CloudyWing.SpreadsheetExporter.Sheeter.md 'CloudyWing\.SpreadsheetExporter\.Sheeter')  
The sheeter\.

#### Exceptions

[System\.ArgumentOutOfRangeException](https://learn.microsoft.com/en-us/dotnet/api/system.argumentoutofrangeexception 'System\.ArgumentOutOfRangeException')  
[index](CloudyWing.SpreadsheetExporter.ExporterBase.md#CloudyWing.SpreadsheetExporter.ExporterBase.GetSheeter(int).index 'CloudyWing\.SpreadsheetExporter\.ExporterBase\.GetSheeter\(int\)\.index') is less than 0 or greater than or equal to the number of sheeters\.

<a name='CloudyWing.SpreadsheetExporter.ExporterBase.OnSheetCreated(CloudyWing.SpreadsheetExporter.SheetCreatedEventArgs)'></a>

## ExporterBase\.OnSheetCreated\(SheetCreatedEventArgs\) Method

Raises the [SheetCreated](https://learn.microsoft.com/en-us/dotnet/api/sheetcreated 'SheetCreated') event\.

```csharp
protected virtual void OnSheetCreated(CloudyWing.SpreadsheetExporter.SheetCreatedEventArgs args);
```
#### Parameters

<a name='CloudyWing.SpreadsheetExporter.ExporterBase.OnSheetCreated(CloudyWing.SpreadsheetExporter.SheetCreatedEventArgs).args'></a>

`args` [SheetCreatedEventArgs](CloudyWing.SpreadsheetExporter.SheetCreatedEventArgs.md 'CloudyWing\.SpreadsheetExporter\.SheetCreatedEventArgs')

The [SheetCreatedEventArgs](CloudyWing.SpreadsheetExporter.SheetCreatedEventArgs.md 'CloudyWing\.SpreadsheetExporter\.SheetCreatedEventArgs') instance containing the event data\.

<a name='CloudyWing.SpreadsheetExporter.ExporterBase.OnSpreadsheetExported(CloudyWing.SpreadsheetExporter.SpreadsheetExportedEventArgs)'></a>

## ExporterBase\.OnSpreadsheetExported\(SpreadsheetExportedEventArgs\) Method

Raises the [SpreadsheetExported](https://learn.microsoft.com/en-us/dotnet/api/spreadsheetexported 'SpreadsheetExported') event\.

```csharp
protected virtual void OnSpreadsheetExported(CloudyWing.SpreadsheetExporter.SpreadsheetExportedEventArgs args);
```
#### Parameters

<a name='CloudyWing.SpreadsheetExporter.ExporterBase.OnSpreadsheetExported(CloudyWing.SpreadsheetExporter.SpreadsheetExportedEventArgs).args'></a>

`args` [SpreadsheetExportedEventArgs](CloudyWing.SpreadsheetExporter.SpreadsheetExportedEventArgs.md 'CloudyWing\.SpreadsheetExporter\.SpreadsheetExportedEventArgs')

The [SpreadsheetExportedEventArgs](CloudyWing.SpreadsheetExporter.SpreadsheetExportedEventArgs.md 'CloudyWing\.SpreadsheetExporter\.SpreadsheetExportedEventArgs') instance containing the event data\.

<a name='CloudyWing.SpreadsheetExporter.ExporterBase.OnSpreadsheetExporting(CloudyWing.SpreadsheetExporter.SpreadsheetExportingEventArgs)'></a>

## ExporterBase\.OnSpreadsheetExporting\(SpreadsheetExportingEventArgs\) Method

Raises the [SpreadsheetExporting](https://learn.microsoft.com/en-us/dotnet/api/spreadsheetexporting 'SpreadsheetExporting') event\.

```csharp
protected virtual void OnSpreadsheetExporting(CloudyWing.SpreadsheetExporter.SpreadsheetExportingEventArgs args);
```
#### Parameters

<a name='CloudyWing.SpreadsheetExporter.ExporterBase.OnSpreadsheetExporting(CloudyWing.SpreadsheetExporter.SpreadsheetExportingEventArgs).args'></a>

`args` [SpreadsheetExportingEventArgs](CloudyWing.SpreadsheetExporter.SpreadsheetExportingEventArgs.md 'CloudyWing\.SpreadsheetExporter\.SpreadsheetExportingEventArgs')

The [SpreadsheetExportingEventArgs](CloudyWing.SpreadsheetExporter.SpreadsheetExportingEventArgs.md 'CloudyWing\.SpreadsheetExporter\.SpreadsheetExportingEventArgs') instance containing the event data\.

<a name='CloudyWing.SpreadsheetExporter.ExporterBase.TryGetSheeter(int,CloudyWing.SpreadsheetExporter.Sheeter)'></a>

## ExporterBase\.TryGetSheeter\(int, Sheeter\) Method

Tries to get the sheeter at the specified index\.

```csharp
public bool TryGetSheeter(int index, out CloudyWing.SpreadsheetExporter.Sheeter sheeter);
```
#### Parameters

<a name='CloudyWing.SpreadsheetExporter.ExporterBase.TryGetSheeter(int,CloudyWing.SpreadsheetExporter.Sheeter).index'></a>

`index` [System\.Int32](https://learn.microsoft.com/en-us/dotnet/api/system.int32 'System\.Int32')

The zero\-based index of the sheeter to get\.

<a name='CloudyWing.SpreadsheetExporter.ExporterBase.TryGetSheeter(int,CloudyWing.SpreadsheetExporter.Sheeter).sheeter'></a>

`sheeter` [Sheeter](CloudyWing.SpreadsheetExporter.Sheeter.md 'CloudyWing\.SpreadsheetExporter\.Sheeter')

When this method returns, contains the sheeter at the specified index if the index is valid;
otherwise, `null`\. This parameter is passed uninitialized\.

#### Returns
[System\.Boolean](https://learn.microsoft.com/en-us/dotnet/api/system.boolean 'System\.Boolean')  
`true` if the index is valid and the sheeter was retrieved; otherwise, `false`\.
### Events

<a name='CloudyWing.SpreadsheetExporter.ExporterBase.SheetCreatedEvent'></a>

## ExporterBase\.SheetCreatedEvent Event

Occurs when \[sheet created event\]\.

```csharp
public event EventHandler<SheetCreatedEventArgs> SheetCreatedEvent;
```

Implements [SheetCreatedEvent](CloudyWing.SpreadsheetExporter.ISpreadsheetExporter.md#CloudyWing.SpreadsheetExporter.ISpreadsheetExporter.SheetCreatedEvent 'CloudyWing\.SpreadsheetExporter\.ISpreadsheetExporter\.SheetCreatedEvent')

#### Event Type
[System\.EventHandler&lt;](https://learn.microsoft.com/en-us/dotnet/api/system.eventhandler-1 'System\.EventHandler\`1')[SheetCreatedEventArgs](CloudyWing.SpreadsheetExporter.SheetCreatedEventArgs.md 'CloudyWing\.SpreadsheetExporter\.SheetCreatedEventArgs')[&gt;](https://learn.microsoft.com/en-us/dotnet/api/system.eventhandler-1 'System\.EventHandler\`1')

<a name='CloudyWing.SpreadsheetExporter.ExporterBase.SpreadsheetExportedEvent'></a>

## ExporterBase\.SpreadsheetExportedEvent Event

Occurs when \[spreadsheet exported event\]\.

```csharp
public event EventHandler<SpreadsheetExportedEventArgs> SpreadsheetExportedEvent;
```

Implements [SpreadsheetExportedEvent](CloudyWing.SpreadsheetExporter.ISpreadsheetExporter.md#CloudyWing.SpreadsheetExporter.ISpreadsheetExporter.SpreadsheetExportedEvent 'CloudyWing\.SpreadsheetExporter\.ISpreadsheetExporter\.SpreadsheetExportedEvent')

#### Event Type
[System\.EventHandler&lt;](https://learn.microsoft.com/en-us/dotnet/api/system.eventhandler-1 'System\.EventHandler\`1')[SpreadsheetExportedEventArgs](CloudyWing.SpreadsheetExporter.SpreadsheetExportedEventArgs.md 'CloudyWing\.SpreadsheetExporter\.SpreadsheetExportedEventArgs')[&gt;](https://learn.microsoft.com/en-us/dotnet/api/system.eventhandler-1 'System\.EventHandler\`1')

<a name='CloudyWing.SpreadsheetExporter.ExporterBase.SpreadsheetExportingEvent'></a>

## ExporterBase\.SpreadsheetExportingEvent Event

Occurs when \[spreadsheet exporting event\]\.

```csharp
public event EventHandler<SpreadsheetExportingEventArgs> SpreadsheetExportingEvent;
```

Implements [SpreadsheetExportingEvent](CloudyWing.SpreadsheetExporter.ISpreadsheetExporter.md#CloudyWing.SpreadsheetExporter.ISpreadsheetExporter.SpreadsheetExportingEvent 'CloudyWing\.SpreadsheetExporter\.ISpreadsheetExporter\.SpreadsheetExportingEvent')

#### Event Type
[System\.EventHandler&lt;](https://learn.microsoft.com/en-us/dotnet/api/system.eventhandler-1 'System\.EventHandler\`1')[SpreadsheetExportingEventArgs](CloudyWing.SpreadsheetExporter.SpreadsheetExportingEventArgs.md 'CloudyWing\.SpreadsheetExporter\.SpreadsheetExportingEventArgs')[&gt;](https://learn.microsoft.com/en-us/dotnet/api/system.eventhandler-1 'System\.EventHandler\`1')