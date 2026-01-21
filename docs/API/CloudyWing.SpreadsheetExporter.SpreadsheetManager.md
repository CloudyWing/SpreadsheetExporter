#### [CloudyWing\.SpreadsheetExporter](index.md 'index')
### [CloudyWing\.SpreadsheetExporter](CloudyWing.SpreadsheetExporter.md 'CloudyWing\.SpreadsheetExporter')

## SpreadsheetManager Class

The spreadsheet manager\.

```csharp
public static class SpreadsheetManager
```

Inheritance [System\.Object](https://learn.microsoft.com/en-us/dotnet/api/system.object 'System\.Object') &#129106; SpreadsheetManager
### Properties

<a name='CloudyWing.SpreadsheetExporter.SpreadsheetManager.DefaultCellStyles'></a>

## SpreadsheetManager\.DefaultCellStyles Property

Gets or sets the configuration\.

```csharp
public static CloudyWing.SpreadsheetExporter.Config.CellStyleConfiguration DefaultCellStyles { get; set; }
```

#### Property Value
[CellStyleConfiguration](CloudyWing.SpreadsheetExporter.Config.CellStyleConfiguration.md 'CloudyWing\.SpreadsheetExporter\.Config\.CellStyleConfiguration')  
The configuration\.
### Methods

<a name='CloudyWing.SpreadsheetExporter.SpreadsheetManager.CreateExporter()'></a>

## SpreadsheetManager\.CreateExporter\(\) Method

Creates the exporter\.

```csharp
public static CloudyWing.SpreadsheetExporter.ISpreadsheetExporter CreateExporter();
```

#### Returns
[ISpreadsheetExporter](CloudyWing.SpreadsheetExporter.ISpreadsheetExporter.md 'CloudyWing\.SpreadsheetExporter\.ISpreadsheetExporter')  
The exporter\.

#### Exceptions

[System\.InvalidOperationException](https://learn.microsoft.com/en-us/dotnet/api/system.invalidoperationexception 'System\.InvalidOperationException')  
Exporter factory is not set\.

<a name='CloudyWing.SpreadsheetExporter.SpreadsheetManager.SetExporter(System.Func_CloudyWing.SpreadsheetExporter.ISpreadsheetExporter_)'></a>

## SpreadsheetManager\.SetExporter\(Func\<ISpreadsheetExporter\>\) Method

Sets the exporter\.

```csharp
public static void SetExporter(System.Func<CloudyWing.SpreadsheetExporter.ISpreadsheetExporter> exporterFactory);
```
#### Parameters

<a name='CloudyWing.SpreadsheetExporter.SpreadsheetManager.SetExporter(System.Func_CloudyWing.SpreadsheetExporter.ISpreadsheetExporter_).exporterFactory'></a>

`exporterFactory` [System\.Func&lt;](https://learn.microsoft.com/en-us/dotnet/api/system.func-1 'System\.Func\`1')[ISpreadsheetExporter](CloudyWing.SpreadsheetExporter.ISpreadsheetExporter.md 'CloudyWing\.SpreadsheetExporter\.ISpreadsheetExporter')[&gt;](https://learn.microsoft.com/en-us/dotnet/api/system.func-1 'System\.Func\`1')

The exporter factory\.

#### Exceptions

[System\.ArgumentNullException](https://learn.microsoft.com/en-us/dotnet/api/system.argumentnullexception 'System\.ArgumentNullException')  
exporterFactory

[System\.ArgumentException](https://learn.microsoft.com/en-us/dotnet/api/system.argumentexception 'System\.ArgumentException')  
Factory return value cannot be null\. \- exporterFactory