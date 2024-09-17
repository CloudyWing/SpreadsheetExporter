#### [CloudyWing.SpreadsheetExporter](index.md 'index')
### [CloudyWing.SpreadsheetExporter](CloudyWing.SpreadsheetExporter.md 'CloudyWing.SpreadsheetExporter')

## SpreadsheetExportingEventArgs Class

Event arguments before spreadsheet export.

```csharp
public class SpreadsheetExportingEventArgs
```

Inheritance [System.Object](https://docs.microsoft.com/en-us/dotnet/api/System.Object 'System.Object') &#129106; SpreadsheetExportingEventArgs

#### Exceptions

[System.ArgumentNullException](https://docs.microsoft.com/en-us/dotnet/api/System.ArgumentNullException 'System.ArgumentNullException')  
sheeterContexts
### Constructors

<a name='CloudyWing.SpreadsheetExporter.SpreadsheetExportingEventArgs.SpreadsheetExportingEventArgs(System.Collections.Generic.IEnumerable_CloudyWing.SpreadsheetExporter.SheeterContext_)'></a>

## SpreadsheetExportingEventArgs(IEnumerable<SheeterContext>) Constructor

Event arguments before spreadsheet export.

```csharp
public SpreadsheetExportingEventArgs(System.Collections.Generic.IEnumerable<CloudyWing.SpreadsheetExporter.SheeterContext> sheeterContexts);
```
#### Parameters

<a name='CloudyWing.SpreadsheetExporter.SpreadsheetExportingEventArgs.SpreadsheetExportingEventArgs(System.Collections.Generic.IEnumerable_CloudyWing.SpreadsheetExporter.SheeterContext_).sheeterContexts'></a>

`sheeterContexts` [System.Collections.Generic.IEnumerable&lt;](https://docs.microsoft.com/en-us/dotnet/api/System.Collections.Generic.IEnumerable-1 'System.Collections.Generic.IEnumerable`1')[SheeterContext](CloudyWing.SpreadsheetExporter.SheeterContext.md 'CloudyWing.SpreadsheetExporter.SheeterContext')[&gt;](https://docs.microsoft.com/en-us/dotnet/api/System.Collections.Generic.IEnumerable-1 'System.Collections.Generic.IEnumerable`1')

The sheeter contexts.

#### Exceptions

[System.ArgumentNullException](https://docs.microsoft.com/en-us/dotnet/api/System.ArgumentNullException 'System.ArgumentNullException')  
sheeterContexts
### Properties

<a name='CloudyWing.SpreadsheetExporter.SpreadsheetExportingEventArgs.SheeterContexts'></a>

## SpreadsheetExportingEventArgs.SheeterContexts Property

Gets the sheeter contexts.

```csharp
public System.Collections.Generic.IEnumerable<CloudyWing.SpreadsheetExporter.SheeterContext> SheeterContexts { get; }
```

#### Property Value
[System.Collections.Generic.IEnumerable&lt;](https://docs.microsoft.com/en-us/dotnet/api/System.Collections.Generic.IEnumerable-1 'System.Collections.Generic.IEnumerable`1')[SheeterContext](CloudyWing.SpreadsheetExporter.SheeterContext.md 'CloudyWing.SpreadsheetExporter.SheeterContext')[&gt;](https://docs.microsoft.com/en-us/dotnet/api/System.Collections.Generic.IEnumerable-1 'System.Collections.Generic.IEnumerable`1')  
The sheeter contexts.