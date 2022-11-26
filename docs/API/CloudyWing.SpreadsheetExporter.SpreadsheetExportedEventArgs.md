#### [CloudyWing.SpreadsheetExporter](index.md 'index')
### [CloudyWing.SpreadsheetExporter](CloudyWing.SpreadsheetExporter.md 'CloudyWing.SpreadsheetExporter')

## SpreadsheetExportedEventArgs Class

Event arguments after spreadsheet export.

```csharp
public class SpreadsheetExportedEventArgs
```

Inheritance [System.Object](https://docs.microsoft.com/en-us/dotnet/api/System.Object 'System.Object') &#129106; SpreadsheetExportedEventArgs
### Constructors

<a name='CloudyWing.SpreadsheetExporter.SpreadsheetExportedEventArgs.SpreadsheetExportedEventArgs(System.Collections.Generic.IEnumerable_CloudyWing.SpreadsheetExporter.SheeterContext_)'></a>

## SpreadsheetExportedEventArgs(IEnumerable<SheeterContext>) Constructor

Initializes a new instance of the [SpreadsheetExportedEventArgs](CloudyWing.SpreadsheetExporter.SpreadsheetExportedEventArgs.md 'CloudyWing.SpreadsheetExporter.SpreadsheetExportedEventArgs') class.

```csharp
public SpreadsheetExportedEventArgs(System.Collections.Generic.IEnumerable<CloudyWing.SpreadsheetExporter.SheeterContext> sheeterContexts);
```
#### Parameters

<a name='CloudyWing.SpreadsheetExporter.SpreadsheetExportedEventArgs.SpreadsheetExportedEventArgs(System.Collections.Generic.IEnumerable_CloudyWing.SpreadsheetExporter.SheeterContext_).sheeterContexts'></a>

`sheeterContexts` [System.Collections.Generic.IEnumerable&lt;](https://docs.microsoft.com/en-us/dotnet/api/System.Collections.Generic.IEnumerable-1 'System.Collections.Generic.IEnumerable`1')[SheeterContext](CloudyWing.SpreadsheetExporter.SheeterContext.md 'CloudyWing.SpreadsheetExporter.SheeterContext')[&gt;](https://docs.microsoft.com/en-us/dotnet/api/System.Collections.Generic.IEnumerable-1 'System.Collections.Generic.IEnumerable`1')

The sheeter contexts.

#### Exceptions

[System.ArgumentNullException](https://docs.microsoft.com/en-us/dotnet/api/System.ArgumentNullException 'System.ArgumentNullException')  
sheeterContexts
### Properties

<a name='CloudyWing.SpreadsheetExporter.SpreadsheetExportedEventArgs.SheeterContexts'></a>

## SpreadsheetExportedEventArgs.SheeterContexts Property

Gets the sheeter contexts.

```csharp
public System.Collections.Generic.IEnumerable<CloudyWing.SpreadsheetExporter.SheeterContext> SheeterContexts { get; }
```

#### Property Value
[System.Collections.Generic.IEnumerable&lt;](https://docs.microsoft.com/en-us/dotnet/api/System.Collections.Generic.IEnumerable-1 'System.Collections.Generic.IEnumerable`1')[SheeterContext](CloudyWing.SpreadsheetExporter.SheeterContext.md 'CloudyWing.SpreadsheetExporter.SheeterContext')[&gt;](https://docs.microsoft.com/en-us/dotnet/api/System.Collections.Generic.IEnumerable-1 'System.Collections.Generic.IEnumerable`1')  
The sheeter contexts.