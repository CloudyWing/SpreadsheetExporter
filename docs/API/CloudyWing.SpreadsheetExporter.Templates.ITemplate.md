#### [CloudyWing.SpreadsheetExporter](index.md 'index')
### [CloudyWing.SpreadsheetExporter.Templates](CloudyWing.SpreadsheetExporter.Templates.md 'CloudyWing.SpreadsheetExporter.Templates')

## ITemplate Interface

The template interface.

```csharp
public interface ITemplate
```

Derived  
&#8627; [GridTemplate](CloudyWing.SpreadsheetExporter.Templates.Grid.GridTemplate.md 'CloudyWing.SpreadsheetExporter.Templates.Grid.GridTemplate')  
&#8627; [MergedTemplate](CloudyWing.SpreadsheetExporter.Templates.MergedTemplate.md 'CloudyWing.SpreadsheetExporter.Templates.MergedTemplate')  
&#8627; [RecordSetTemplate&lt;T&gt;](CloudyWing.SpreadsheetExporter.Templates.RecordSet.RecordSetTemplate_T_.md 'CloudyWing.SpreadsheetExporter.Templates.RecordSet.RecordSetTemplate<T>')
### Properties

<a name='CloudyWing.SpreadsheetExporter.Templates.ITemplate.RowHeights'></a>

## ITemplate.RowHeights Property

Gets the height of rows.

```csharp
System.Collections.Generic.IReadOnlyDictionary<int,double> RowHeights { get; }
```

#### Property Value
[System.Collections.Generic.IReadOnlyDictionary&lt;](https://docs.microsoft.com/en-us/dotnet/api/System.Collections.Generic.IReadOnlyDictionary-2 'System.Collections.Generic.IReadOnlyDictionary`2')[System.Int32](https://docs.microsoft.com/en-us/dotnet/api/System.Int32 'System.Int32')[,](https://docs.microsoft.com/en-us/dotnet/api/System.Collections.Generic.IReadOnlyDictionary-2 'System.Collections.Generic.IReadOnlyDictionary`2')[System.Double](https://docs.microsoft.com/en-us/dotnet/api/System.Double 'System.Double')[&gt;](https://docs.microsoft.com/en-us/dotnet/api/System.Collections.Generic.IReadOnlyDictionary-2 'System.Collections.Generic.IReadOnlyDictionary`2')  
The height of rows.
### Methods

<a name='CloudyWing.SpreadsheetExporter.Templates.ITemplate.GetContext()'></a>

## ITemplate.GetContext() Method

Gets the context.

```csharp
CloudyWing.SpreadsheetExporter.Templates.TemplateContext GetContext();
```

#### Returns
[TemplateContext](CloudyWing.SpreadsheetExporter.Templates.TemplateContext.md 'CloudyWing.SpreadsheetExporter.Templates.TemplateContext')  
The template context.