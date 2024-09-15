#### [CloudyWing.SpreadsheetExporter](index.md 'index')
### [CloudyWing.SpreadsheetExporter.Templates](CloudyWing.SpreadsheetExporter.Templates.md 'CloudyWing.SpreadsheetExporter.Templates')

## TemplateContext Class

The template context.

```csharp
public class TemplateContext
```

Inheritance [System.Object](https://docs.microsoft.com/en-us/dotnet/api/System.Object 'System.Object') &#129106; TemplateContext
### Constructors

<a name='CloudyWing.SpreadsheetExporter.Templates.TemplateContext.TemplateContext(System.Collections.Generic.IEnumerable_CloudyWing.SpreadsheetExporter.Cell_,int,System.Collections.Generic.IReadOnlyDictionary_int,System.Nullable_double__)'></a>

## TemplateContext(IEnumerable<Cell>, int, IReadOnlyDictionary<int,Nullable<double>>) Constructor

Initializes a new instance of the [TemplateContext](CloudyWing.SpreadsheetExporter.Templates.TemplateContext.md 'CloudyWing.SpreadsheetExporter.Templates.TemplateContext') class.

```csharp
public TemplateContext(System.Collections.Generic.IEnumerable<CloudyWing.SpreadsheetExporter.Cell> cells, int rowSpan, System.Collections.Generic.IReadOnlyDictionary<int,System.Nullable<double>> rowHeights);
```
#### Parameters

<a name='CloudyWing.SpreadsheetExporter.Templates.TemplateContext.TemplateContext(System.Collections.Generic.IEnumerable_CloudyWing.SpreadsheetExporter.Cell_,int,System.Collections.Generic.IReadOnlyDictionary_int,System.Nullable_double__).cells'></a>

`cells` [System.Collections.Generic.IEnumerable&lt;](https://docs.microsoft.com/en-us/dotnet/api/System.Collections.Generic.IEnumerable-1 'System.Collections.Generic.IEnumerable`1')[Cell](CloudyWing.SpreadsheetExporter.Cell.md 'CloudyWing.SpreadsheetExporter.Cell')[&gt;](https://docs.microsoft.com/en-us/dotnet/api/System.Collections.Generic.IEnumerable-1 'System.Collections.Generic.IEnumerable`1')

The cells.

<a name='CloudyWing.SpreadsheetExporter.Templates.TemplateContext.TemplateContext(System.Collections.Generic.IEnumerable_CloudyWing.SpreadsheetExporter.Cell_,int,System.Collections.Generic.IReadOnlyDictionary_int,System.Nullable_double__).rowSpan'></a>

`rowSpan` [System.Int32](https://docs.microsoft.com/en-us/dotnet/api/System.Int32 'System.Int32')

The row span.

<a name='CloudyWing.SpreadsheetExporter.Templates.TemplateContext.TemplateContext(System.Collections.Generic.IEnumerable_CloudyWing.SpreadsheetExporter.Cell_,int,System.Collections.Generic.IReadOnlyDictionary_int,System.Nullable_double__).rowHeights'></a>

`rowHeights` [System.Collections.Generic.IReadOnlyDictionary&lt;](https://docs.microsoft.com/en-us/dotnet/api/System.Collections.Generic.IReadOnlyDictionary-2 'System.Collections.Generic.IReadOnlyDictionary`2')[System.Int32](https://docs.microsoft.com/en-us/dotnet/api/System.Int32 'System.Int32')[,](https://docs.microsoft.com/en-us/dotnet/api/System.Collections.Generic.IReadOnlyDictionary-2 'System.Collections.Generic.IReadOnlyDictionary`2')[System.Nullable&lt;](https://docs.microsoft.com/en-us/dotnet/api/System.Nullable-1 'System.Nullable`1')[System.Double](https://docs.microsoft.com/en-us/dotnet/api/System.Double 'System.Double')[&gt;](https://docs.microsoft.com/en-us/dotnet/api/System.Nullable-1 'System.Nullable`1')[&gt;](https://docs.microsoft.com/en-us/dotnet/api/System.Collections.Generic.IReadOnlyDictionary-2 'System.Collections.Generic.IReadOnlyDictionary`2')

The row heights.
### Properties

<a name='CloudyWing.SpreadsheetExporter.Templates.TemplateContext.Cells'></a>

## TemplateContext.Cells Property

Gets the cells.

```csharp
public System.Collections.Generic.IReadOnlyList<CloudyWing.SpreadsheetExporter.Cell> Cells { get; }
```

#### Property Value
[System.Collections.Generic.IReadOnlyList&lt;](https://docs.microsoft.com/en-us/dotnet/api/System.Collections.Generic.IReadOnlyList-1 'System.Collections.Generic.IReadOnlyList`1')[Cell](CloudyWing.SpreadsheetExporter.Cell.md 'CloudyWing.SpreadsheetExporter.Cell')[&gt;](https://docs.microsoft.com/en-us/dotnet/api/System.Collections.Generic.IReadOnlyList-1 'System.Collections.Generic.IReadOnlyList`1')  
The cells.

<a name='CloudyWing.SpreadsheetExporter.Templates.TemplateContext.RowHeights'></a>

## TemplateContext.RowHeights Property

Gets the height of rows.

```csharp
public System.Collections.Generic.IReadOnlyDictionary<int,System.Nullable<double>> RowHeights { get; }
```

#### Property Value
[System.Collections.Generic.IReadOnlyDictionary&lt;](https://docs.microsoft.com/en-us/dotnet/api/System.Collections.Generic.IReadOnlyDictionary-2 'System.Collections.Generic.IReadOnlyDictionary`2')[System.Int32](https://docs.microsoft.com/en-us/dotnet/api/System.Int32 'System.Int32')[,](https://docs.microsoft.com/en-us/dotnet/api/System.Collections.Generic.IReadOnlyDictionary-2 'System.Collections.Generic.IReadOnlyDictionary`2')[System.Nullable&lt;](https://docs.microsoft.com/en-us/dotnet/api/System.Nullable-1 'System.Nullable`1')[System.Double](https://docs.microsoft.com/en-us/dotnet/api/System.Double 'System.Double')[&gt;](https://docs.microsoft.com/en-us/dotnet/api/System.Nullable-1 'System.Nullable`1')[&gt;](https://docs.microsoft.com/en-us/dotnet/api/System.Collections.Generic.IReadOnlyDictionary-2 'System.Collections.Generic.IReadOnlyDictionary`2')  
The height of rows.

<a name='CloudyWing.SpreadsheetExporter.Templates.TemplateContext.RowSpan'></a>

## TemplateContext.RowSpan Property

Gets the row span.

```csharp
public int RowSpan { get; }
```

#### Property Value
[System.Int32](https://docs.microsoft.com/en-us/dotnet/api/System.Int32 'System.Int32')  
The row span.