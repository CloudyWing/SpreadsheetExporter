#### [CloudyWing\.SpreadsheetExporter](index.md 'index')
### [CloudyWing\.SpreadsheetExporter\.Templates](CloudyWing.SpreadsheetExporter.Templates.md 'CloudyWing\.SpreadsheetExporter\.Templates')

## TemplateContext Class

The template context\.

```csharp
public class TemplateContext
```

Inheritance [System\.Object](https://learn.microsoft.com/en-us/dotnet/api/system.object 'System\.Object') &#129106; TemplateContext
### Constructors

<a name='CloudyWing.SpreadsheetExporter.Templates.TemplateContext.TemplateContext(System.Collections.Generic.IEnumerable_CloudyWing.SpreadsheetExporter.Cell_,int,System.Collections.Generic.IReadOnlyDictionary_int,System.Nullable_double__)'></a>

## TemplateContext\(IEnumerable\<Cell\>, int, IReadOnlyDictionary\<int,Nullable\<double\>\>\) Constructor

The template context\.

```csharp
public TemplateContext(System.Collections.Generic.IEnumerable<CloudyWing.SpreadsheetExporter.Cell> cells, int rowSpan, System.Collections.Generic.IReadOnlyDictionary<int,System.Nullable<double>> rowHeights);
```
#### Parameters

<a name='CloudyWing.SpreadsheetExporter.Templates.TemplateContext.TemplateContext(System.Collections.Generic.IEnumerable_CloudyWing.SpreadsheetExporter.Cell_,int,System.Collections.Generic.IReadOnlyDictionary_int,System.Nullable_double__).cells'></a>

`cells` [System\.Collections\.Generic\.IEnumerable&lt;](https://learn.microsoft.com/en-us/dotnet/api/system.collections.generic.ienumerable-1 'System\.Collections\.Generic\.IEnumerable\`1')[Cell](CloudyWing.SpreadsheetExporter.Cell.md 'CloudyWing\.SpreadsheetExporter\.Cell')[&gt;](https://learn.microsoft.com/en-us/dotnet/api/system.collections.generic.ienumerable-1 'System\.Collections\.Generic\.IEnumerable\`1')

The cells\.

<a name='CloudyWing.SpreadsheetExporter.Templates.TemplateContext.TemplateContext(System.Collections.Generic.IEnumerable_CloudyWing.SpreadsheetExporter.Cell_,int,System.Collections.Generic.IReadOnlyDictionary_int,System.Nullable_double__).rowSpan'></a>

`rowSpan` [System\.Int32](https://learn.microsoft.com/en-us/dotnet/api/system.int32 'System\.Int32')

The row span\.

<a name='CloudyWing.SpreadsheetExporter.Templates.TemplateContext.TemplateContext(System.Collections.Generic.IEnumerable_CloudyWing.SpreadsheetExporter.Cell_,int,System.Collections.Generic.IReadOnlyDictionary_int,System.Nullable_double__).rowHeights'></a>

`rowHeights` [System\.Collections\.Generic\.IReadOnlyDictionary&lt;](https://learn.microsoft.com/en-us/dotnet/api/system.collections.generic.ireadonlydictionary-2 'System\.Collections\.Generic\.IReadOnlyDictionary\`2')[System\.Int32](https://learn.microsoft.com/en-us/dotnet/api/system.int32 'System\.Int32')[,](https://learn.microsoft.com/en-us/dotnet/api/system.collections.generic.ireadonlydictionary-2 'System\.Collections\.Generic\.IReadOnlyDictionary\`2')[System\.Nullable&lt;](https://learn.microsoft.com/en-us/dotnet/api/system.nullable-1 'System\.Nullable\`1')[System\.Double](https://learn.microsoft.com/en-us/dotnet/api/system.double 'System\.Double')[&gt;](https://learn.microsoft.com/en-us/dotnet/api/system.nullable-1 'System\.Nullable\`1')[&gt;](https://learn.microsoft.com/en-us/dotnet/api/system.collections.generic.ireadonlydictionary-2 'System\.Collections\.Generic\.IReadOnlyDictionary\`2')

The row heights\.
### Properties

<a name='CloudyWing.SpreadsheetExporter.Templates.TemplateContext.Cells'></a>

## TemplateContext\.Cells Property

Gets the cells\.

```csharp
public System.Collections.Generic.IReadOnlyList<CloudyWing.SpreadsheetExporter.Cell> Cells { get; }
```

#### Property Value
[System\.Collections\.Generic\.IReadOnlyList&lt;](https://learn.microsoft.com/en-us/dotnet/api/system.collections.generic.ireadonlylist-1 'System\.Collections\.Generic\.IReadOnlyList\`1')[Cell](CloudyWing.SpreadsheetExporter.Cell.md 'CloudyWing\.SpreadsheetExporter\.Cell')[&gt;](https://learn.microsoft.com/en-us/dotnet/api/system.collections.generic.ireadonlylist-1 'System\.Collections\.Generic\.IReadOnlyList\`1')  
The cells\.

<a name='CloudyWing.SpreadsheetExporter.Templates.TemplateContext.RowHeights'></a>

## TemplateContext\.RowHeights Property

Gets the height of rows\.

```csharp
public System.Collections.Generic.IReadOnlyDictionary<int,System.Nullable<double>> RowHeights { get; }
```

#### Property Value
[System\.Collections\.Generic\.IReadOnlyDictionary&lt;](https://learn.microsoft.com/en-us/dotnet/api/system.collections.generic.ireadonlydictionary-2 'System\.Collections\.Generic\.IReadOnlyDictionary\`2')[System\.Int32](https://learn.microsoft.com/en-us/dotnet/api/system.int32 'System\.Int32')[,](https://learn.microsoft.com/en-us/dotnet/api/system.collections.generic.ireadonlydictionary-2 'System\.Collections\.Generic\.IReadOnlyDictionary\`2')[System\.Nullable&lt;](https://learn.microsoft.com/en-us/dotnet/api/system.nullable-1 'System\.Nullable\`1')[System\.Double](https://learn.microsoft.com/en-us/dotnet/api/system.double 'System\.Double')[&gt;](https://learn.microsoft.com/en-us/dotnet/api/system.nullable-1 'System\.Nullable\`1')[&gt;](https://learn.microsoft.com/en-us/dotnet/api/system.collections.generic.ireadonlydictionary-2 'System\.Collections\.Generic\.IReadOnlyDictionary\`2')  
The height of rows\.

<a name='CloudyWing.SpreadsheetExporter.Templates.TemplateContext.RowSpan'></a>

## TemplateContext\.RowSpan Property

Gets the row span\.

```csharp
public int RowSpan { get; }
```

#### Property Value
[System\.Int32](https://learn.microsoft.com/en-us/dotnet/api/system.int32 'System\.Int32')  
The row span\.