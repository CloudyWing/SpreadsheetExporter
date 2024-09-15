#### [CloudyWing.SpreadsheetExporter](index.md 'index')
### [CloudyWing.SpreadsheetExporter](CloudyWing.SpreadsheetExporter.md 'CloudyWing.SpreadsheetExporter')

## SheeterContext Class

The sheeter context.

```csharp
public class SheeterContext
```

Inheritance [System.Object](https://docs.microsoft.com/en-us/dotnet/api/System.Object 'System.Object') &#129106; SheeterContext
### Constructors

<a name='CloudyWing.SpreadsheetExporter.SheeterContext.SheeterContext(CloudyWing.SpreadsheetExporter.Sheeter)'></a>

## SheeterContext(Sheeter) Constructor

Initializes a new instance of the [SheeterContext](CloudyWing.SpreadsheetExporter.SheeterContext.md 'CloudyWing.SpreadsheetExporter.SheeterContext') class.

```csharp
public SheeterContext(CloudyWing.SpreadsheetExporter.Sheeter sheeter);
```
#### Parameters

<a name='CloudyWing.SpreadsheetExporter.SheeterContext.SheeterContext(CloudyWing.SpreadsheetExporter.Sheeter).sheeter'></a>

`sheeter` [Sheeter](CloudyWing.SpreadsheetExporter.Sheeter.md 'CloudyWing.SpreadsheetExporter.Sheeter')

The sheeter.
### Properties

<a name='CloudyWing.SpreadsheetExporter.SheeterContext.Cells'></a>

## SheeterContext.Cells Property

Gets the cells.

```csharp
public System.Collections.Generic.IReadOnlyList<CloudyWing.SpreadsheetExporter.Cell> Cells { get; }
```

#### Property Value
[System.Collections.Generic.IReadOnlyList&lt;](https://docs.microsoft.com/en-us/dotnet/api/System.Collections.Generic.IReadOnlyList-1 'System.Collections.Generic.IReadOnlyList`1')[Cell](CloudyWing.SpreadsheetExporter.Cell.md 'CloudyWing.SpreadsheetExporter.Cell')[&gt;](https://docs.microsoft.com/en-us/dotnet/api/System.Collections.Generic.IReadOnlyList-1 'System.Collections.Generic.IReadOnlyList`1')  
The cells.

<a name='CloudyWing.SpreadsheetExporter.SheeterContext.ColumnWidths'></a>

## SheeterContext.ColumnWidths Property

Gets the width of columns.

```csharp
public System.Collections.Generic.IReadOnlyDictionary<int,double> ColumnWidths { get; }
```

#### Property Value
[System.Collections.Generic.IReadOnlyDictionary&lt;](https://docs.microsoft.com/en-us/dotnet/api/System.Collections.Generic.IReadOnlyDictionary-2 'System.Collections.Generic.IReadOnlyDictionary`2')[System.Int32](https://docs.microsoft.com/en-us/dotnet/api/System.Int32 'System.Int32')[,](https://docs.microsoft.com/en-us/dotnet/api/System.Collections.Generic.IReadOnlyDictionary-2 'System.Collections.Generic.IReadOnlyDictionary`2')[System.Double](https://docs.microsoft.com/en-us/dotnet/api/System.Double 'System.Double')[&gt;](https://docs.microsoft.com/en-us/dotnet/api/System.Collections.Generic.IReadOnlyDictionary-2 'System.Collections.Generic.IReadOnlyDictionary`2')  
The width of columns.

<a name='CloudyWing.SpreadsheetExporter.SheeterContext.DefaultRowHeight'></a>

## SheeterContext.DefaultRowHeight Property

Gets or sets the default height of the row.

```csharp
public System.Nullable<double> DefaultRowHeight { get; set; }
```

#### Property Value
[System.Nullable&lt;](https://docs.microsoft.com/en-us/dotnet/api/System.Nullable-1 'System.Nullable`1')[System.Double](https://docs.microsoft.com/en-us/dotnet/api/System.Double 'System.Double')[&gt;](https://docs.microsoft.com/en-us/dotnet/api/System.Nullable-1 'System.Nullable`1')  
The default height of the row.

<a name='CloudyWing.SpreadsheetExporter.SheeterContext.IsProtected'></a>

## SheeterContext.IsProtected Property

Gets a value indicating whether this instance is protected.

```csharp
public bool IsProtected { get; }
```

#### Property Value
[System.Boolean](https://docs.microsoft.com/en-us/dotnet/api/System.Boolean 'System.Boolean')  
`true` if this instance is protected; otherwise, `false`.

<a name='CloudyWing.SpreadsheetExporter.SheeterContext.PageSettings'></a>

## SheeterContext.PageSettings Property

Gets the page settings.

```csharp
public CloudyWing.SpreadsheetExporter.PageSettings PageSettings { get; }
```

#### Property Value
[PageSettings](CloudyWing.SpreadsheetExporter.PageSettings.md 'CloudyWing.SpreadsheetExporter.PageSettings')  
The page settings.

<a name='CloudyWing.SpreadsheetExporter.SheeterContext.Password'></a>

## SheeterContext.Password Property

Gets the password.

```csharp
public string Password { get; }
```

#### Property Value
[System.String](https://docs.microsoft.com/en-us/dotnet/api/System.String 'System.String')  
The password.

<a name='CloudyWing.SpreadsheetExporter.SheeterContext.RowHeights'></a>

## SheeterContext.RowHeights Property

Gets the height of rows.

```csharp
public System.Collections.Generic.IReadOnlyDictionary<int,System.Nullable<double>> RowHeights { get; set; }
```

#### Property Value
[System.Collections.Generic.IReadOnlyDictionary&lt;](https://docs.microsoft.com/en-us/dotnet/api/System.Collections.Generic.IReadOnlyDictionary-2 'System.Collections.Generic.IReadOnlyDictionary`2')[System.Int32](https://docs.microsoft.com/en-us/dotnet/api/System.Int32 'System.Int32')[,](https://docs.microsoft.com/en-us/dotnet/api/System.Collections.Generic.IReadOnlyDictionary-2 'System.Collections.Generic.IReadOnlyDictionary`2')[System.Nullable&lt;](https://docs.microsoft.com/en-us/dotnet/api/System.Nullable-1 'System.Nullable`1')[System.Double](https://docs.microsoft.com/en-us/dotnet/api/System.Double 'System.Double')[&gt;](https://docs.microsoft.com/en-us/dotnet/api/System.Nullable-1 'System.Nullable`1')[&gt;](https://docs.microsoft.com/en-us/dotnet/api/System.Collections.Generic.IReadOnlyDictionary-2 'System.Collections.Generic.IReadOnlyDictionary`2')  
The height of rows.

<a name='CloudyWing.SpreadsheetExporter.SheeterContext.SheetName'></a>

## SheeterContext.SheetName Property

Gets the name of the sheet.

```csharp
public string SheetName { get; }
```

#### Property Value
[System.String](https://docs.microsoft.com/en-us/dotnet/api/System.String 'System.String')  
The name of the sheet.