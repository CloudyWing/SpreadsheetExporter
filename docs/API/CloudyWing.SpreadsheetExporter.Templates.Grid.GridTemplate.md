#### [CloudyWing.SpreadsheetExporter](index.md 'index')
### [CloudyWing.SpreadsheetExporter.Templates.Grid](CloudyWing.SpreadsheetExporter.Templates.Grid.md 'CloudyWing.SpreadsheetExporter.Templates.Grid')

## GridTemplate Class

The grid template. Create cell information using `CreateRow()` and `CreateCell()`.

```csharp
public class GridTemplate :
CloudyWing.SpreadsheetExporter.Templates.ITemplate
```

Inheritance [System.Object](https://docs.microsoft.com/en-us/dotnet/api/System.Object 'System.Object') &#129106; GridTemplate

Implements [ITemplate](CloudyWing.SpreadsheetExporter.Templates.ITemplate.md 'CloudyWing.SpreadsheetExporter.Templates.ITemplate')
### Properties

<a name='CloudyWing.SpreadsheetExporter.Templates.Grid.GridTemplate.ColumnSpan'></a>

## GridTemplate.ColumnSpan Property

Gets the column span.

```csharp
public int ColumnSpan { get; }
```

#### Property Value
[System.Int32](https://docs.microsoft.com/en-us/dotnet/api/System.Int32 'System.Int32')  
The column span.

<a name='CloudyWing.SpreadsheetExporter.Templates.Grid.GridTemplate.RowSpan'></a>

## GridTemplate.RowSpan Property

Gets the row span.

```csharp
public int RowSpan { get; }
```

#### Property Value
[System.Int32](https://docs.microsoft.com/en-us/dotnet/api/System.Int32 'System.Int32')  
The row span.
### Methods

<a name='CloudyWing.SpreadsheetExporter.Templates.Grid.GridTemplate.CreateCell(object,int,int,System.Nullable_CloudyWing.SpreadsheetExporter.CellStyle_)'></a>

## GridTemplate.CreateCell(object, int, int, Nullable<CellStyle>) Method

Creates the cell.

```csharp
public CloudyWing.SpreadsheetExporter.Templates.Grid.GridTemplate CreateCell(object value, int columnSpan=1, int rowSpan=1, System.Nullable<CloudyWing.SpreadsheetExporter.CellStyle> cellStyle=null);
```
#### Parameters

<a name='CloudyWing.SpreadsheetExporter.Templates.Grid.GridTemplate.CreateCell(object,int,int,System.Nullable_CloudyWing.SpreadsheetExporter.CellStyle_).value'></a>

`value` [System.Object](https://docs.microsoft.com/en-us/dotnet/api/System.Object 'System.Object')

The value.

<a name='CloudyWing.SpreadsheetExporter.Templates.Grid.GridTemplate.CreateCell(object,int,int,System.Nullable_CloudyWing.SpreadsheetExporter.CellStyle_).columnSpan'></a>

`columnSpan` [System.Int32](https://docs.microsoft.com/en-us/dotnet/api/System.Int32 'System.Int32')

The column span.

<a name='CloudyWing.SpreadsheetExporter.Templates.Grid.GridTemplate.CreateCell(object,int,int,System.Nullable_CloudyWing.SpreadsheetExporter.CellStyle_).rowSpan'></a>

`rowSpan` [System.Int32](https://docs.microsoft.com/en-us/dotnet/api/System.Int32 'System.Int32')

The row span.

<a name='CloudyWing.SpreadsheetExporter.Templates.Grid.GridTemplate.CreateCell(object,int,int,System.Nullable_CloudyWing.SpreadsheetExporter.CellStyle_).cellStyle'></a>

`cellStyle` [System.Nullable&lt;](https://docs.microsoft.com/en-us/dotnet/api/System.Nullable-1 'System.Nullable`1')[CellStyle](CloudyWing.SpreadsheetExporter.CellStyle.md 'CloudyWing.SpreadsheetExporter.CellStyle')[&gt;](https://docs.microsoft.com/en-us/dotnet/api/System.Nullable-1 'System.Nullable`1')

The cell style. The default is `SpreadsheetManager.DefaultCellStyles.GridCellStyle`.

#### Returns
[GridTemplate](CloudyWing.SpreadsheetExporter.Templates.Grid.GridTemplate.md 'CloudyWing.SpreadsheetExporter.Templates.Grid.GridTemplate')  
The self.

#### Exceptions

[System.ArgumentOutOfRangeException](https://docs.microsoft.com/en-us/dotnet/api/System.ArgumentOutOfRangeException 'System.ArgumentOutOfRangeException')  
columnSpan - Must be greater than 0.  
            or  
            rowSpan - Must be greater than 0.

<a name='CloudyWing.SpreadsheetExporter.Templates.Grid.GridTemplate.CreateCell(System.Func_int,int,string_,int,int,System.Nullable_CloudyWing.SpreadsheetExporter.CellStyle_)'></a>

## GridTemplate.CreateCell(Func<int,int,string>, int, int, Nullable<CellStyle>) Method

Create cell that contain formula.

```csharp
public CloudyWing.SpreadsheetExporter.Templates.Grid.GridTemplate CreateCell(System.Func<int,int,string> formulaGenerator, int columnSpan=1, int rowSpan=1, System.Nullable<CloudyWing.SpreadsheetExporter.CellStyle> cellStyle=null);
```
#### Parameters

<a name='CloudyWing.SpreadsheetExporter.Templates.Grid.GridTemplate.CreateCell(System.Func_int,int,string_,int,int,System.Nullable_CloudyWing.SpreadsheetExporter.CellStyle_).formulaGenerator'></a>

`formulaGenerator` [System.Func&lt;](https://docs.microsoft.com/en-us/dotnet/api/System.Func-3 'System.Func`3')[System.Int32](https://docs.microsoft.com/en-us/dotnet/api/System.Int32 'System.Int32')[,](https://docs.microsoft.com/en-us/dotnet/api/System.Func-3 'System.Func`3')[System.Int32](https://docs.microsoft.com/en-us/dotnet/api/System.Int32 'System.Int32')[,](https://docs.microsoft.com/en-us/dotnet/api/System.Func-3 'System.Func`3')[System.String](https://docs.microsoft.com/en-us/dotnet/api/System.String 'System.String')[&gt;](https://docs.microsoft.com/en-us/dotnet/api/System.Func-3 'System.Func`3')

The formula generator. Pass the cell index and row index to the generator. The  index start at 0.

<a name='CloudyWing.SpreadsheetExporter.Templates.Grid.GridTemplate.CreateCell(System.Func_int,int,string_,int,int,System.Nullable_CloudyWing.SpreadsheetExporter.CellStyle_).columnSpan'></a>

`columnSpan` [System.Int32](https://docs.microsoft.com/en-us/dotnet/api/System.Int32 'System.Int32')

The column span.

<a name='CloudyWing.SpreadsheetExporter.Templates.Grid.GridTemplate.CreateCell(System.Func_int,int,string_,int,int,System.Nullable_CloudyWing.SpreadsheetExporter.CellStyle_).rowSpan'></a>

`rowSpan` [System.Int32](https://docs.microsoft.com/en-us/dotnet/api/System.Int32 'System.Int32')

The row span.

<a name='CloudyWing.SpreadsheetExporter.Templates.Grid.GridTemplate.CreateCell(System.Func_int,int,string_,int,int,System.Nullable_CloudyWing.SpreadsheetExporter.CellStyle_).cellStyle'></a>

`cellStyle` [System.Nullable&lt;](https://docs.microsoft.com/en-us/dotnet/api/System.Nullable-1 'System.Nullable`1')[CellStyle](CloudyWing.SpreadsheetExporter.CellStyle.md 'CloudyWing.SpreadsheetExporter.CellStyle')[&gt;](https://docs.microsoft.com/en-us/dotnet/api/System.Nullable-1 'System.Nullable`1')

The cell style. The default is `SpreadsheetManager.DefaultCellStyles.GridCellStyle`.

#### Returns
[GridTemplate](CloudyWing.SpreadsheetExporter.Templates.Grid.GridTemplate.md 'CloudyWing.SpreadsheetExporter.Templates.Grid.GridTemplate')  
The self.

#### Exceptions

[System.ArgumentOutOfRangeException](https://docs.microsoft.com/en-us/dotnet/api/System.ArgumentOutOfRangeException 'System.ArgumentOutOfRangeException')  
columnSpan - Must be greater than 0.  
            or  
            rowSpan - Must be greater than 0.

<a name='CloudyWing.SpreadsheetExporter.Templates.Grid.GridTemplate.CreateRow(double)'></a>

## GridTemplate.CreateRow(double) Method

Creates the row.

```csharp
public CloudyWing.SpreadsheetExporter.Templates.Grid.GridTemplate CreateRow(double height=16.5);
```
#### Parameters

<a name='CloudyWing.SpreadsheetExporter.Templates.Grid.GridTemplate.CreateRow(double).height'></a>

`height` [System.Double](https://docs.microsoft.com/en-us/dotnet/api/System.Double 'System.Double')

The height.

#### Returns
[GridTemplate](CloudyWing.SpreadsheetExporter.Templates.Grid.GridTemplate.md 'CloudyWing.SpreadsheetExporter.Templates.Grid.GridTemplate')  
The self.

<a name='CloudyWing.SpreadsheetExporter.Templates.Grid.GridTemplate.GetContext()'></a>

## GridTemplate.GetContext() Method

Gets the context.

```csharp
public CloudyWing.SpreadsheetExporter.Templates.TemplateContext GetContext();
```

Implements [GetContext()](CloudyWing.SpreadsheetExporter.Templates.ITemplate.md#CloudyWing.SpreadsheetExporter.Templates.ITemplate.GetContext() 'CloudyWing.SpreadsheetExporter.Templates.ITemplate.GetContext()')

#### Returns
[TemplateContext](CloudyWing.SpreadsheetExporter.Templates.TemplateContext.md 'CloudyWing.SpreadsheetExporter.Templates.TemplateContext')  
The template context.