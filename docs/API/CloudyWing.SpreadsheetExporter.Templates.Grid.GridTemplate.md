#### [CloudyWing\.SpreadsheetExporter](index.md 'index')
### [CloudyWing\.SpreadsheetExporter\.Templates\.Grid](CloudyWing.SpreadsheetExporter.Templates.Grid.md 'CloudyWing\.SpreadsheetExporter\.Templates\.Grid')

## GridTemplate Class

The grid template\. Create cell information using `CreateRow()` and `CreateCell()`\.

```csharp
public class GridTemplate : CloudyWing.SpreadsheetExporter.Templates.ITemplate
```

Inheritance [System\.Object](https://learn.microsoft.com/en-us/dotnet/api/system.object 'System\.Object') &#129106; GridTemplate

Implements [ITemplate](CloudyWing.SpreadsheetExporter.Templates.ITemplate.md 'CloudyWing\.SpreadsheetExporter\.Templates\.ITemplate')
### Properties

<a name='CloudyWing.SpreadsheetExporter.Templates.Grid.GridTemplate.ColumnSpan'></a>

## GridTemplate\.ColumnSpan Property

Gets the column span\.

```csharp
public int ColumnSpan { get; }
```

#### Property Value
[System\.Int32](https://learn.microsoft.com/en-us/dotnet/api/system.int32 'System\.Int32')  
The column span\.

<a name='CloudyWing.SpreadsheetExporter.Templates.Grid.GridTemplate.RowSpan'></a>

## GridTemplate\.RowSpan Property

Gets the row span\.

```csharp
public int RowSpan { get; }
```

#### Property Value
[System\.Int32](https://learn.microsoft.com/en-us/dotnet/api/system.int32 'System\.Int32')  
The row span\.
### Methods

<a name='CloudyWing.SpreadsheetExporter.Templates.Grid.GridTemplate.CreateCell(object,int,int,System.Nullable_CloudyWing.SpreadsheetExporter.CellStyle_)'></a>

## GridTemplate\.CreateCell\(object, int, int, Nullable\<CellStyle\>\) Method

Creates the cell\.

```csharp
public CloudyWing.SpreadsheetExporter.Templates.Grid.GridTemplate CreateCell(object value, int columnSpan=1, int rowSpan=1, System.Nullable<CloudyWing.SpreadsheetExporter.CellStyle> cellStyle=null);
```
#### Parameters

<a name='CloudyWing.SpreadsheetExporter.Templates.Grid.GridTemplate.CreateCell(object,int,int,System.Nullable_CloudyWing.SpreadsheetExporter.CellStyle_).value'></a>

`value` [System\.Object](https://learn.microsoft.com/en-us/dotnet/api/system.object 'System\.Object')

The value\.

<a name='CloudyWing.SpreadsheetExporter.Templates.Grid.GridTemplate.CreateCell(object,int,int,System.Nullable_CloudyWing.SpreadsheetExporter.CellStyle_).columnSpan'></a>

`columnSpan` [System\.Int32](https://learn.microsoft.com/en-us/dotnet/api/system.int32 'System\.Int32')

The column span\.

<a name='CloudyWing.SpreadsheetExporter.Templates.Grid.GridTemplate.CreateCell(object,int,int,System.Nullable_CloudyWing.SpreadsheetExporter.CellStyle_).rowSpan'></a>

`rowSpan` [System\.Int32](https://learn.microsoft.com/en-us/dotnet/api/system.int32 'System\.Int32')

The row span\.

<a name='CloudyWing.SpreadsheetExporter.Templates.Grid.GridTemplate.CreateCell(object,int,int,System.Nullable_CloudyWing.SpreadsheetExporter.CellStyle_).cellStyle'></a>

`cellStyle` [System\.Nullable&lt;](https://learn.microsoft.com/en-us/dotnet/api/system.nullable-1 'System\.Nullable\`1')[CellStyle](CloudyWing.SpreadsheetExporter.CellStyle.md 'CloudyWing\.SpreadsheetExporter\.CellStyle')[&gt;](https://learn.microsoft.com/en-us/dotnet/api/system.nullable-1 'System\.Nullable\`1')

The cell style\. The default is `SpreadsheetManager.DefaultCellStyles.GridCellStyle`\.

#### Returns
[GridTemplate](CloudyWing.SpreadsheetExporter.Templates.Grid.GridTemplate.md 'CloudyWing\.SpreadsheetExporter\.Templates\.Grid\.GridTemplate')  
The self\.

#### Exceptions

[System\.ArgumentOutOfRangeException](https://learn.microsoft.com/en-us/dotnet/api/system.argumentoutofrangeexception 'System\.ArgumentOutOfRangeException')  
[columnSpan](CloudyWing.SpreadsheetExporter.Templates.Grid.GridTemplate.md#CloudyWing.SpreadsheetExporter.Templates.Grid.GridTemplate.CreateCell(object,int,int,System.Nullable_CloudyWing.SpreadsheetExporter.CellStyle_).columnSpan 'CloudyWing\.SpreadsheetExporter\.Templates\.Grid\.GridTemplate\.CreateCell\(object, int, int, System\.Nullable\<CloudyWing\.SpreadsheetExporter\.CellStyle\>\)\.columnSpan') must be greater than 0\.

[System\.ArgumentOutOfRangeException](https://learn.microsoft.com/en-us/dotnet/api/system.argumentoutofrangeexception 'System\.ArgumentOutOfRangeException')  
[rowSpan](CloudyWing.SpreadsheetExporter.Templates.Grid.GridTemplate.md#CloudyWing.SpreadsheetExporter.Templates.Grid.GridTemplate.CreateCell(object,int,int,System.Nullable_CloudyWing.SpreadsheetExporter.CellStyle_).rowSpan 'CloudyWing\.SpreadsheetExporter\.Templates\.Grid\.GridTemplate\.CreateCell\(object, int, int, System\.Nullable\<CloudyWing\.SpreadsheetExporter\.CellStyle\>\)\.rowSpan') must be greater than 0\.

<a name='CloudyWing.SpreadsheetExporter.Templates.Grid.GridTemplate.CreateCell(System.Action_CloudyWing.SpreadsheetExporter.Cell_,int,int,System.Nullable_CloudyWing.SpreadsheetExporter.CellStyle_)'></a>

## GridTemplate\.CreateCell\(Action\<Cell\>, int, int, Nullable\<CellStyle\>\) Method

Creates a cell with custom configuration using a configurator action\.
This method provides full flexibility for configuring cell properties including value, formula, data validation, and style\.

```csharp
public CloudyWing.SpreadsheetExporter.Templates.Grid.GridTemplate CreateCell(System.Action<CloudyWing.SpreadsheetExporter.Cell> configurator, int columnSpan=1, int rowSpan=1, System.Nullable<CloudyWing.SpreadsheetExporter.CellStyle> cellStyle=null);
```
#### Parameters

<a name='CloudyWing.SpreadsheetExporter.Templates.Grid.GridTemplate.CreateCell(System.Action_CloudyWing.SpreadsheetExporter.Cell_,int,int,System.Nullable_CloudyWing.SpreadsheetExporter.CellStyle_).configurator'></a>

`configurator` [System\.Action&lt;](https://learn.microsoft.com/en-us/dotnet/api/system.action-1 'System\.Action\`1')[Cell](CloudyWing.SpreadsheetExporter.Cell.md 'CloudyWing\.SpreadsheetExporter\.Cell')[&gt;](https://learn.microsoft.com/en-us/dotnet/api/system.action-1 'System\.Action\`1')

The action to configure the cell\. The action receives the cell instance to configure\.

<a name='CloudyWing.SpreadsheetExporter.Templates.Grid.GridTemplate.CreateCell(System.Action_CloudyWing.SpreadsheetExporter.Cell_,int,int,System.Nullable_CloudyWing.SpreadsheetExporter.CellStyle_).columnSpan'></a>

`columnSpan` [System\.Int32](https://learn.microsoft.com/en-us/dotnet/api/system.int32 'System\.Int32')

The column span\.

<a name='CloudyWing.SpreadsheetExporter.Templates.Grid.GridTemplate.CreateCell(System.Action_CloudyWing.SpreadsheetExporter.Cell_,int,int,System.Nullable_CloudyWing.SpreadsheetExporter.CellStyle_).rowSpan'></a>

`rowSpan` [System\.Int32](https://learn.microsoft.com/en-us/dotnet/api/system.int32 'System\.Int32')

The row span\.

<a name='CloudyWing.SpreadsheetExporter.Templates.Grid.GridTemplate.CreateCell(System.Action_CloudyWing.SpreadsheetExporter.Cell_,int,int,System.Nullable_CloudyWing.SpreadsheetExporter.CellStyle_).cellStyle'></a>

`cellStyle` [System\.Nullable&lt;](https://learn.microsoft.com/en-us/dotnet/api/system.nullable-1 'System\.Nullable\`1')[CellStyle](CloudyWing.SpreadsheetExporter.CellStyle.md 'CloudyWing\.SpreadsheetExporter\.CellStyle')[&gt;](https://learn.microsoft.com/en-us/dotnet/api/system.nullable-1 'System\.Nullable\`1')

The cell style\. The default is `SpreadsheetManager.DefaultCellStyles.GridCellStyle`\.

#### Returns
[GridTemplate](CloudyWing.SpreadsheetExporter.Templates.Grid.GridTemplate.md 'CloudyWing\.SpreadsheetExporter\.Templates\.Grid\.GridTemplate')  
The self\.

#### Exceptions

[System\.ArgumentOutOfRangeException](https://learn.microsoft.com/en-us/dotnet/api/system.argumentoutofrangeexception 'System\.ArgumentOutOfRangeException')  
[columnSpan](CloudyWing.SpreadsheetExporter.Templates.Grid.GridTemplate.md#CloudyWing.SpreadsheetExporter.Templates.Grid.GridTemplate.CreateCell(System.Action_CloudyWing.SpreadsheetExporter.Cell_,int,int,System.Nullable_CloudyWing.SpreadsheetExporter.CellStyle_).columnSpan 'CloudyWing\.SpreadsheetExporter\.Templates\.Grid\.GridTemplate\.CreateCell\(System\.Action\<CloudyWing\.SpreadsheetExporter\.Cell\>, int, int, System\.Nullable\<CloudyWing\.SpreadsheetExporter\.CellStyle\>\)\.columnSpan') must be greater than 0\.

[System\.ArgumentOutOfRangeException](https://learn.microsoft.com/en-us/dotnet/api/system.argumentoutofrangeexception 'System\.ArgumentOutOfRangeException')  
[rowSpan](CloudyWing.SpreadsheetExporter.Templates.Grid.GridTemplate.md#CloudyWing.SpreadsheetExporter.Templates.Grid.GridTemplate.CreateCell(System.Action_CloudyWing.SpreadsheetExporter.Cell_,int,int,System.Nullable_CloudyWing.SpreadsheetExporter.CellStyle_).rowSpan 'CloudyWing\.SpreadsheetExporter\.Templates\.Grid\.GridTemplate\.CreateCell\(System\.Action\<CloudyWing\.SpreadsheetExporter\.Cell\>, int, int, System\.Nullable\<CloudyWing\.SpreadsheetExporter\.CellStyle\>\)\.rowSpan') must be greater than 0\.

<a name='CloudyWing.SpreadsheetExporter.Templates.Grid.GridTemplate.CreateCell(System.Func_int,int,string_,int,int,System.Nullable_CloudyWing.SpreadsheetExporter.CellStyle_)'></a>

## GridTemplate\.CreateCell\(Func\<int,int,string\>, int, int, Nullable\<CellStyle\>\) Method

Create cell that contain formula\.

```csharp
public CloudyWing.SpreadsheetExporter.Templates.Grid.GridTemplate CreateCell(System.Func<int,int,string> formulaGenerator, int columnSpan=1, int rowSpan=1, System.Nullable<CloudyWing.SpreadsheetExporter.CellStyle> cellStyle=null);
```
#### Parameters

<a name='CloudyWing.SpreadsheetExporter.Templates.Grid.GridTemplate.CreateCell(System.Func_int,int,string_,int,int,System.Nullable_CloudyWing.SpreadsheetExporter.CellStyle_).formulaGenerator'></a>

`formulaGenerator` [System\.Func&lt;](https://learn.microsoft.com/en-us/dotnet/api/system.func-3 'System\.Func\`3')[System\.Int32](https://learn.microsoft.com/en-us/dotnet/api/system.int32 'System\.Int32')[,](https://learn.microsoft.com/en-us/dotnet/api/system.func-3 'System\.Func\`3')[System\.Int32](https://learn.microsoft.com/en-us/dotnet/api/system.int32 'System\.Int32')[,](https://learn.microsoft.com/en-us/dotnet/api/system.func-3 'System\.Func\`3')[System\.String](https://learn.microsoft.com/en-us/dotnet/api/system.string 'System\.String')[&gt;](https://learn.microsoft.com/en-us/dotnet/api/system.func-3 'System\.Func\`3')

The formula generator\. Pass the cell index and row index to the generator\. The  index start at 0\.

<a name='CloudyWing.SpreadsheetExporter.Templates.Grid.GridTemplate.CreateCell(System.Func_int,int,string_,int,int,System.Nullable_CloudyWing.SpreadsheetExporter.CellStyle_).columnSpan'></a>

`columnSpan` [System\.Int32](https://learn.microsoft.com/en-us/dotnet/api/system.int32 'System\.Int32')

The column span\.

<a name='CloudyWing.SpreadsheetExporter.Templates.Grid.GridTemplate.CreateCell(System.Func_int,int,string_,int,int,System.Nullable_CloudyWing.SpreadsheetExporter.CellStyle_).rowSpan'></a>

`rowSpan` [System\.Int32](https://learn.microsoft.com/en-us/dotnet/api/system.int32 'System\.Int32')

The row span\.

<a name='CloudyWing.SpreadsheetExporter.Templates.Grid.GridTemplate.CreateCell(System.Func_int,int,string_,int,int,System.Nullable_CloudyWing.SpreadsheetExporter.CellStyle_).cellStyle'></a>

`cellStyle` [System\.Nullable&lt;](https://learn.microsoft.com/en-us/dotnet/api/system.nullable-1 'System\.Nullable\`1')[CellStyle](CloudyWing.SpreadsheetExporter.CellStyle.md 'CloudyWing\.SpreadsheetExporter\.CellStyle')[&gt;](https://learn.microsoft.com/en-us/dotnet/api/system.nullable-1 'System\.Nullable\`1')

The cell style\. The default is `SpreadsheetManager.DefaultCellStyles.GridCellStyle`\.

#### Returns
[GridTemplate](CloudyWing.SpreadsheetExporter.Templates.Grid.GridTemplate.md 'CloudyWing\.SpreadsheetExporter\.Templates\.Grid\.GridTemplate')  
The self\.

#### Exceptions

[System\.ArgumentOutOfRangeException](https://learn.microsoft.com/en-us/dotnet/api/system.argumentoutofrangeexception 'System\.ArgumentOutOfRangeException')  
[columnSpan](CloudyWing.SpreadsheetExporter.Templates.Grid.GridTemplate.md#CloudyWing.SpreadsheetExporter.Templates.Grid.GridTemplate.CreateCell(System.Func_int,int,string_,int,int,System.Nullable_CloudyWing.SpreadsheetExporter.CellStyle_).columnSpan 'CloudyWing\.SpreadsheetExporter\.Templates\.Grid\.GridTemplate\.CreateCell\(System\.Func\<int,int,string\>, int, int, System\.Nullable\<CloudyWing\.SpreadsheetExporter\.CellStyle\>\)\.columnSpan') must be greater than 0\.

[System\.ArgumentOutOfRangeException](https://learn.microsoft.com/en-us/dotnet/api/system.argumentoutofrangeexception 'System\.ArgumentOutOfRangeException')  
[rowSpan](CloudyWing.SpreadsheetExporter.Templates.Grid.GridTemplate.md#CloudyWing.SpreadsheetExporter.Templates.Grid.GridTemplate.CreateCell(System.Func_int,int,string_,int,int,System.Nullable_CloudyWing.SpreadsheetExporter.CellStyle_).rowSpan 'CloudyWing\.SpreadsheetExporter\.Templates\.Grid\.GridTemplate\.CreateCell\(System\.Func\<int,int,string\>, int, int, System\.Nullable\<CloudyWing\.SpreadsheetExporter\.CellStyle\>\)\.rowSpan') must be greater than 0\.

<a name='CloudyWing.SpreadsheetExporter.Templates.Grid.GridTemplate.CreateRow(System.Nullable_double_)'></a>

## GridTemplate\.CreateRow\(Nullable\<double\>\) Method

Creates the row\.

```csharp
public CloudyWing.SpreadsheetExporter.Templates.Grid.GridTemplate CreateRow(System.Nullable<double> height=null);
```
#### Parameters

<a name='CloudyWing.SpreadsheetExporter.Templates.Grid.GridTemplate.CreateRow(System.Nullable_double_).height'></a>

`height` [System\.Nullable&lt;](https://learn.microsoft.com/en-us/dotnet/api/system.nullable-1 'System\.Nullable\`1')[System\.Double](https://learn.microsoft.com/en-us/dotnet/api/system.double 'System\.Double')[&gt;](https://learn.microsoft.com/en-us/dotnet/api/system.nullable-1 'System\.Nullable\`1')

The height\.

#### Returns
[GridTemplate](CloudyWing.SpreadsheetExporter.Templates.Grid.GridTemplate.md 'CloudyWing\.SpreadsheetExporter\.Templates\.Grid\.GridTemplate')  
The self\.

<a name='CloudyWing.SpreadsheetExporter.Templates.Grid.GridTemplate.GetContext()'></a>

## GridTemplate\.GetContext\(\) Method

Gets the context\.

```csharp
public CloudyWing.SpreadsheetExporter.Templates.TemplateContext GetContext();
```

Implements [GetContext\(\)](CloudyWing.SpreadsheetExporter.Templates.ITemplate.md#CloudyWing.SpreadsheetExporter.Templates.ITemplate.GetContext() 'CloudyWing\.SpreadsheetExporter\.Templates\.ITemplate\.GetContext\(\)')

#### Returns
[TemplateContext](CloudyWing.SpreadsheetExporter.Templates.TemplateContext.md 'CloudyWing\.SpreadsheetExporter\.Templates\.TemplateContext')  
The template context\.