#### [CloudyWing\.SpreadsheetExporter](index.md 'index')
### [CloudyWing\.SpreadsheetExporter](CloudyWing.SpreadsheetExporter.md 'CloudyWing\.SpreadsheetExporter')

## Cell Class

The spreadsheet cell\.

```csharp
public class Cell
```

Inheritance [System\.Object](https://learn.microsoft.com/en-us/dotnet/api/system.object 'System\.Object') &#129106; Cell
### Properties

<a name='CloudyWing.SpreadsheetExporter.Cell.CellStyleGenerator'></a>

## Cell\.CellStyleGenerator Property

Gets or sets the cell style generator\. Pass the cell index and row index to the generator\. The  index start at 0\.

```csharp
public System.Func<int,int,CloudyWing.SpreadsheetExporter.CellStyle> CellStyleGenerator { get; set; }
```

#### Property Value
[System\.Func&lt;](https://learn.microsoft.com/en-us/dotnet/api/system.func-3 'System\.Func\`3')[System\.Int32](https://learn.microsoft.com/en-us/dotnet/api/system.int32 'System\.Int32')[,](https://learn.microsoft.com/en-us/dotnet/api/system.func-3 'System\.Func\`3')[System\.Int32](https://learn.microsoft.com/en-us/dotnet/api/system.int32 'System\.Int32')[,](https://learn.microsoft.com/en-us/dotnet/api/system.func-3 'System\.Func\`3')[CellStyle](CloudyWing.SpreadsheetExporter.CellStyle.md 'CloudyWing\.SpreadsheetExporter\.CellStyle')[&gt;](https://learn.microsoft.com/en-us/dotnet/api/system.func-3 'System\.Func\`3')  
The cell style\.

<a name='CloudyWing.SpreadsheetExporter.Cell.FormulaGenerator'></a>

## Cell\.FormulaGenerator Property

Gets or sets the formula generator\. Pass the cell index and row index to the generator\. The  index start at 0\.

```csharp
public System.Func<int,int,string> FormulaGenerator { get; set; }
```

#### Property Value
[System\.Func&lt;](https://learn.microsoft.com/en-us/dotnet/api/system.func-3 'System\.Func\`3')[System\.Int32](https://learn.microsoft.com/en-us/dotnet/api/system.int32 'System\.Int32')[,](https://learn.microsoft.com/en-us/dotnet/api/system.func-3 'System\.Func\`3')[System\.Int32](https://learn.microsoft.com/en-us/dotnet/api/system.int32 'System\.Int32')[,](https://learn.microsoft.com/en-us/dotnet/api/system.func-3 'System\.Func\`3')[System\.String](https://learn.microsoft.com/en-us/dotnet/api/system.string 'System\.String')[&gt;](https://learn.microsoft.com/en-us/dotnet/api/system.func-3 'System\.Func\`3')  
The formula generator\.

<a name='CloudyWing.SpreadsheetExporter.Cell.Point'></a>

## Cell\.Point Property

Gets or sets the point\.

```csharp
public System.Drawing.Point Point { get; set; }
```

#### Property Value
[System\.Drawing\.Point](https://learn.microsoft.com/en-us/dotnet/api/system.drawing.point 'System\.Drawing\.Point')  
The point\.

<a name='CloudyWing.SpreadsheetExporter.Cell.Size'></a>

## Cell\.Size Property

Gets or sets the cell size\.

```csharp
public System.Drawing.Size Size { get; set; }
```

#### Property Value
[System\.Drawing\.Size](https://learn.microsoft.com/en-us/dotnet/api/system.drawing.size 'System\.Drawing\.Size')  
The cell size\.

<a name='CloudyWing.SpreadsheetExporter.Cell.ValueGenerator'></a>

## Cell\.ValueGenerator Property

Gets or sets the cell content value generator\. Pass the cell index and row index to the generator\. The  index start at 0\.

```csharp
public System.Func<int,int,object> ValueGenerator { get; set; }
```

#### Property Value
[System\.Func&lt;](https://learn.microsoft.com/en-us/dotnet/api/system.func-3 'System\.Func\`3')[System\.Int32](https://learn.microsoft.com/en-us/dotnet/api/system.int32 'System\.Int32')[,](https://learn.microsoft.com/en-us/dotnet/api/system.func-3 'System\.Func\`3')[System\.Int32](https://learn.microsoft.com/en-us/dotnet/api/system.int32 'System\.Int32')[,](https://learn.microsoft.com/en-us/dotnet/api/system.func-3 'System\.Func\`3')[System\.Object](https://learn.microsoft.com/en-us/dotnet/api/system.object 'System\.Object')[&gt;](https://learn.microsoft.com/en-us/dotnet/api/system.func-3 'System\.Func\`3')  
The cell content value\.
### Methods

<a name='CloudyWing.SpreadsheetExporter.Cell.GetCellStyle()'></a>

## Cell\.GetCellStyle\(\) Method

Gets the cell style\.

```csharp
public CloudyWing.SpreadsheetExporter.CellStyle GetCellStyle();
```

#### Returns
[CellStyle](CloudyWing.SpreadsheetExporter.CellStyle.md 'CloudyWing\.SpreadsheetExporter\.CellStyle')  
The cell style\.

<a name='CloudyWing.SpreadsheetExporter.Cell.GetFormula()'></a>

## Cell\.GetFormula\(\) Method

Gets the formula\. if formula starts with `=`, automatically removed\.

```csharp
public string GetFormula();
```

#### Returns
[System\.String](https://learn.microsoft.com/en-us/dotnet/api/system.string 'System\.String')  
The the formula\.

<a name='CloudyWing.SpreadsheetExporter.Cell.GetValue()'></a>

## Cell\.GetValue\(\) Method

Gets the content value\.

```csharp
public object GetValue();
```

#### Returns
[System\.Object](https://learn.microsoft.com/en-us/dotnet/api/system.object 'System\.Object')  
The content value\.

<a name='CloudyWing.SpreadsheetExporter.Cell.ShallowCopy()'></a>

## Cell\.ShallowCopy\(\) Method

Shallows the copy\.

```csharp
public CloudyWing.SpreadsheetExporter.Cell ShallowCopy();
```

#### Returns
[Cell](CloudyWing.SpreadsheetExporter.Cell.md 'CloudyWing\.SpreadsheetExporter\.Cell')  
The spreadsheet cell\.