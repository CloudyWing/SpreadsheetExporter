#### [CloudyWing.SpreadsheetExporter](index.md 'index')
### [CloudyWing.SpreadsheetExporter](CloudyWing.SpreadsheetExporter.md 'CloudyWing.SpreadsheetExporter')

## Cell Class

The spreadsheet cell.

```csharp
public class Cell
```

Inheritance [System.Object](https://docs.microsoft.com/en-us/dotnet/api/System.Object 'System.Object') &#129106; Cell
### Properties

<a name='CloudyWing.SpreadsheetExporter.Cell.CellStyle'></a>

## Cell.CellStyle Property

Gets or sets the cell style.

```csharp
public CloudyWing.SpreadsheetExporter.CellStyle CellStyle { get; set; }
```

#### Property Value
[CellStyle](CloudyWing.SpreadsheetExporter.CellStyle.md 'CloudyWing.SpreadsheetExporter.CellStyle')  
The cell style.

<a name='CloudyWing.SpreadsheetExporter.Cell.Formula'></a>

## Cell.Formula Property

Gets or sets the formula.

```csharp
public string Formula { get; set; }
```

#### Property Value
[System.String](https://docs.microsoft.com/en-us/dotnet/api/System.String 'System.String')  
The formula.

<a name='CloudyWing.SpreadsheetExporter.Cell.Point'></a>

## Cell.Point Property

Gets or sets the point.

```csharp
public System.Drawing.Point Point { get; set; }
```

#### Property Value
[System.Drawing.Point](https://docs.microsoft.com/en-us/dotnet/api/System.Drawing.Point 'System.Drawing.Point')  
The point.

<a name='CloudyWing.SpreadsheetExporter.Cell.Size'></a>

## Cell.Size Property

Gets or sets the cell size.

```csharp
public System.Drawing.Size Size { get; set; }
```

#### Property Value
[System.Drawing.Size](https://docs.microsoft.com/en-us/dotnet/api/System.Drawing.Size 'System.Drawing.Size')  
The cell size.

<a name='CloudyWing.SpreadsheetExporter.Cell.Value'></a>

## Cell.Value Property

Gets or sets the cell content value.

```csharp
public object Value { get; set; }
```

#### Property Value
[System.Object](https://docs.microsoft.com/en-us/dotnet/api/System.Object 'System.Object')  
The cell content value.
### Methods

<a name='CloudyWing.SpreadsheetExporter.Cell.ShallowCopy()'></a>

## Cell.ShallowCopy() Method

Shallows the copy.

```csharp
public CloudyWing.SpreadsheetExporter.Cell ShallowCopy();
```

#### Returns
[Cell](CloudyWing.SpreadsheetExporter.Cell.md 'CloudyWing.SpreadsheetExporter.Cell')  
The spreadsheet cell.