#### [CloudyWing\.SpreadsheetExporter](index.md 'index')
### [CloudyWing\.SpreadsheetExporter](CloudyWing.SpreadsheetExporter.md 'CloudyWing\.SpreadsheetExporter')

## CellFont Struct

The cell font\.

```csharp
public record struct CellFont : System.IEquatable<CloudyWing.SpreadsheetExporter.CellFont>
```

Implements [System\.IEquatable&lt;](https://learn.microsoft.com/en-us/dotnet/api/system.iequatable-1 'System\.IEquatable\`1')[CellFont](CloudyWing.SpreadsheetExporter.CellFont.md 'CloudyWing\.SpreadsheetExporter\.CellFont')[&gt;](https://learn.microsoft.com/en-us/dotnet/api/system.iequatable-1 'System\.IEquatable\`1')
### Constructors

<a name='CloudyWing.SpreadsheetExporter.CellFont.CellFont(string,short,System.Drawing.Color,CloudyWing.SpreadsheetExporter.FontStyles)'></a>

## CellFont\(string, short, Color, FontStyles\) Constructor

The cell font\.

```csharp
public CellFont(string Name=null, short Size=0, System.Drawing.Color Color=default(System.Drawing.Color), CloudyWing.SpreadsheetExporter.FontStyles Style=CloudyWing.SpreadsheetExporter.FontStyles.None);
```
#### Parameters

<a name='CloudyWing.SpreadsheetExporter.CellFont.CellFont(string,short,System.Drawing.Color,CloudyWing.SpreadsheetExporter.FontStyles).Name'></a>

`Name` [System\.String](https://learn.microsoft.com/en-us/dotnet/api/system.string 'System\.String')

The font name\.

<a name='CloudyWing.SpreadsheetExporter.CellFont.CellFont(string,short,System.Drawing.Color,CloudyWing.SpreadsheetExporter.FontStyles).Size'></a>

`Size` [System\.Int16](https://learn.microsoft.com/en-us/dotnet/api/system.int16 'System\.Int16')

The size\.

<a name='CloudyWing.SpreadsheetExporter.CellFont.CellFont(string,short,System.Drawing.Color,CloudyWing.SpreadsheetExporter.FontStyles).Color'></a>

`Color` [System\.Drawing\.Color](https://learn.microsoft.com/en-us/dotnet/api/system.drawing.color 'System\.Drawing\.Color')

The color\.

<a name='CloudyWing.SpreadsheetExporter.CellFont.CellFont(string,short,System.Drawing.Color,CloudyWing.SpreadsheetExporter.FontStyles).Style'></a>

`Style` [FontStyles](CloudyWing.SpreadsheetExporter.FontStyles.md 'CloudyWing\.SpreadsheetExporter\.FontStyles')

The style\.
### Properties

<a name='CloudyWing.SpreadsheetExporter.CellFont.Color'></a>

## CellFont\.Color Property

The color\.

```csharp
public System.Drawing.Color Color { get; set; }
```

#### Property Value
[System\.Drawing\.Color](https://learn.microsoft.com/en-us/dotnet/api/system.drawing.color 'System\.Drawing\.Color')

<a name='CloudyWing.SpreadsheetExporter.CellFont.Empty'></a>

## CellFont\.Empty Property

The cell font equals to `new CellFont()`\.

```csharp
public static CloudyWing.SpreadsheetExporter.CellFont Empty { get; }
```

#### Property Value
[CellFont](CloudyWing.SpreadsheetExporter.CellFont.md 'CloudyWing\.SpreadsheetExporter\.CellFont')

<a name='CloudyWing.SpreadsheetExporter.CellFont.Name'></a>

## CellFont\.Name Property

The font name\.

```csharp
public string Name { get; set; }
```

#### Property Value
[System\.String](https://learn.microsoft.com/en-us/dotnet/api/system.string 'System\.String')

<a name='CloudyWing.SpreadsheetExporter.CellFont.Size'></a>

## CellFont\.Size Property

The size\.

```csharp
public short Size { get; set; }
```

#### Property Value
[System\.Int16](https://learn.microsoft.com/en-us/dotnet/api/system.int16 'System\.Int16')

<a name='CloudyWing.SpreadsheetExporter.CellFont.Style'></a>

## CellFont\.Style Property

The style\.

```csharp
public CloudyWing.SpreadsheetExporter.FontStyles Style { get; set; }
```

#### Property Value
[FontStyles](CloudyWing.SpreadsheetExporter.FontStyles.md 'CloudyWing\.SpreadsheetExporter\.FontStyles')