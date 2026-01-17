#### [CloudyWing\.SpreadsheetExporter](index.md 'index')
### [CloudyWing\.SpreadsheetExporter](CloudyWing.SpreadsheetExporter.md 'CloudyWing\.SpreadsheetExporter')

## CellStyle Struct

The cell style\.

```csharp
public record struct CellStyle : System.IEquatable<CloudyWing.SpreadsheetExporter.CellStyle>
```

Implements [System\.IEquatable&lt;](https://learn.microsoft.com/en-us/dotnet/api/system.iequatable-1 'System\.IEquatable\`1')[CellStyle](CloudyWing.SpreadsheetExporter.CellStyle.md 'CloudyWing\.SpreadsheetExporter\.CellStyle')[&gt;](https://learn.microsoft.com/en-us/dotnet/api/system.iequatable-1 'System\.IEquatable\`1')
### Constructors

<a name='CloudyWing.SpreadsheetExporter.CellStyle.CellStyle(CloudyWing.SpreadsheetExporter.HorizontalAlignment,CloudyWing.SpreadsheetExporter.VerticalAlignment,bool,bool,System.Drawing.Color,CloudyWing.SpreadsheetExporter.CellFont,string,bool)'></a>

## CellStyle\(HorizontalAlignment, VerticalAlignment, bool, bool, Color, CellFont, string, bool\) Constructor

The cell style\.

```csharp
public CellStyle(CloudyWing.SpreadsheetExporter.HorizontalAlignment HorizontalAlignment=CloudyWing.SpreadsheetExporter.HorizontalAlignment.General, CloudyWing.SpreadsheetExporter.VerticalAlignment VerticalAlignment=CloudyWing.SpreadsheetExporter.VerticalAlignment.Top, bool HasBorder=false, bool WrapText=false, System.Drawing.Color BackgroundColor=default(System.Drawing.Color), CloudyWing.SpreadsheetExporter.CellFont Font=default(CloudyWing.SpreadsheetExporter.CellFont), string DataFormat=null, bool IsLocked=false);
```
#### Parameters

<a name='CloudyWing.SpreadsheetExporter.CellStyle.CellStyle(CloudyWing.SpreadsheetExporter.HorizontalAlignment,CloudyWing.SpreadsheetExporter.VerticalAlignment,bool,bool,System.Drawing.Color,CloudyWing.SpreadsheetExporter.CellFont,string,bool).HorizontalAlignment'></a>

`HorizontalAlignment` [HorizontalAlignment](CloudyWing.SpreadsheetExporter.HorizontalAlignment.md 'CloudyWing\.SpreadsheetExporter\.HorizontalAlignment')

Gets the horizontal alignment\.

<a name='CloudyWing.SpreadsheetExporter.CellStyle.CellStyle(CloudyWing.SpreadsheetExporter.HorizontalAlignment,CloudyWing.SpreadsheetExporter.VerticalAlignment,bool,bool,System.Drawing.Color,CloudyWing.SpreadsheetExporter.CellFont,string,bool).VerticalAlignment'></a>

`VerticalAlignment` [VerticalAlignment](CloudyWing.SpreadsheetExporter.VerticalAlignment.md 'CloudyWing\.SpreadsheetExporter\.VerticalAlignment')

Gets the vertical alignment\.

<a name='CloudyWing.SpreadsheetExporter.CellStyle.CellStyle(CloudyWing.SpreadsheetExporter.HorizontalAlignment,CloudyWing.SpreadsheetExporter.VerticalAlignment,bool,bool,System.Drawing.Color,CloudyWing.SpreadsheetExporter.CellFont,string,bool).HasBorder'></a>

`HasBorder` [System\.Boolean](https://learn.microsoft.com/en-us/dotnet/api/system.boolean 'System\.Boolean')

Gets a value indicating whether this instance has border\.

<a name='CloudyWing.SpreadsheetExporter.CellStyle.CellStyle(CloudyWing.SpreadsheetExporter.HorizontalAlignment,CloudyWing.SpreadsheetExporter.VerticalAlignment,bool,bool,System.Drawing.Color,CloudyWing.SpreadsheetExporter.CellFont,string,bool).WrapText'></a>

`WrapText` [System\.Boolean](https://learn.microsoft.com/en-us/dotnet/api/system.boolean 'System\.Boolean')

Gets a value indicating whether \[wrap text\]\.

<a name='CloudyWing.SpreadsheetExporter.CellStyle.CellStyle(CloudyWing.SpreadsheetExporter.HorizontalAlignment,CloudyWing.SpreadsheetExporter.VerticalAlignment,bool,bool,System.Drawing.Color,CloudyWing.SpreadsheetExporter.CellFont,string,bool).BackgroundColor'></a>

`BackgroundColor` [System\.Drawing\.Color](https://learn.microsoft.com/en-us/dotnet/api/system.drawing.color 'System\.Drawing\.Color')

Gets the color of the background\.

<a name='CloudyWing.SpreadsheetExporter.CellStyle.CellStyle(CloudyWing.SpreadsheetExporter.HorizontalAlignment,CloudyWing.SpreadsheetExporter.VerticalAlignment,bool,bool,System.Drawing.Color,CloudyWing.SpreadsheetExporter.CellFont,string,bool).Font'></a>

`Font` [CellFont](CloudyWing.SpreadsheetExporter.CellFont.md 'CloudyWing\.SpreadsheetExporter\.CellFont')

Gets the font\.

<a name='CloudyWing.SpreadsheetExporter.CellStyle.CellStyle(CloudyWing.SpreadsheetExporter.HorizontalAlignment,CloudyWing.SpreadsheetExporter.VerticalAlignment,bool,bool,System.Drawing.Color,CloudyWing.SpreadsheetExporter.CellFont,string,bool).DataFormat'></a>

`DataFormat` [System\.String](https://learn.microsoft.com/en-us/dotnet/api/system.string 'System\.String')

Gets the data format\.

<a name='CloudyWing.SpreadsheetExporter.CellStyle.CellStyle(CloudyWing.SpreadsheetExporter.HorizontalAlignment,CloudyWing.SpreadsheetExporter.VerticalAlignment,bool,bool,System.Drawing.Color,CloudyWing.SpreadsheetExporter.CellFont,string,bool).IsLocked'></a>

`IsLocked` [System\.Boolean](https://learn.microsoft.com/en-us/dotnet/api/system.boolean 'System\.Boolean')

Gets a value indicating whether this instance is locked\.
### Properties

<a name='CloudyWing.SpreadsheetExporter.CellStyle.BackgroundColor'></a>

## CellStyle\.BackgroundColor Property

Gets the color of the background\.

```csharp
public System.Drawing.Color BackgroundColor { get; set; }
```

#### Property Value
[System\.Drawing\.Color](https://learn.microsoft.com/en-us/dotnet/api/system.drawing.color 'System\.Drawing\.Color')

<a name='CloudyWing.SpreadsheetExporter.CellStyle.DataFormat'></a>

## CellStyle\.DataFormat Property

Gets the data format\.

```csharp
public string DataFormat { get; set; }
```

#### Property Value
[System\.String](https://learn.microsoft.com/en-us/dotnet/api/system.string 'System\.String')

<a name='CloudyWing.SpreadsheetExporter.CellStyle.Empty'></a>

## CellStyle\.Empty Property

The cell style equals to `new CellStyle()`\.

```csharp
public static CloudyWing.SpreadsheetExporter.CellStyle Empty { get; }
```

#### Property Value
[CellStyle](CloudyWing.SpreadsheetExporter.CellStyle.md 'CloudyWing\.SpreadsheetExporter\.CellStyle')

<a name='CloudyWing.SpreadsheetExporter.CellStyle.Font'></a>

## CellStyle\.Font Property

Gets the font\.

```csharp
public CloudyWing.SpreadsheetExporter.CellFont Font { get; set; }
```

#### Property Value
[CellFont](CloudyWing.SpreadsheetExporter.CellFont.md 'CloudyWing\.SpreadsheetExporter\.CellFont')

<a name='CloudyWing.SpreadsheetExporter.CellStyle.HasBorder'></a>

## CellStyle\.HasBorder Property

Gets a value indicating whether this instance has border\.

```csharp
public bool HasBorder { get; set; }
```

#### Property Value
[System\.Boolean](https://learn.microsoft.com/en-us/dotnet/api/system.boolean 'System\.Boolean')

<a name='CloudyWing.SpreadsheetExporter.CellStyle.HorizontalAlignment'></a>

## CellStyle\.HorizontalAlignment Property

Gets the horizontal alignment\.

```csharp
public CloudyWing.SpreadsheetExporter.HorizontalAlignment HorizontalAlignment { get; set; }
```

#### Property Value
[HorizontalAlignment](CloudyWing.SpreadsheetExporter.HorizontalAlignment.md 'CloudyWing\.SpreadsheetExporter\.HorizontalAlignment')

<a name='CloudyWing.SpreadsheetExporter.CellStyle.IsLocked'></a>

## CellStyle\.IsLocked Property

Gets a value indicating whether this instance is locked\.

```csharp
public bool IsLocked { get; set; }
```

#### Property Value
[System\.Boolean](https://learn.microsoft.com/en-us/dotnet/api/system.boolean 'System\.Boolean')

<a name='CloudyWing.SpreadsheetExporter.CellStyle.VerticalAlignment'></a>

## CellStyle\.VerticalAlignment Property

Gets the vertical alignment\.

```csharp
public CloudyWing.SpreadsheetExporter.VerticalAlignment VerticalAlignment { get; set; }
```

#### Property Value
[VerticalAlignment](CloudyWing.SpreadsheetExporter.VerticalAlignment.md 'CloudyWing\.SpreadsheetExporter\.VerticalAlignment')

<a name='CloudyWing.SpreadsheetExporter.CellStyle.WrapText'></a>

## CellStyle\.WrapText Property

Gets a value indicating whether \[wrap text\]\.

```csharp
public bool WrapText { get; set; }
```

#### Property Value
[System\.Boolean](https://learn.microsoft.com/en-us/dotnet/api/system.boolean 'System\.Boolean')