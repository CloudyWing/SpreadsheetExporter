#### [CloudyWing.SpreadsheetExporter](index.md 'index')
### [CloudyWing.SpreadsheetExporter](CloudyWing.SpreadsheetExporter.md 'CloudyWing.SpreadsheetExporter')

## CellStyle Struct

The cell style.

```csharp
public struct CellStyle :
System.IEquatable<CloudyWing.SpreadsheetExporter.CellStyle>
```

Implements [System.IEquatable&lt;](https://docs.microsoft.com/en-us/dotnet/api/System.IEquatable-1 'System.IEquatable`1')[CellStyle](CloudyWing.SpreadsheetExporter.CellStyle.md 'CloudyWing.SpreadsheetExporter.CellStyle')[&gt;](https://docs.microsoft.com/en-us/dotnet/api/System.IEquatable-1 'System.IEquatable`1')

### See Also
- [System.IEquatable&lt;&gt;](https://docs.microsoft.com/en-us/dotnet/api/System.IEquatable-1 'System.IEquatable`1')
### Constructors

<a name='CloudyWing.SpreadsheetExporter.CellStyle.CellStyle(CloudyWing.SpreadsheetExporter.HorizontalAlignment,CloudyWing.SpreadsheetExporter.VerticalAlignment,bool,bool,System.Nullable_System.Drawing.Color_,System.Nullable_CloudyWing.SpreadsheetExporter.CellFont_,string,bool)'></a>

## CellStyle(HorizontalAlignment, VerticalAlignment, bool, bool, Nullable<Color>, Nullable<CellFont>, string, bool) Constructor

Initializes a new instance of the [CellStyle](CloudyWing.SpreadsheetExporter.CellStyle.md 'CloudyWing.SpreadsheetExporter.CellStyle') struct.

```csharp
public CellStyle(CloudyWing.SpreadsheetExporter.HorizontalAlignment halign=CloudyWing.SpreadsheetExporter.HorizontalAlignment.Center, CloudyWing.SpreadsheetExporter.VerticalAlignment valign=CloudyWing.SpreadsheetExporter.VerticalAlignment.Middle, bool hasBorder=false, bool wrapText=false, System.Nullable<System.Drawing.Color> backgroundColor=null, System.Nullable<CloudyWing.SpreadsheetExporter.CellFont> font=null, string dataFormat=null, bool isLocked=false);
```
#### Parameters

<a name='CloudyWing.SpreadsheetExporter.CellStyle.CellStyle(CloudyWing.SpreadsheetExporter.HorizontalAlignment,CloudyWing.SpreadsheetExporter.VerticalAlignment,bool,bool,System.Nullable_System.Drawing.Color_,System.Nullable_CloudyWing.SpreadsheetExporter.CellFont_,string,bool).halign'></a>

`halign` [HorizontalAlignment](CloudyWing.SpreadsheetExporter.HorizontalAlignment.md 'CloudyWing.SpreadsheetExporter.HorizontalAlignment')

The halign.

<a name='CloudyWing.SpreadsheetExporter.CellStyle.CellStyle(CloudyWing.SpreadsheetExporter.HorizontalAlignment,CloudyWing.SpreadsheetExporter.VerticalAlignment,bool,bool,System.Nullable_System.Drawing.Color_,System.Nullable_CloudyWing.SpreadsheetExporter.CellFont_,string,bool).valign'></a>

`valign` [VerticalAlignment](CloudyWing.SpreadsheetExporter.VerticalAlignment.md 'CloudyWing.SpreadsheetExporter.VerticalAlignment')

The valign.

<a name='CloudyWing.SpreadsheetExporter.CellStyle.CellStyle(CloudyWing.SpreadsheetExporter.HorizontalAlignment,CloudyWing.SpreadsheetExporter.VerticalAlignment,bool,bool,System.Nullable_System.Drawing.Color_,System.Nullable_CloudyWing.SpreadsheetExporter.CellFont_,string,bool).hasBorder'></a>

`hasBorder` [System.Boolean](https://docs.microsoft.com/en-us/dotnet/api/System.Boolean 'System.Boolean')

if set to `true` [has border].

<a name='CloudyWing.SpreadsheetExporter.CellStyle.CellStyle(CloudyWing.SpreadsheetExporter.HorizontalAlignment,CloudyWing.SpreadsheetExporter.VerticalAlignment,bool,bool,System.Nullable_System.Drawing.Color_,System.Nullable_CloudyWing.SpreadsheetExporter.CellFont_,string,bool).wrapText'></a>

`wrapText` [System.Boolean](https://docs.microsoft.com/en-us/dotnet/api/System.Boolean 'System.Boolean')

if set to `true` [wrap text].

<a name='CloudyWing.SpreadsheetExporter.CellStyle.CellStyle(CloudyWing.SpreadsheetExporter.HorizontalAlignment,CloudyWing.SpreadsheetExporter.VerticalAlignment,bool,bool,System.Nullable_System.Drawing.Color_,System.Nullable_CloudyWing.SpreadsheetExporter.CellFont_,string,bool).backgroundColor'></a>

`backgroundColor` [System.Nullable&lt;](https://docs.microsoft.com/en-us/dotnet/api/System.Nullable-1 'System.Nullable`1')[System.Drawing.Color](https://docs.microsoft.com/en-us/dotnet/api/System.Drawing.Color 'System.Drawing.Color')[&gt;](https://docs.microsoft.com/en-us/dotnet/api/System.Nullable-1 'System.Nullable`1')

Color of the background. The default is `Color.Empty`.

<a name='CloudyWing.SpreadsheetExporter.CellStyle.CellStyle(CloudyWing.SpreadsheetExporter.HorizontalAlignment,CloudyWing.SpreadsheetExporter.VerticalAlignment,bool,bool,System.Nullable_System.Drawing.Color_,System.Nullable_CloudyWing.SpreadsheetExporter.CellFont_,string,bool).font'></a>

`font` [System.Nullable&lt;](https://docs.microsoft.com/en-us/dotnet/api/System.Nullable-1 'System.Nullable`1')[CellFont](CloudyWing.SpreadsheetExporter.CellFont.md 'CloudyWing.SpreadsheetExporter.CellFont')[&gt;](https://docs.microsoft.com/en-us/dotnet/api/System.Nullable-1 'System.Nullable`1')

The font. The default is `CellFont.Empty`.

<a name='CloudyWing.SpreadsheetExporter.CellStyle.CellStyle(CloudyWing.SpreadsheetExporter.HorizontalAlignment,CloudyWing.SpreadsheetExporter.VerticalAlignment,bool,bool,System.Nullable_System.Drawing.Color_,System.Nullable_CloudyWing.SpreadsheetExporter.CellFont_,string,bool).dataFormat'></a>

`dataFormat` [System.String](https://docs.microsoft.com/en-us/dotnet/api/System.String 'System.String')

The data format.

<a name='CloudyWing.SpreadsheetExporter.CellStyle.CellStyle(CloudyWing.SpreadsheetExporter.HorizontalAlignment,CloudyWing.SpreadsheetExporter.VerticalAlignment,bool,bool,System.Nullable_System.Drawing.Color_,System.Nullable_CloudyWing.SpreadsheetExporter.CellFont_,string,bool).isLocked'></a>

`isLocked` [System.Boolean](https://docs.microsoft.com/en-us/dotnet/api/System.Boolean 'System.Boolean')

if set to `true` [is locked].
### Fields

<a name='CloudyWing.SpreadsheetExporter.CellStyle.Empty'></a>

## CellStyle.Empty Field

The cell style equals to `new CellStyle()`.

```csharp
public static CellStyle Empty;
```

#### Field Value
[CellStyle](CloudyWing.SpreadsheetExporter.CellStyle.md 'CloudyWing.SpreadsheetExporter.CellStyle')
### Properties

<a name='CloudyWing.SpreadsheetExporter.CellStyle.BackgroundColor'></a>

## CellStyle.BackgroundColor Property

Gets the color of the background.

```csharp
public System.Drawing.Color BackgroundColor { get; set; }
```

#### Property Value
[System.Drawing.Color](https://docs.microsoft.com/en-us/dotnet/api/System.Drawing.Color 'System.Drawing.Color')  
The color of the background.

<a name='CloudyWing.SpreadsheetExporter.CellStyle.DataFormat'></a>

## CellStyle.DataFormat Property

Gets the data format.

```csharp
public string DataFormat { get; set; }
```

#### Property Value
[System.String](https://docs.microsoft.com/en-us/dotnet/api/System.String 'System.String')  
The data format.

<a name='CloudyWing.SpreadsheetExporter.CellStyle.Font'></a>

## CellStyle.Font Property

Gets the font.

```csharp
public CloudyWing.SpreadsheetExporter.CellFont Font { get; set; }
```

#### Property Value
[CellFont](CloudyWing.SpreadsheetExporter.CellFont.md 'CloudyWing.SpreadsheetExporter.CellFont')  
The font.

<a name='CloudyWing.SpreadsheetExporter.CellStyle.HasBorder'></a>

## CellStyle.HasBorder Property

Gets a value indicating whether this instance has border.

```csharp
public bool HasBorder { get; set; }
```

#### Property Value
[System.Boolean](https://docs.microsoft.com/en-us/dotnet/api/System.Boolean 'System.Boolean')  
`true` if this instance has border; otherwise, `false`.

<a name='CloudyWing.SpreadsheetExporter.CellStyle.HorizontalAlignment'></a>

## CellStyle.HorizontalAlignment Property

Gets the horizontal alignment.

```csharp
public CloudyWing.SpreadsheetExporter.HorizontalAlignment HorizontalAlignment { get; set; }
```

#### Property Value
[HorizontalAlignment](CloudyWing.SpreadsheetExporter.HorizontalAlignment.md 'CloudyWing.SpreadsheetExporter.HorizontalAlignment')  
The horizontal alignment.

<a name='CloudyWing.SpreadsheetExporter.CellStyle.IsLocked'></a>

## CellStyle.IsLocked Property

Gets a value indicating whether this instance is locked.

```csharp
public bool IsLocked { get; set; }
```

#### Property Value
[System.Boolean](https://docs.microsoft.com/en-us/dotnet/api/System.Boolean 'System.Boolean')  
`true` if this instance is locked; otherwise, `false`.

<a name='CloudyWing.SpreadsheetExporter.CellStyle.VerticalAlignment'></a>

## CellStyle.VerticalAlignment Property

Gets the vertical alignment.

```csharp
public CloudyWing.SpreadsheetExporter.VerticalAlignment VerticalAlignment { get; set; }
```

#### Property Value
[VerticalAlignment](CloudyWing.SpreadsheetExporter.VerticalAlignment.md 'CloudyWing.SpreadsheetExporter.VerticalAlignment')  
The vertical alignment.

<a name='CloudyWing.SpreadsheetExporter.CellStyle.WrapText'></a>

## CellStyle.WrapText Property

Gets a value indicating whether [wrap text].

```csharp
public bool WrapText { get; set; }
```

#### Property Value
[System.Boolean](https://docs.microsoft.com/en-us/dotnet/api/System.Boolean 'System.Boolean')  
`true` if [wrap text]; otherwise, `false`.
### Methods

<a name='CloudyWing.SpreadsheetExporter.CellStyle.CloneAndSetBackgroundColor(System.Drawing.Color)'></a>

## CellStyle.CloneAndSetBackgroundColor(Color) Method

Clones and set background color of new instance.

```csharp
public CloudyWing.SpreadsheetExporter.CellStyle CloneAndSetBackgroundColor(System.Drawing.Color backgroundColor);
```
#### Parameters

<a name='CloudyWing.SpreadsheetExporter.CellStyle.CloneAndSetBackgroundColor(System.Drawing.Color).backgroundColor'></a>

`backgroundColor` [System.Drawing.Color](https://docs.microsoft.com/en-us/dotnet/api/System.Drawing.Color 'System.Drawing.Color')

Color of the background.

#### Returns
[CellStyle](CloudyWing.SpreadsheetExporter.CellStyle.md 'CloudyWing.SpreadsheetExporter.CellStyle')  
The cloned new cell style.

<a name='CloudyWing.SpreadsheetExporter.CellStyle.CloneAndSetBorder(bool)'></a>

## CellStyle.CloneAndSetBorder(bool) Method

Clones and set border of new instance.

```csharp
public CloudyWing.SpreadsheetExporter.CellStyle CloneAndSetBorder(bool hasBolder);
```
#### Parameters

<a name='CloudyWing.SpreadsheetExporter.CellStyle.CloneAndSetBorder(bool).hasBolder'></a>

`hasBolder` [System.Boolean](https://docs.microsoft.com/en-us/dotnet/api/System.Boolean 'System.Boolean')

if set to `true` [has bolder].

#### Returns
[CellStyle](CloudyWing.SpreadsheetExporter.CellStyle.md 'CloudyWing.SpreadsheetExporter.CellStyle')  
The cloned new cell style.

<a name='CloudyWing.SpreadsheetExporter.CellStyle.CloneAndSetDataFormat(string)'></a>

## CellStyle.CloneAndSetDataFormat(string) Method

Clones and set data format of new instance.

```csharp
public CloudyWing.SpreadsheetExporter.CellStyle CloneAndSetDataFormat(string dataForamt);
```
#### Parameters

<a name='CloudyWing.SpreadsheetExporter.CellStyle.CloneAndSetDataFormat(string).dataForamt'></a>

`dataForamt` [System.String](https://docs.microsoft.com/en-us/dotnet/api/System.String 'System.String')

The data foramt.

#### Returns
[CellStyle](CloudyWing.SpreadsheetExporter.CellStyle.md 'CloudyWing.SpreadsheetExporter.CellStyle')  
The cloned new cell style.

<a name='CloudyWing.SpreadsheetExporter.CellStyle.CloneAndSetFont(CloudyWing.SpreadsheetExporter.CellFont)'></a>

## CellStyle.CloneAndSetFont(CellFont) Method

Clones and set font of new instance.

```csharp
public CloudyWing.SpreadsheetExporter.CellStyle CloneAndSetFont(CloudyWing.SpreadsheetExporter.CellFont font);
```
#### Parameters

<a name='CloudyWing.SpreadsheetExporter.CellStyle.CloneAndSetFont(CloudyWing.SpreadsheetExporter.CellFont).font'></a>

`font` [CellFont](CloudyWing.SpreadsheetExporter.CellFont.md 'CloudyWing.SpreadsheetExporter.CellFont')

The font.

#### Returns
[CellStyle](CloudyWing.SpreadsheetExporter.CellStyle.md 'CloudyWing.SpreadsheetExporter.CellStyle')  
The cloned new cell style.

<a name='CloudyWing.SpreadsheetExporter.CellStyle.CloneAndSetHorizontalAlignment(CloudyWing.SpreadsheetExporter.HorizontalAlignment)'></a>

## CellStyle.CloneAndSetHorizontalAlignment(HorizontalAlignment) Method

Clones and set horizontal alignment of new instance.

```csharp
public CloudyWing.SpreadsheetExporter.CellStyle CloneAndSetHorizontalAlignment(CloudyWing.SpreadsheetExporter.HorizontalAlignment align);
```
#### Parameters

<a name='CloudyWing.SpreadsheetExporter.CellStyle.CloneAndSetHorizontalAlignment(CloudyWing.SpreadsheetExporter.HorizontalAlignment).align'></a>

`align` [HorizontalAlignment](CloudyWing.SpreadsheetExporter.HorizontalAlignment.md 'CloudyWing.SpreadsheetExporter.HorizontalAlignment')

The align.

#### Returns
[CellStyle](CloudyWing.SpreadsheetExporter.CellStyle.md 'CloudyWing.SpreadsheetExporter.CellStyle')  
The cloned new cell style.

<a name='CloudyWing.SpreadsheetExporter.CellStyle.CloneAndSetLocked(bool)'></a>

## CellStyle.CloneAndSetLocked(bool) Method

Clones and set lockedof of new instance.

```csharp
public CloudyWing.SpreadsheetExporter.CellStyle CloneAndSetLocked(bool isLocked);
```
#### Parameters

<a name='CloudyWing.SpreadsheetExporter.CellStyle.CloneAndSetLocked(bool).isLocked'></a>

`isLocked` [System.Boolean](https://docs.microsoft.com/en-us/dotnet/api/System.Boolean 'System.Boolean')

if set to `true` [is locked].

#### Returns
[CellStyle](CloudyWing.SpreadsheetExporter.CellStyle.md 'CloudyWing.SpreadsheetExporter.CellStyle')  
The cloned new cell style.

<a name='CloudyWing.SpreadsheetExporter.CellStyle.CloneAndSetVerticalAlignment(CloudyWing.SpreadsheetExporter.VerticalAlignment)'></a>

## CellStyle.CloneAndSetVerticalAlignment(VerticalAlignment) Method

Clones and set vertical alignment of new instance.

```csharp
public CloudyWing.SpreadsheetExporter.CellStyle CloneAndSetVerticalAlignment(CloudyWing.SpreadsheetExporter.VerticalAlignment valign);
```
#### Parameters

<a name='CloudyWing.SpreadsheetExporter.CellStyle.CloneAndSetVerticalAlignment(CloudyWing.SpreadsheetExporter.VerticalAlignment).valign'></a>

`valign` [VerticalAlignment](CloudyWing.SpreadsheetExporter.VerticalAlignment.md 'CloudyWing.SpreadsheetExporter.VerticalAlignment')

The valign.

#### Returns
[CellStyle](CloudyWing.SpreadsheetExporter.CellStyle.md 'CloudyWing.SpreadsheetExporter.CellStyle')  
The cloned new cell style.

<a name='CloudyWing.SpreadsheetExporter.CellStyle.CloneAndSetWrapText(bool)'></a>

## CellStyle.CloneAndSetWrapText(bool) Method

Clones and set wrap text of new instance.

```csharp
public CloudyWing.SpreadsheetExporter.CellStyle CloneAndSetWrapText(bool wrapText);
```
#### Parameters

<a name='CloudyWing.SpreadsheetExporter.CellStyle.CloneAndSetWrapText(bool).wrapText'></a>

`wrapText` [System.Boolean](https://docs.microsoft.com/en-us/dotnet/api/System.Boolean 'System.Boolean')

if set to `true` [wrap text].

#### Returns
[CellStyle](CloudyWing.SpreadsheetExporter.CellStyle.md 'CloudyWing.SpreadsheetExporter.CellStyle')  
The cloned new cell style.
### Operators

<a name='CloudyWing.SpreadsheetExporter.CellStyle.op_Equality(CloudyWing.SpreadsheetExporter.CellStyle,CloudyWing.SpreadsheetExporter.CellStyle)'></a>

## CellStyle.operator ==(CellStyle, CellStyle) Operator

Implements the operator op_Equality.

```csharp
public static bool operator ==(CloudyWing.SpreadsheetExporter.CellStyle left, CloudyWing.SpreadsheetExporter.CellStyle right);
```
#### Parameters

<a name='CloudyWing.SpreadsheetExporter.CellStyle.op_Equality(CloudyWing.SpreadsheetExporter.CellStyle,CloudyWing.SpreadsheetExporter.CellStyle).left'></a>

`left` [CellStyle](CloudyWing.SpreadsheetExporter.CellStyle.md 'CloudyWing.SpreadsheetExporter.CellStyle')

The left.

<a name='CloudyWing.SpreadsheetExporter.CellStyle.op_Equality(CloudyWing.SpreadsheetExporter.CellStyle,CloudyWing.SpreadsheetExporter.CellStyle).right'></a>

`right` [CellStyle](CloudyWing.SpreadsheetExporter.CellStyle.md 'CloudyWing.SpreadsheetExporter.CellStyle')

The right.

#### Returns
[System.Boolean](https://docs.microsoft.com/en-us/dotnet/api/System.Boolean 'System.Boolean')  
The result of the operator.

<a name='CloudyWing.SpreadsheetExporter.CellStyle.op_Inequality(CloudyWing.SpreadsheetExporter.CellStyle,CloudyWing.SpreadsheetExporter.CellStyle)'></a>

## CellStyle.operator !=(CellStyle, CellStyle) Operator

Implements the operator op_Inequality.

```csharp
public static bool operator !=(CloudyWing.SpreadsheetExporter.CellStyle left, CloudyWing.SpreadsheetExporter.CellStyle right);
```
#### Parameters

<a name='CloudyWing.SpreadsheetExporter.CellStyle.op_Inequality(CloudyWing.SpreadsheetExporter.CellStyle,CloudyWing.SpreadsheetExporter.CellStyle).left'></a>

`left` [CellStyle](CloudyWing.SpreadsheetExporter.CellStyle.md 'CloudyWing.SpreadsheetExporter.CellStyle')

The left.

<a name='CloudyWing.SpreadsheetExporter.CellStyle.op_Inequality(CloudyWing.SpreadsheetExporter.CellStyle,CloudyWing.SpreadsheetExporter.CellStyle).right'></a>

`right` [CellStyle](CloudyWing.SpreadsheetExporter.CellStyle.md 'CloudyWing.SpreadsheetExporter.CellStyle')

The right.

#### Returns
[System.Boolean](https://docs.microsoft.com/en-us/dotnet/api/System.Boolean 'System.Boolean')  
The result of the operator.