#### [CloudyWing.SpreadsheetExporter](index.md 'index')
### [CloudyWing.SpreadsheetExporter](CloudyWing.SpreadsheetExporter.md 'CloudyWing.SpreadsheetExporter')

## CellFont Struct

The cell font.

```csharp
public struct CellFont :
System.IEquatable<CloudyWing.SpreadsheetExporter.CellFont>
```

Implements [System.IEquatable&lt;](https://docs.microsoft.com/en-us/dotnet/api/System.IEquatable-1 'System.IEquatable`1')[CellFont](CloudyWing.SpreadsheetExporter.CellFont.md 'CloudyWing.SpreadsheetExporter.CellFont')[&gt;](https://docs.microsoft.com/en-us/dotnet/api/System.IEquatable-1 'System.IEquatable`1')
### Constructors

<a name='CloudyWing.SpreadsheetExporter.CellFont.CellFont(string,short,System.Nullable_System.Drawing.Color_,CloudyWing.SpreadsheetExporter.FontStyles)'></a>

## CellFont(string, short, Nullable<Color>, FontStyles) Constructor

Initializes a new instance of the [CellFont](CloudyWing.SpreadsheetExporter.CellFont.md 'CloudyWing.SpreadsheetExporter.CellFont') struct.

```csharp
public CellFont(string name, short size=0, System.Nullable<System.Drawing.Color> color=null, CloudyWing.SpreadsheetExporter.FontStyles style=CloudyWing.SpreadsheetExporter.FontStyles.None);
```
#### Parameters

<a name='CloudyWing.SpreadsheetExporter.CellFont.CellFont(string,short,System.Nullable_System.Drawing.Color_,CloudyWing.SpreadsheetExporter.FontStyles).name'></a>

`name` [System.String](https://docs.microsoft.com/en-us/dotnet/api/System.String 'System.String')

The font name.

<a name='CloudyWing.SpreadsheetExporter.CellFont.CellFont(string,short,System.Nullable_System.Drawing.Color_,CloudyWing.SpreadsheetExporter.FontStyles).size'></a>

`size` [System.Int16](https://docs.microsoft.com/en-us/dotnet/api/System.Int16 'System.Int16')

The size.

<a name='CloudyWing.SpreadsheetExporter.CellFont.CellFont(string,short,System.Nullable_System.Drawing.Color_,CloudyWing.SpreadsheetExporter.FontStyles).color'></a>

`color` [System.Nullable&lt;](https://docs.microsoft.com/en-us/dotnet/api/System.Nullable-1 'System.Nullable`1')[System.Drawing.Color](https://docs.microsoft.com/en-us/dotnet/api/System.Drawing.Color 'System.Drawing.Color')[&gt;](https://docs.microsoft.com/en-us/dotnet/api/System.Nullable-1 'System.Nullable`1')

The color.

<a name='CloudyWing.SpreadsheetExporter.CellFont.CellFont(string,short,System.Nullable_System.Drawing.Color_,CloudyWing.SpreadsheetExporter.FontStyles).style'></a>

`style` [FontStyles](CloudyWing.SpreadsheetExporter.FontStyles.md 'CloudyWing.SpreadsheetExporter.FontStyles')

The style.
### Fields

<a name='CloudyWing.SpreadsheetExporter.CellFont.Empty'></a>

## CellFont.Empty Field

The cell font equals to `new CellFont()`.

```csharp
public static CellFont Empty;
```

#### Field Value
[CellFont](CloudyWing.SpreadsheetExporter.CellFont.md 'CloudyWing.SpreadsheetExporter.CellFont')
### Properties

<a name='CloudyWing.SpreadsheetExporter.CellFont.Color'></a>

## CellFont.Color Property

Gets the color. The default is `Color.Empty`.

```csharp
public System.Drawing.Color Color { get; set; }
```

#### Property Value
[System.Drawing.Color](https://docs.microsoft.com/en-us/dotnet/api/System.Drawing.Color 'System.Drawing.Color')  
The color.

<a name='CloudyWing.SpreadsheetExporter.CellFont.Name'></a>

## CellFont.Name Property

Gets the font name.

```csharp
public string Name { get; set; }
```

#### Property Value
[System.String](https://docs.microsoft.com/en-us/dotnet/api/System.String 'System.String')  
The font name.

<a name='CloudyWing.SpreadsheetExporter.CellFont.Size'></a>

## CellFont.Size Property

Gets the size.

```csharp
public short Size { get; set; }
```

#### Property Value
[System.Int16](https://docs.microsoft.com/en-us/dotnet/api/System.Int16 'System.Int16')  
The size.

<a name='CloudyWing.SpreadsheetExporter.CellFont.Style'></a>

## CellFont.Style Property

Gets the style.

```csharp
public CloudyWing.SpreadsheetExporter.FontStyles Style { get; set; }
```

#### Property Value
[FontStyles](CloudyWing.SpreadsheetExporter.FontStyles.md 'CloudyWing.SpreadsheetExporter.FontStyles')  
The style.
### Methods

<a name='CloudyWing.SpreadsheetExporter.CellFont.CloneAndSetColor(System.Drawing.Color)'></a>

## CellFont.CloneAndSetColor(Color) Method

Clones and set the color of new instance.

```csharp
public CloudyWing.SpreadsheetExporter.CellFont CloneAndSetColor(System.Drawing.Color color);
```
#### Parameters

<a name='CloudyWing.SpreadsheetExporter.CellFont.CloneAndSetColor(System.Drawing.Color).color'></a>

`color` [System.Drawing.Color](https://docs.microsoft.com/en-us/dotnet/api/System.Drawing.Color 'System.Drawing.Color')

The color.

#### Returns
[CellFont](CloudyWing.SpreadsheetExporter.CellFont.md 'CloudyWing.SpreadsheetExporter.CellFont')  
The cloned new cell.

<a name='CloudyWing.SpreadsheetExporter.CellFont.CloneAndSetName(string)'></a>

## CellFont.CloneAndSetName(string) Method

Clones and set the name of new instance.

```csharp
public CloudyWing.SpreadsheetExporter.CellFont CloneAndSetName(string name);
```
#### Parameters

<a name='CloudyWing.SpreadsheetExporter.CellFont.CloneAndSetName(string).name'></a>

`name` [System.String](https://docs.microsoft.com/en-us/dotnet/api/System.String 'System.String')

The name.

#### Returns
[CellFont](CloudyWing.SpreadsheetExporter.CellFont.md 'CloudyWing.SpreadsheetExporter.CellFont')  
The cloned new cell.

<a name='CloudyWing.SpreadsheetExporter.CellFont.CloneAndSetSize(short)'></a>

## CellFont.CloneAndSetSize(short) Method

Clones and set the size of new instance.

```csharp
public CloudyWing.SpreadsheetExporter.CellFont CloneAndSetSize(short size);
```
#### Parameters

<a name='CloudyWing.SpreadsheetExporter.CellFont.CloneAndSetSize(short).size'></a>

`size` [System.Int16](https://docs.microsoft.com/en-us/dotnet/api/System.Int16 'System.Int16')

The size.

#### Returns
[CellFont](CloudyWing.SpreadsheetExporter.CellFont.md 'CloudyWing.SpreadsheetExporter.CellFont')  
The cloned new cell.

<a name='CloudyWing.SpreadsheetExporter.CellFont.CloneAndSetStyle(CloudyWing.SpreadsheetExporter.FontStyles)'></a>

## CellFont.CloneAndSetStyle(FontStyles) Method

Clones and set the style of new instance.

```csharp
public CloudyWing.SpreadsheetExporter.CellFont CloneAndSetStyle(CloudyWing.SpreadsheetExporter.FontStyles style);
```
#### Parameters

<a name='CloudyWing.SpreadsheetExporter.CellFont.CloneAndSetStyle(CloudyWing.SpreadsheetExporter.FontStyles).style'></a>

`style` [FontStyles](CloudyWing.SpreadsheetExporter.FontStyles.md 'CloudyWing.SpreadsheetExporter.FontStyles')

The style.

#### Returns
[CellFont](CloudyWing.SpreadsheetExporter.CellFont.md 'CloudyWing.SpreadsheetExporter.CellFont')  
The cloned new cell.

<a name='CloudyWing.SpreadsheetExporter.CellFont.Equals(object)'></a>

## CellFont.Equals(object) Method

Indicates whether this instance and a specified object are equal.

```csharp
public override bool Equals(object obj);
```
#### Parameters

<a name='CloudyWing.SpreadsheetExporter.CellFont.Equals(object).obj'></a>

`obj` [System.Object](https://docs.microsoft.com/en-us/dotnet/api/System.Object 'System.Object')

Another object to compare to.

#### Returns
[System.Boolean](https://docs.microsoft.com/en-us/dotnet/api/System.Boolean 'System.Boolean')  
true if [obj](CloudyWing.SpreadsheetExporter.CellFont.md#CloudyWing.SpreadsheetExporter.CellFont.Equals(object).obj 'CloudyWing.SpreadsheetExporter.CellFont.Equals(object).obj') and this instance are the same type and represent the same value; otherwise, false.

<a name='CloudyWing.SpreadsheetExporter.CellFont.GetHashCode()'></a>

## CellFont.GetHashCode() Method

Returns the hash code for this instance.

```csharp
public override int GetHashCode();
```

#### Returns
[System.Int32](https://docs.microsoft.com/en-us/dotnet/api/System.Int32 'System.Int32')  
A 32-bit signed integer that is the hash code for this instance.
### Operators

<a name='CloudyWing.SpreadsheetExporter.CellFont.op_Equality(CloudyWing.SpreadsheetExporter.CellFont,CloudyWing.SpreadsheetExporter.CellFont)'></a>

## CellFont.operator ==(CellFont, CellFont) Operator

Implements the operator op_Equality.

```csharp
public static bool operator ==(CloudyWing.SpreadsheetExporter.CellFont left, CloudyWing.SpreadsheetExporter.CellFont right);
```
#### Parameters

<a name='CloudyWing.SpreadsheetExporter.CellFont.op_Equality(CloudyWing.SpreadsheetExporter.CellFont,CloudyWing.SpreadsheetExporter.CellFont).left'></a>

`left` [CellFont](CloudyWing.SpreadsheetExporter.CellFont.md 'CloudyWing.SpreadsheetExporter.CellFont')

The left.

<a name='CloudyWing.SpreadsheetExporter.CellFont.op_Equality(CloudyWing.SpreadsheetExporter.CellFont,CloudyWing.SpreadsheetExporter.CellFont).right'></a>

`right` [CellFont](CloudyWing.SpreadsheetExporter.CellFont.md 'CloudyWing.SpreadsheetExporter.CellFont')

The right.

#### Returns
[System.Boolean](https://docs.microsoft.com/en-us/dotnet/api/System.Boolean 'System.Boolean')  
The result of the operator.

<a name='CloudyWing.SpreadsheetExporter.CellFont.op_Inequality(CloudyWing.SpreadsheetExporter.CellFont,CloudyWing.SpreadsheetExporter.CellFont)'></a>

## CellFont.operator !=(CellFont, CellFont) Operator

Implements the operator op_Inequality.

```csharp
public static bool operator !=(CloudyWing.SpreadsheetExporter.CellFont left, CloudyWing.SpreadsheetExporter.CellFont right);
```
#### Parameters

<a name='CloudyWing.SpreadsheetExporter.CellFont.op_Inequality(CloudyWing.SpreadsheetExporter.CellFont,CloudyWing.SpreadsheetExporter.CellFont).left'></a>

`left` [CellFont](CloudyWing.SpreadsheetExporter.CellFont.md 'CloudyWing.SpreadsheetExporter.CellFont')

The left.

<a name='CloudyWing.SpreadsheetExporter.CellFont.op_Inequality(CloudyWing.SpreadsheetExporter.CellFont,CloudyWing.SpreadsheetExporter.CellFont).right'></a>

`right` [CellFont](CloudyWing.SpreadsheetExporter.CellFont.md 'CloudyWing.SpreadsheetExporter.CellFont')

The right.

#### Returns
[System.Boolean](https://docs.microsoft.com/en-us/dotnet/api/System.Boolean 'System.Boolean')  
The result of the operator.