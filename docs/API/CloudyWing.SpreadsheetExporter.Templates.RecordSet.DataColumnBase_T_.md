#### [CloudyWing\.SpreadsheetExporter](index.md 'index')
### [CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet](CloudyWing.SpreadsheetExporter.Templates.RecordSet.md 'CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet')

## DataColumnBase\<T\> Class

The data column base\.

```csharp
public abstract class DataColumnBase<T>
```
#### Type parameters

<a name='CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnBase_T_.T'></a>

`T`

The type of the record\.

Inheritance [System\.Object](https://learn.microsoft.com/en-us/dotnet/api/system.object 'System\.Object') &#129106; DataColumnBase\<T\>

Derived  
&#8627; [DataColumn&lt;TRecord,TField&gt;](CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumn_TRecord,TField_.md 'CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.DataColumn\<TRecord,TField\>')  
&#8627; [RecordDataColumn&lt;T&gt;](CloudyWing.SpreadsheetExporter.Templates.RecordSet.RecordDataColumn_T_.md 'CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.RecordDataColumn\<T\>')
### Properties

<a name='CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnBase_T_.ChildColumns'></a>

## DataColumnBase\<T\>\.ChildColumns Property

Gets the child columns\.

```csharp
public CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection<T> ChildColumns { get; }
```

#### Property Value
[CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.DataColumnCollection&lt;](CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.md 'CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.DataColumnCollection\<T\>')[T](CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnBase_T_.md#CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnBase_T_.T 'CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.DataColumnBase\<T\>\.T')[&gt;](CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.md 'CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.DataColumnCollection\<T\>')  
The child columns\.

<a name='CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnBase_T_.ColumnLayers'></a>

## DataColumnBase\<T\>\.ColumnLayers Property

Gets the column layers\.

```csharp
public int ColumnLayers { get; }
```

#### Property Value
[System\.Int32](https://learn.microsoft.com/en-us/dotnet/api/system.int32 'System\.Int32')  
The column layers\.

### Remarks
自己和子層共幾層 Column，用來計算 RowSpan

<a name='CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnBase_T_.ColumnSpan'></a>

## DataColumnBase\<T\>\.ColumnSpan Property

Gets the column span\.

```csharp
public int ColumnSpan { get; }
```

#### Property Value
[System\.Int32](https://learn.microsoft.com/en-us/dotnet/api/system.int32 'System\.Int32')  
The column span\.

<a name='CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnBase_T_.HeaderStyle'></a>

## DataColumnBase\<T\>\.HeaderStyle Property

Gets or sets the header style\.

```csharp
public CloudyWing.SpreadsheetExporter.CellStyle HeaderStyle { get; set; }
```

#### Property Value
[CellStyle](CloudyWing.SpreadsheetExporter.CellStyle.md 'CloudyWing\.SpreadsheetExporter\.CellStyle')  
The header style\.

<a name='CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnBase_T_.HeaderText'></a>

## DataColumnBase\<T\>\.HeaderText Property

Gets or sets the header text\.

```csharp
public string HeaderText { get; set; }
```

#### Property Value
[System\.String](https://learn.microsoft.com/en-us/dotnet/api/system.string 'System\.String')  
The header text\.

<a name='CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnBase_T_.ParentColumns'></a>

## DataColumnBase\<T\>\.ParentColumns Property

Gets the parent columns\.

```csharp
public CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection<T> ParentColumns { get; internal set; }
```

#### Property Value
[CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.DataColumnCollection&lt;](CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.md 'CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.DataColumnCollection\<T\>')[T](CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnBase_T_.md#CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnBase_T_.T 'CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.DataColumnBase\<T\>\.T')[&gt;](CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.md 'CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.DataColumnCollection\<T\>')  
The parent columns\.

<a name='CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnBase_T_.Point'></a>

## DataColumnBase\<T\>\.Point Property

Gets the point\.

```csharp
public System.Drawing.Point Point { get; internal set; }
```

#### Property Value
[System\.Drawing\.Point](https://learn.microsoft.com/en-us/dotnet/api/system.drawing.point 'System\.Drawing\.Point')  
The point\.

<a name='CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnBase_T_.RowSpan'></a>

## DataColumnBase\<T\>\.RowSpan Property

Gets the row span\.

```csharp
public int RowSpan { get; }
```

#### Property Value
[System\.Int32](https://learn.microsoft.com/en-us/dotnet/api/system.int32 'System\.Int32')  
The row span\.
### Methods

<a name='CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnBase_T_.GetFieldFormula(CloudyWing.SpreadsheetExporter.Templates.RecordSet.RecordContext_T_)'></a>

## DataColumnBase\<T\>\.GetFieldFormula\(RecordContext\<T\>\) Method

Gets the field formula\.

```csharp
public abstract string GetFieldFormula(CloudyWing.SpreadsheetExporter.Templates.RecordSet.RecordContext<T> context);
```
#### Parameters

<a name='CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnBase_T_.GetFieldFormula(CloudyWing.SpreadsheetExporter.Templates.RecordSet.RecordContext_T_).context'></a>

`context` [CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.RecordContext&lt;](CloudyWing.SpreadsheetExporter.Templates.RecordSet.RecordContext_T_.md 'CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.RecordContext\<T\>')[T](CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnBase_T_.md#CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnBase_T_.T 'CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.DataColumnBase\<T\>\.T')[&gt;](CloudyWing.SpreadsheetExporter.Templates.RecordSet.RecordContext_T_.md 'CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.RecordContext\<T\>')

The context\.

#### Returns
[System\.String](https://learn.microsoft.com/en-us/dotnet/api/system.string 'System\.String')  
Returns the field formula\.

<a name='CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnBase_T_.GetFieldStyle(CloudyWing.SpreadsheetExporter.Templates.RecordSet.RecordContext_T_)'></a>

## DataColumnBase\<T\>\.GetFieldStyle\(RecordContext\<T\>\) Method

Gets the field style\.

```csharp
public abstract CloudyWing.SpreadsheetExporter.CellStyle GetFieldStyle(CloudyWing.SpreadsheetExporter.Templates.RecordSet.RecordContext<T> context);
```
#### Parameters

<a name='CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnBase_T_.GetFieldStyle(CloudyWing.SpreadsheetExporter.Templates.RecordSet.RecordContext_T_).context'></a>

`context` [CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.RecordContext&lt;](CloudyWing.SpreadsheetExporter.Templates.RecordSet.RecordContext_T_.md 'CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.RecordContext\<T\>')[T](CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnBase_T_.md#CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnBase_T_.T 'CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.DataColumnBase\<T\>\.T')[&gt;](CloudyWing.SpreadsheetExporter.Templates.RecordSet.RecordContext_T_.md 'CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.RecordContext\<T\>')

The context\.

#### Returns
[CellStyle](CloudyWing.SpreadsheetExporter.CellStyle.md 'CloudyWing\.SpreadsheetExporter\.CellStyle')  
Returns the field style\.

<a name='CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnBase_T_.GetFieldValue(CloudyWing.SpreadsheetExporter.Templates.RecordSet.RecordContext_T_)'></a>

## DataColumnBase\<T\>\.GetFieldValue\(RecordContext\<T\>\) Method

Gets the field value\.

```csharp
public abstract object GetFieldValue(CloudyWing.SpreadsheetExporter.Templates.RecordSet.RecordContext<T> context);
```
#### Parameters

<a name='CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnBase_T_.GetFieldValue(CloudyWing.SpreadsheetExporter.Templates.RecordSet.RecordContext_T_).context'></a>

`context` [CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.RecordContext&lt;](CloudyWing.SpreadsheetExporter.Templates.RecordSet.RecordContext_T_.md 'CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.RecordContext\<T\>')[T](CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnBase_T_.md#CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnBase_T_.T 'CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.DataColumnBase\<T\>\.T')[&gt;](CloudyWing.SpreadsheetExporter.Templates.RecordSet.RecordContext_T_.md 'CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.RecordContext\<T\>')

The context\.

#### Returns
[System\.Object](https://learn.microsoft.com/en-us/dotnet/api/system.object 'System\.Object')  
Returns the field value\.