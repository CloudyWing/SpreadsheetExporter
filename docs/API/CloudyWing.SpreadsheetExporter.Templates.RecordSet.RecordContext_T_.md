#### [CloudyWing.SpreadsheetExporter](index.md 'index')
### [CloudyWing.SpreadsheetExporter.Templates.RecordSet](CloudyWing.SpreadsheetExporter.Templates.RecordSet.md 'CloudyWing.SpreadsheetExporter.Templates.RecordSet')

## RecordContext<T> Class

The record context.

```csharp
public class RecordContext<T>
```
#### Type parameters

<a name='CloudyWing.SpreadsheetExporter.Templates.RecordSet.RecordContext_T_.T'></a>

`T`

The type of the record.

Inheritance [System.Object](https://docs.microsoft.com/en-us/dotnet/api/System.Object 'System.Object') &#129106; RecordContext<T>

Derived  
&#8627; [FieldContext&lt;TRecord,TField&gt;](CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_TRecord,TField_.md 'CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext<TRecord,TField>')
### Constructors

<a name='CloudyWing.SpreadsheetExporter.Templates.RecordSet.RecordContext_T_.RecordContext(int,int,T)'></a>

## RecordContext(int, int, T) Constructor

The record context.

```csharp
public RecordContext(int cellIndex, int rowIndex, T record);
```
#### Parameters

<a name='CloudyWing.SpreadsheetExporter.Templates.RecordSet.RecordContext_T_.RecordContext(int,int,T).cellIndex'></a>

`cellIndex` [System.Int32](https://docs.microsoft.com/en-us/dotnet/api/System.Int32 'System.Int32')

Index of the cell.

<a name='CloudyWing.SpreadsheetExporter.Templates.RecordSet.RecordContext_T_.RecordContext(int,int,T).rowIndex'></a>

`rowIndex` [System.Int32](https://docs.microsoft.com/en-us/dotnet/api/System.Int32 'System.Int32')

Index of the row.

<a name='CloudyWing.SpreadsheetExporter.Templates.RecordSet.RecordContext_T_.RecordContext(int,int,T).record'></a>

`record` [T](CloudyWing.SpreadsheetExporter.Templates.RecordSet.RecordContext_T_.md#CloudyWing.SpreadsheetExporter.Templates.RecordSet.RecordContext_T_.T 'CloudyWing.SpreadsheetExporter.Templates.RecordSet.RecordContext<T>.T')

The record.
### Properties

<a name='CloudyWing.SpreadsheetExporter.Templates.RecordSet.RecordContext_T_.CellIndex'></a>

## RecordContext<T>.CellIndex Property

Gets the index of the cell. The index start at 0.

```csharp
public int CellIndex { get; }
```

#### Property Value
[System.Int32](https://docs.microsoft.com/en-us/dotnet/api/System.Int32 'System.Int32')  
The index of the cell.

<a name='CloudyWing.SpreadsheetExporter.Templates.RecordSet.RecordContext_T_.Record'></a>

## RecordContext<T>.Record Property

Gets the record.

```csharp
public T Record { get; }
```

#### Property Value
[T](CloudyWing.SpreadsheetExporter.Templates.RecordSet.RecordContext_T_.md#CloudyWing.SpreadsheetExporter.Templates.RecordSet.RecordContext_T_.T 'CloudyWing.SpreadsheetExporter.Templates.RecordSet.RecordContext<T>.T')  
The record.

<a name='CloudyWing.SpreadsheetExporter.Templates.RecordSet.RecordContext_T_.RowIndex'></a>

## RecordContext<T>.RowIndex Property

Gets the index of the row. The index start at 0.

```csharp
public int RowIndex { get; }
```

#### Property Value
[System.Int32](https://docs.microsoft.com/en-us/dotnet/api/System.Int32 'System.Int32')  
The index of the row.