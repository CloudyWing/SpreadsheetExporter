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
&#8627; [FieldContext&lt;TField,TRecoed&gt;](CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_TField,TRecoed_.md 'CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext<TField,TRecoed>')
### Constructors

<a name='CloudyWing.SpreadsheetExporter.Templates.RecordSet.RecordContext_T_.RecordContext(int,int,T)'></a>

## RecordContext(int, int, T) Constructor

Initializes a new instance of the [RecordContext&lt;T&gt;](CloudyWing.SpreadsheetExporter.Templates.RecordSet.RecordContext_T_.md 'CloudyWing.SpreadsheetExporter.Templates.RecordSet.RecordContext<T>') class.

```csharp
public RecordContext(int headerRowSpan, int recordIndex, T record);
```
#### Parameters

<a name='CloudyWing.SpreadsheetExporter.Templates.RecordSet.RecordContext_T_.RecordContext(int,int,T).headerRowSpan'></a>

`headerRowSpan` [System.Int32](https://docs.microsoft.com/en-us/dotnet/api/System.Int32 'System.Int32')

The header row span.

<a name='CloudyWing.SpreadsheetExporter.Templates.RecordSet.RecordContext_T_.RecordContext(int,int,T).recordIndex'></a>

`recordIndex` [System.Int32](https://docs.microsoft.com/en-us/dotnet/api/System.Int32 'System.Int32')

Index of the record.

<a name='CloudyWing.SpreadsheetExporter.Templates.RecordSet.RecordContext_T_.RecordContext(int,int,T).record'></a>

`record` [T](CloudyWing.SpreadsheetExporter.Templates.RecordSet.RecordContext_T_.md#CloudyWing.SpreadsheetExporter.Templates.RecordSet.RecordContext_T_.T 'CloudyWing.SpreadsheetExporter.Templates.RecordSet.RecordContext<T>.T')

The record.
### Properties

<a name='CloudyWing.SpreadsheetExporter.Templates.RecordSet.RecordContext_T_.HeaderRowSpan'></a>

## RecordContext<T>.HeaderRowSpan Property

Gets the header row span.

```csharp
public int HeaderRowSpan { get; }
```

#### Property Value
[System.Int32](https://docs.microsoft.com/en-us/dotnet/api/System.Int32 'System.Int32')  
The header row span.

<a name='CloudyWing.SpreadsheetExporter.Templates.RecordSet.RecordContext_T_.Record'></a>

## RecordContext<T>.Record Property

Gets the record.

```csharp
public T Record { get; }
```

#### Property Value
[T](CloudyWing.SpreadsheetExporter.Templates.RecordSet.RecordContext_T_.md#CloudyWing.SpreadsheetExporter.Templates.RecordSet.RecordContext_T_.T 'CloudyWing.SpreadsheetExporter.Templates.RecordSet.RecordContext<T>.T')  
The record.

<a name='CloudyWing.SpreadsheetExporter.Templates.RecordSet.RecordContext_T_.RecordIndex'></a>

## RecordContext<T>.RecordIndex Property

Gets the index of the record.

```csharp
public int RecordIndex { get; }
```

#### Property Value
[System.Int32](https://docs.microsoft.com/en-us/dotnet/api/System.Int32 'System.Int32')  
The index of the record.