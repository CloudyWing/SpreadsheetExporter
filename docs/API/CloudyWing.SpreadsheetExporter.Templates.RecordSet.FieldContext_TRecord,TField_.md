#### [CloudyWing.SpreadsheetExporter](index.md 'index')
### [CloudyWing.SpreadsheetExporter.Templates.RecordSet](CloudyWing.SpreadsheetExporter.Templates.RecordSet.md 'CloudyWing.SpreadsheetExporter.Templates.RecordSet')

## FieldContext<TRecord,TField> Class

The field context.

```csharp
public class FieldContext<TRecord,TField> : CloudyWing.SpreadsheetExporter.Templates.RecordSet.RecordContext<TRecord>
```
#### Type parameters

<a name='CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_TRecord,TField_.TRecord'></a>

`TRecord`

The type of the record.

<a name='CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_TRecord,TField_.TField'></a>

`TField`

The type of the field.

Inheritance [System.Object](https://docs.microsoft.com/en-us/dotnet/api/System.Object 'System.Object') &#129106; [CloudyWing.SpreadsheetExporter.Templates.RecordSet.RecordContext&lt;](CloudyWing.SpreadsheetExporter.Templates.RecordSet.RecordContext_T_.md 'CloudyWing.SpreadsheetExporter.Templates.RecordSet.RecordContext<T>')[TRecord](CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_TRecord,TField_.md#CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_TRecord,TField_.TRecord 'CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext<TRecord,TField>.TRecord')[&gt;](CloudyWing.SpreadsheetExporter.Templates.RecordSet.RecordContext_T_.md 'CloudyWing.SpreadsheetExporter.Templates.RecordSet.RecordContext<T>') &#129106; FieldContext<TRecord,TField>

#### Exceptions

[System.ArgumentNullException](https://docs.microsoft.com/en-us/dotnet/api/System.ArgumentNullException 'System.ArgumentNullException')  
key
### Constructors

<a name='CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_TRecord,TField_.FieldContext(CloudyWing.SpreadsheetExporter.Templates.RecordSet.RecordContext_TRecord_,string,TField)'></a>

## FieldContext(RecordContext<TRecord>, string, TField) Constructor

The field context.

```csharp
public FieldContext(CloudyWing.SpreadsheetExporter.Templates.RecordSet.RecordContext<TRecord> recordContext, string key, TField value);
```
#### Parameters

<a name='CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_TRecord,TField_.FieldContext(CloudyWing.SpreadsheetExporter.Templates.RecordSet.RecordContext_TRecord_,string,TField).recordContext'></a>

`recordContext` [CloudyWing.SpreadsheetExporter.Templates.RecordSet.RecordContext&lt;](CloudyWing.SpreadsheetExporter.Templates.RecordSet.RecordContext_T_.md 'CloudyWing.SpreadsheetExporter.Templates.RecordSet.RecordContext<T>')[TRecord](CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_TRecord,TField_.md#CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_TRecord,TField_.TRecord 'CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext<TRecord,TField>.TRecord')[&gt;](CloudyWing.SpreadsheetExporter.Templates.RecordSet.RecordContext_T_.md 'CloudyWing.SpreadsheetExporter.Templates.RecordSet.RecordContext<T>')

The record context.

<a name='CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_TRecord,TField_.FieldContext(CloudyWing.SpreadsheetExporter.Templates.RecordSet.RecordContext_TRecord_,string,TField).key'></a>

`key` [System.String](https://docs.microsoft.com/en-us/dotnet/api/System.String 'System.String')

The key.

<a name='CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_TRecord,TField_.FieldContext(CloudyWing.SpreadsheetExporter.Templates.RecordSet.RecordContext_TRecord_,string,TField).value'></a>

`value` [TField](CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_TRecord,TField_.md#CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_TRecord,TField_.TField 'CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext<TRecord,TField>.TField')

The value.

#### Exceptions

[System.ArgumentNullException](https://docs.microsoft.com/en-us/dotnet/api/System.ArgumentNullException 'System.ArgumentNullException')  
key
### Properties

<a name='CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_TRecord,TField_.Key'></a>

## FieldContext<TRecord,TField>.Key Property

Gets the key.

```csharp
public string Key { get; }
```

#### Property Value
[System.String](https://docs.microsoft.com/en-us/dotnet/api/System.String 'System.String')  
The key.

<a name='CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_TRecord,TField_.Value'></a>

## FieldContext<TRecord,TField>.Value Property

Gets the value.

```csharp
public TField Value { get; }
```

#### Property Value
[TField](CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_TRecord,TField_.md#CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_TRecord,TField_.TField 'CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext<TRecord,TField>.TField')  
The value.