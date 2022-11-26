#### [CloudyWing.SpreadsheetExporter](index.md 'index')
### [CloudyWing.SpreadsheetExporter.Templates.RecordSet](CloudyWing.SpreadsheetExporter.Templates.RecordSet.md 'CloudyWing.SpreadsheetExporter.Templates.RecordSet')

## FieldContext<TField,TRecoed> Class

The field context.

```csharp
public class FieldContext<TField,TRecoed> : CloudyWing.SpreadsheetExporter.Templates.RecordSet.RecordContext<TRecoed>
```
#### Type parameters

<a name='CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_TField,TRecoed_.TField'></a>

`TField`

The type of the field.

<a name='CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_TField,TRecoed_.TRecoed'></a>

`TRecoed`

The type of the recoed.

Inheritance [System.Object](https://docs.microsoft.com/en-us/dotnet/api/System.Object 'System.Object') &#129106; [CloudyWing.SpreadsheetExporter.Templates.RecordSet.RecordContext&lt;](CloudyWing.SpreadsheetExporter.Templates.RecordSet.RecordContext_T_.md 'CloudyWing.SpreadsheetExporter.Templates.RecordSet.RecordContext<T>')[TRecoed](CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_TField,TRecoed_.md#CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_TField,TRecoed_.TRecoed 'CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext<TField,TRecoed>.TRecoed')[&gt;](CloudyWing.SpreadsheetExporter.Templates.RecordSet.RecordContext_T_.md 'CloudyWing.SpreadsheetExporter.Templates.RecordSet.RecordContext<T>') &#129106; FieldContext<TField,TRecoed>
### Constructors

<a name='CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_TField,TRecoed_.FieldContext(CloudyWing.SpreadsheetExporter.Templates.RecordSet.RecordContext_TRecoed_,string,TField)'></a>

## FieldContext(RecordContext<TRecoed>, string, TField) Constructor

Initializes a new instance of the [FieldContext&lt;TField,TRecoed&gt;](CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_TField,TRecoed_.md 'CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext<TField,TRecoed>') class.

```csharp
public FieldContext(CloudyWing.SpreadsheetExporter.Templates.RecordSet.RecordContext<TRecoed> recordContext, string key, TField value);
```
#### Parameters

<a name='CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_TField,TRecoed_.FieldContext(CloudyWing.SpreadsheetExporter.Templates.RecordSet.RecordContext_TRecoed_,string,TField).recordContext'></a>

`recordContext` [CloudyWing.SpreadsheetExporter.Templates.RecordSet.RecordContext&lt;](CloudyWing.SpreadsheetExporter.Templates.RecordSet.RecordContext_T_.md 'CloudyWing.SpreadsheetExporter.Templates.RecordSet.RecordContext<T>')[TRecoed](CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_TField,TRecoed_.md#CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_TField,TRecoed_.TRecoed 'CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext<TField,TRecoed>.TRecoed')[&gt;](CloudyWing.SpreadsheetExporter.Templates.RecordSet.RecordContext_T_.md 'CloudyWing.SpreadsheetExporter.Templates.RecordSet.RecordContext<T>')

The record context.

<a name='CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_TField,TRecoed_.FieldContext(CloudyWing.SpreadsheetExporter.Templates.RecordSet.RecordContext_TRecoed_,string,TField).key'></a>

`key` [System.String](https://docs.microsoft.com/en-us/dotnet/api/System.String 'System.String')

The key.

<a name='CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_TField,TRecoed_.FieldContext(CloudyWing.SpreadsheetExporter.Templates.RecordSet.RecordContext_TRecoed_,string,TField).value'></a>

`value` [TField](CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_TField,TRecoed_.md#CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_TField,TRecoed_.TField 'CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext<TField,TRecoed>.TField')

The value.

#### Exceptions

[System.ArgumentNullException](https://docs.microsoft.com/en-us/dotnet/api/System.ArgumentNullException 'System.ArgumentNullException')  
key
### Properties

<a name='CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_TField,TRecoed_.Key'></a>

## FieldContext<TField,TRecoed>.Key Property

Gets the key.

```csharp
public string Key { get; }
```

#### Property Value
[System.String](https://docs.microsoft.com/en-us/dotnet/api/System.String 'System.String')  
The key.

<a name='CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_TField,TRecoed_.Value'></a>

## FieldContext<TField,TRecoed>.Value Property

Gets the value.

```csharp
public TField Value { get; }
```

#### Property Value
[TField](CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_TField,TRecoed_.md#CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_TField,TRecoed_.TField 'CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext<TField,TRecoed>.TField')  
The value.