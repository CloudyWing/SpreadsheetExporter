#### [CloudyWing.SpreadsheetExporter](index.md 'index')
### [CloudyWing.SpreadsheetExporter.Templates.RecordSet](CloudyWing.SpreadsheetExporter.Templates.RecordSet.md 'CloudyWing.SpreadsheetExporter.Templates.RecordSet')

## FieldContext<TRecoed,TField> Class

The field context.

```csharp
public class FieldContext<TRecoed,TField> : CloudyWing.SpreadsheetExporter.Templates.RecordSet.RecordContext<TRecoed>
```
#### Type parameters

<a name='CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_TRecoed,TField_.TRecoed'></a>

`TRecoed`

The type of the recoed.

<a name='CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_TRecoed,TField_.TField'></a>

`TField`

The type of the field.

Inheritance [System.Object](https://docs.microsoft.com/en-us/dotnet/api/System.Object 'System.Object') &#129106; [CloudyWing.SpreadsheetExporter.Templates.RecordSet.RecordContext&lt;](CloudyWing.SpreadsheetExporter.Templates.RecordSet.RecordContext_T_.md 'CloudyWing.SpreadsheetExporter.Templates.RecordSet.RecordContext<T>')[TRecoed](CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_TRecoed,TField_.md#CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_TRecoed,TField_.TRecoed 'CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext<TRecoed,TField>.TRecoed')[&gt;](CloudyWing.SpreadsheetExporter.Templates.RecordSet.RecordContext_T_.md 'CloudyWing.SpreadsheetExporter.Templates.RecordSet.RecordContext<T>') &#129106; FieldContext<TRecoed,TField>
### Constructors

<a name='CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_TRecoed,TField_.FieldContext(CloudyWing.SpreadsheetExporter.Templates.RecordSet.RecordContext_TRecoed_,string,TField)'></a>

## FieldContext(RecordContext<TRecoed>, string, TField) Constructor

Initializes a new instance of the [FieldContext&lt;TRecoed,TField&gt;](CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_TRecoed,TField_.md 'CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext<TRecoed,TField>') class.

```csharp
public FieldContext(CloudyWing.SpreadsheetExporter.Templates.RecordSet.RecordContext<TRecoed> recordContext, string key, TField value);
```
#### Parameters

<a name='CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_TRecoed,TField_.FieldContext(CloudyWing.SpreadsheetExporter.Templates.RecordSet.RecordContext_TRecoed_,string,TField).recordContext'></a>

`recordContext` [CloudyWing.SpreadsheetExporter.Templates.RecordSet.RecordContext&lt;](CloudyWing.SpreadsheetExporter.Templates.RecordSet.RecordContext_T_.md 'CloudyWing.SpreadsheetExporter.Templates.RecordSet.RecordContext<T>')[TRecoed](CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_TRecoed,TField_.md#CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_TRecoed,TField_.TRecoed 'CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext<TRecoed,TField>.TRecoed')[&gt;](CloudyWing.SpreadsheetExporter.Templates.RecordSet.RecordContext_T_.md 'CloudyWing.SpreadsheetExporter.Templates.RecordSet.RecordContext<T>')

The record context.

<a name='CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_TRecoed,TField_.FieldContext(CloudyWing.SpreadsheetExporter.Templates.RecordSet.RecordContext_TRecoed_,string,TField).key'></a>

`key` [System.String](https://docs.microsoft.com/en-us/dotnet/api/System.String 'System.String')

The key.

<a name='CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_TRecoed,TField_.FieldContext(CloudyWing.SpreadsheetExporter.Templates.RecordSet.RecordContext_TRecoed_,string,TField).value'></a>

`value` [TField](CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_TRecoed,TField_.md#CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_TRecoed,TField_.TField 'CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext<TRecoed,TField>.TField')

The value.

#### Exceptions

[System.ArgumentNullException](https://docs.microsoft.com/en-us/dotnet/api/System.ArgumentNullException 'System.ArgumentNullException')  
key
### Properties

<a name='CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_TRecoed,TField_.Key'></a>

## FieldContext<TRecoed,TField>.Key Property

Gets the key.

```csharp
public string Key { get; }
```

#### Property Value
[System.String](https://docs.microsoft.com/en-us/dotnet/api/System.String 'System.String')  
The key.

<a name='CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_TRecoed,TField_.Value'></a>

## FieldContext<TRecoed,TField>.Value Property

Gets the value.

```csharp
public TField Value { get; }
```

#### Property Value
[TField](CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_TRecoed,TField_.md#CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_TRecoed,TField_.TField 'CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext<TRecoed,TField>.TField')  
The value.