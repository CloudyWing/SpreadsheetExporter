#### [CloudyWing.SpreadsheetExporter](index.md 'index')
### [CloudyWing.SpreadsheetExporter.Templates.RecordSet](CloudyWing.SpreadsheetExporter.Templates.RecordSet.md 'CloudyWing.SpreadsheetExporter.Templates.RecordSet')

## DataColumn<TField,TRecord> Class

The data column.

```csharp
internal class DataColumn<TField,TRecord> : CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnBase<TRecord>
```
#### Type parameters

<a name='CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumn_TField,TRecord_.TField'></a>

`TField`

The type of the record field.

<a name='CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumn_TField,TRecord_.TRecord'></a>

`TRecord`

The type of the record.

Inheritance [System.Object](https://docs.microsoft.com/en-us/dotnet/api/System.Object 'System.Object') &#129106; [CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnBase&lt;](CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnBase_T_.md 'CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnBase<T>')[TRecord](CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumn_TField,TRecord_.md#CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumn_TField,TRecord_.TRecord 'CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumn<TField,TRecord>.TRecord')[&gt;](CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnBase_T_.md 'CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnBase<T>') &#129106; DataColumn<TField,TRecord>
### Constructors

<a name='CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumn_TField,TRecord_.DataColumn(string)'></a>

## DataColumn(string) Constructor

Initializes a new instance of the [DataColumn&lt;TField,TRecord&gt;](CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumn_TField,TRecord_.md 'CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumn<TField,TRecord>') class.

```csharp
public DataColumn(string fieldKey);
```
#### Parameters

<a name='CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumn_TField,TRecord_.DataColumn(string).fieldKey'></a>

`fieldKey` [System.String](https://docs.microsoft.com/en-us/dotnet/api/System.String 'System.String')

The field key.

<a name='CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumn_TField,TRecord_.DataColumn(System.Linq.Expressions.Expression_System.Func_TRecord,TField__)'></a>

## DataColumn(Expression<Func<TRecord,TField>>) Constructor

Initializes a new instance of the [DataColumn&lt;TField,TRecord&gt;](CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumn_TField,TRecord_.md 'CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumn<TField,TRecord>') class.

```csharp
public DataColumn(System.Linq.Expressions.Expression<System.Func<TRecord,TField>> fieldKeyExpression);
```
#### Parameters

<a name='CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumn_TField,TRecord_.DataColumn(System.Linq.Expressions.Expression_System.Func_TRecord,TField__).fieldKeyExpression'></a>

`fieldKeyExpression` [System.Linq.Expressions.Expression&lt;](https://docs.microsoft.com/en-us/dotnet/api/System.Linq.Expressions.Expression-1 'System.Linq.Expressions.Expression`1')[System.Func&lt;](https://docs.microsoft.com/en-us/dotnet/api/System.Func-2 'System.Func`2')[TRecord](CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumn_TField,TRecord_.md#CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumn_TField,TRecord_.TRecord 'CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumn<TField,TRecord>.TRecord')[,](https://docs.microsoft.com/en-us/dotnet/api/System.Func-2 'System.Func`2')[TField](CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumn_TField,TRecord_.md#CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumn_TField,TRecord_.TField 'CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumn<TField,TRecord>.TField')[&gt;](https://docs.microsoft.com/en-us/dotnet/api/System.Func-2 'System.Func`2')[&gt;](https://docs.microsoft.com/en-us/dotnet/api/System.Linq.Expressions.Expression-1 'System.Linq.Expressions.Expression`1')

The field key expression.
### Properties

<a name='CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumn_TField,TRecord_.FieldFormulaGenerator'></a>

## DataColumn<TField,TRecord>.FieldFormulaGenerator Property

Gets or sets the field formula generator.

```csharp
public System.Func<CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext<TField,TRecord>,string> FieldFormulaGenerator { get; set; }
```

#### Property Value
[System.Func&lt;](https://docs.microsoft.com/en-us/dotnet/api/System.Func-2 'System.Func`2')[CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext&lt;](CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_TField,TRecoed_.md 'CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext<TField,TRecoed>')[TField](CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumn_TField,TRecord_.md#CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumn_TField,TRecord_.TField 'CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumn<TField,TRecord>.TField')[,](CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_TField,TRecoed_.md 'CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext<TField,TRecoed>')[TRecord](CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumn_TField,TRecord_.md#CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumn_TField,TRecord_.TRecord 'CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumn<TField,TRecord>.TRecord')[&gt;](CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_TField,TRecoed_.md 'CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext<TField,TRecoed>')[,](https://docs.microsoft.com/en-us/dotnet/api/System.Func-2 'System.Func`2')[System.String](https://docs.microsoft.com/en-us/dotnet/api/System.String 'System.String')[&gt;](https://docs.microsoft.com/en-us/dotnet/api/System.Func-2 'System.Func`2')  
The field formula generator.

<a name='CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumn_TField,TRecord_.FieldKey'></a>

## DataColumn<TField,TRecord>.FieldKey Property

Gets the field key.

```csharp
public string FieldKey { get; }
```

#### Property Value
[System.String](https://docs.microsoft.com/en-us/dotnet/api/System.String 'System.String')  
The field key.

<a name='CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumn_TField,TRecord_.FieldStyleGenerator'></a>

## DataColumn<TField,TRecord>.FieldStyleGenerator Property

Gets or sets the field style generator.

```csharp
public System.Func<CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext<TField,TRecord>,CloudyWing.SpreadsheetExporter.CellStyle> FieldStyleGenerator { get; set; }
```

#### Property Value
[System.Func&lt;](https://docs.microsoft.com/en-us/dotnet/api/System.Func-2 'System.Func`2')[CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext&lt;](CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_TField,TRecoed_.md 'CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext<TField,TRecoed>')[TField](CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumn_TField,TRecord_.md#CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumn_TField,TRecord_.TField 'CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumn<TField,TRecord>.TField')[,](CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_TField,TRecoed_.md 'CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext<TField,TRecoed>')[TRecord](CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumn_TField,TRecord_.md#CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumn_TField,TRecord_.TRecord 'CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumn<TField,TRecord>.TRecord')[&gt;](CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_TField,TRecoed_.md 'CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext<TField,TRecoed>')[,](https://docs.microsoft.com/en-us/dotnet/api/System.Func-2 'System.Func`2')[CellStyle](CloudyWing.SpreadsheetExporter.CellStyle.md 'CloudyWing.SpreadsheetExporter.CellStyle')[&gt;](https://docs.microsoft.com/en-us/dotnet/api/System.Func-2 'System.Func`2')  
The field style generator.

<a name='CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumn_TField,TRecord_.FieldValueGenerator'></a>

## DataColumn<TField,TRecord>.FieldValueGenerator Property

Gets the field value generator.

```csharp
public System.Func<CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext<TField,TRecord>,object> FieldValueGenerator { get; set; }
```

#### Property Value
[System.Func&lt;](https://docs.microsoft.com/en-us/dotnet/api/System.Func-2 'System.Func`2')[CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext&lt;](CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_TField,TRecoed_.md 'CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext<TField,TRecoed>')[TField](CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumn_TField,TRecord_.md#CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumn_TField,TRecord_.TField 'CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumn<TField,TRecord>.TField')[,](CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_TField,TRecoed_.md 'CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext<TField,TRecoed>')[TRecord](CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumn_TField,TRecord_.md#CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumn_TField,TRecord_.TRecord 'CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumn<TField,TRecord>.TRecord')[&gt;](CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_TField,TRecoed_.md 'CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext<TField,TRecoed>')[,](https://docs.microsoft.com/en-us/dotnet/api/System.Func-2 'System.Func`2')[System.Object](https://docs.microsoft.com/en-us/dotnet/api/System.Object 'System.Object')[&gt;](https://docs.microsoft.com/en-us/dotnet/api/System.Func-2 'System.Func`2')  
The field value generator.