#### [CloudyWing.SpreadsheetExporter](index.md 'index')
### [CloudyWing.SpreadsheetExporter.Templates.RecordSet](CloudyWing.SpreadsheetExporter.Templates.RecordSet.md 'CloudyWing.SpreadsheetExporter.Templates.RecordSet')

## DataColumn<TRecord,TField> Class

The data column.

```csharp
internal class DataColumn<TRecord,TField> : CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnBase<TRecord>
```
#### Type parameters

<a name='CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumn_TRecord,TField_.TRecord'></a>

`TRecord`

The type of the record.

<a name='CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumn_TRecord,TField_.TField'></a>

`TField`

The type of the field.

Inheritance [System.Object](https://docs.microsoft.com/en-us/dotnet/api/System.Object 'System.Object') &#129106; [CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnBase&lt;](CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnBase_T_.md 'CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnBase<T>')[TRecord](CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumn_TRecord,TField_.md#CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumn_TRecord,TField_.TRecord 'CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumn<TRecord,TField>.TRecord')[&gt;](CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnBase_T_.md 'CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnBase<T>') &#129106; DataColumn<TRecord,TField>

### See Also
- [DataColumnBase&lt;T&gt;](CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnBase_T_.md 'CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnBase<T>')
### Constructors

<a name='CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumn_TRecord,TField_.DataColumn(string)'></a>

## DataColumn(string) Constructor

Initializes a new instance of the [DataColumn&lt;TRecord,TField&gt;](CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumn_TRecord,TField_.md 'CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumn<TRecord,TField>') class.

```csharp
public DataColumn(string fieldKey);
```
#### Parameters

<a name='CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumn_TRecord,TField_.DataColumn(string).fieldKey'></a>

`fieldKey` [System.String](https://docs.microsoft.com/en-us/dotnet/api/System.String 'System.String')

The field key.

<a name='CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumn_TRecord,TField_.DataColumn(System.Linq.Expressions.Expression_System.Func_TRecord,TField__)'></a>

## DataColumn(Expression<Func<TRecord,TField>>) Constructor

Initializes a new instance of the [DataColumn&lt;TRecord,TField&gt;](CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumn_TRecord,TField_.md 'CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumn<TRecord,TField>') class.

```csharp
public DataColumn(System.Linq.Expressions.Expression<System.Func<TRecord,TField>> fieldKeyExpression);
```
#### Parameters

<a name='CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumn_TRecord,TField_.DataColumn(System.Linq.Expressions.Expression_System.Func_TRecord,TField__).fieldKeyExpression'></a>

`fieldKeyExpression` [System.Linq.Expressions.Expression&lt;](https://docs.microsoft.com/en-us/dotnet/api/System.Linq.Expressions.Expression-1 'System.Linq.Expressions.Expression`1')[System.Func&lt;](https://docs.microsoft.com/en-us/dotnet/api/System.Func-2 'System.Func`2')[TRecord](CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumn_TRecord,TField_.md#CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumn_TRecord,TField_.TRecord 'CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumn<TRecord,TField>.TRecord')[,](https://docs.microsoft.com/en-us/dotnet/api/System.Func-2 'System.Func`2')[TField](CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumn_TRecord,TField_.md#CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumn_TRecord,TField_.TField 'CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumn<TRecord,TField>.TField')[&gt;](https://docs.microsoft.com/en-us/dotnet/api/System.Func-2 'System.Func`2')[&gt;](https://docs.microsoft.com/en-us/dotnet/api/System.Linq.Expressions.Expression-1 'System.Linq.Expressions.Expression`1')

The field key expression.
### Properties

<a name='CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumn_TRecord,TField_.FieldFormulaGenerator'></a>

## DataColumn<TRecord,TField>.FieldFormulaGenerator Property

Gets or sets the field formula generator.

```csharp
public System.Func<CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext<TRecord,TField>,string> FieldFormulaGenerator { get; set; }
```

#### Property Value
[System.Func&lt;](https://docs.microsoft.com/en-us/dotnet/api/System.Func-2 'System.Func`2')[CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext&lt;](CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_TRecord,TField_.md 'CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext<TRecord,TField>')[TRecord](CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumn_TRecord,TField_.md#CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumn_TRecord,TField_.TRecord 'CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumn<TRecord,TField>.TRecord')[,](CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_TRecord,TField_.md 'CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext<TRecord,TField>')[TField](CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumn_TRecord,TField_.md#CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumn_TRecord,TField_.TField 'CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumn<TRecord,TField>.TField')[&gt;](CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_TRecord,TField_.md 'CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext<TRecord,TField>')[,](https://docs.microsoft.com/en-us/dotnet/api/System.Func-2 'System.Func`2')[System.String](https://docs.microsoft.com/en-us/dotnet/api/System.String 'System.String')[&gt;](https://docs.microsoft.com/en-us/dotnet/api/System.Func-2 'System.Func`2')  
The field formula generator.

<a name='CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumn_TRecord,TField_.FieldKey'></a>

## DataColumn<TRecord,TField>.FieldKey Property

Gets the field key.

```csharp
public string FieldKey { get; }
```

#### Property Value
[System.String](https://docs.microsoft.com/en-us/dotnet/api/System.String 'System.String')  
The field key.

<a name='CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumn_TRecord,TField_.FieldStyleGenerator'></a>

## DataColumn<TRecord,TField>.FieldStyleGenerator Property

Gets or sets the field style generator.

```csharp
public System.Func<CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext<TRecord,TField>,CloudyWing.SpreadsheetExporter.CellStyle> FieldStyleGenerator { get; set; }
```

#### Property Value
[System.Func&lt;](https://docs.microsoft.com/en-us/dotnet/api/System.Func-2 'System.Func`2')[CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext&lt;](CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_TRecord,TField_.md 'CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext<TRecord,TField>')[TRecord](CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumn_TRecord,TField_.md#CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumn_TRecord,TField_.TRecord 'CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumn<TRecord,TField>.TRecord')[,](CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_TRecord,TField_.md 'CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext<TRecord,TField>')[TField](CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumn_TRecord,TField_.md#CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumn_TRecord,TField_.TField 'CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumn<TRecord,TField>.TField')[&gt;](CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_TRecord,TField_.md 'CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext<TRecord,TField>')[,](https://docs.microsoft.com/en-us/dotnet/api/System.Func-2 'System.Func`2')[CellStyle](CloudyWing.SpreadsheetExporter.CellStyle.md 'CloudyWing.SpreadsheetExporter.CellStyle')[&gt;](https://docs.microsoft.com/en-us/dotnet/api/System.Func-2 'System.Func`2')  
The field style generator.

<a name='CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumn_TRecord,TField_.FieldValueGenerator'></a>

## DataColumn<TRecord,TField>.FieldValueGenerator Property

Gets the field value generator.

```csharp
public System.Func<CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext<TRecord,TField>,object> FieldValueGenerator { get; set; }
```

#### Property Value
[System.Func&lt;](https://docs.microsoft.com/en-us/dotnet/api/System.Func-2 'System.Func`2')[CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext&lt;](CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_TRecord,TField_.md 'CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext<TRecord,TField>')[TRecord](CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumn_TRecord,TField_.md#CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumn_TRecord,TField_.TRecord 'CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumn<TRecord,TField>.TRecord')[,](CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_TRecord,TField_.md 'CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext<TRecord,TField>')[TField](CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumn_TRecord,TField_.md#CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumn_TRecord,TField_.TField 'CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumn<TRecord,TField>.TField')[&gt;](CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_TRecord,TField_.md 'CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext<TRecord,TField>')[,](https://docs.microsoft.com/en-us/dotnet/api/System.Func-2 'System.Func`2')[System.Object](https://docs.microsoft.com/en-us/dotnet/api/System.Object 'System.Object')[&gt;](https://docs.microsoft.com/en-us/dotnet/api/System.Func-2 'System.Func`2')  
The field value generator.
### Methods

<a name='CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumn_TRecord,TField_.GetFieldFormula(CloudyWing.SpreadsheetExporter.Templates.RecordSet.RecordContext_TRecord_)'></a>

## DataColumn<TRecord,TField>.GetFieldFormula(RecordContext<TRecord>) Method

Gets the field formula.

```csharp
public override string GetFieldFormula(CloudyWing.SpreadsheetExporter.Templates.RecordSet.RecordContext<TRecord> recordContext);
```
#### Parameters

<a name='CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumn_TRecord,TField_.GetFieldFormula(CloudyWing.SpreadsheetExporter.Templates.RecordSet.RecordContext_TRecord_).recordContext'></a>

`recordContext` [CloudyWing.SpreadsheetExporter.Templates.RecordSet.RecordContext&lt;](CloudyWing.SpreadsheetExporter.Templates.RecordSet.RecordContext_T_.md 'CloudyWing.SpreadsheetExporter.Templates.RecordSet.RecordContext<T>')[TRecord](CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumn_TRecord,TField_.md#CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumn_TRecord,TField_.TRecord 'CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumn<TRecord,TField>.TRecord')[&gt;](CloudyWing.SpreadsheetExporter.Templates.RecordSet.RecordContext_T_.md 'CloudyWing.SpreadsheetExporter.Templates.RecordSet.RecordContext<T>')

#### Returns
[System.String](https://docs.microsoft.com/en-us/dotnet/api/System.String 'System.String')  
Returns the field formula.

<a name='CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumn_TRecord,TField_.GetFieldStyle(CloudyWing.SpreadsheetExporter.Templates.RecordSet.RecordContext_TRecord_)'></a>

## DataColumn<TRecord,TField>.GetFieldStyle(RecordContext<TRecord>) Method

Gets the field style.

```csharp
public override CloudyWing.SpreadsheetExporter.CellStyle GetFieldStyle(CloudyWing.SpreadsheetExporter.Templates.RecordSet.RecordContext<TRecord> recordContext);
```
#### Parameters

<a name='CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumn_TRecord,TField_.GetFieldStyle(CloudyWing.SpreadsheetExporter.Templates.RecordSet.RecordContext_TRecord_).recordContext'></a>

`recordContext` [CloudyWing.SpreadsheetExporter.Templates.RecordSet.RecordContext&lt;](CloudyWing.SpreadsheetExporter.Templates.RecordSet.RecordContext_T_.md 'CloudyWing.SpreadsheetExporter.Templates.RecordSet.RecordContext<T>')[TRecord](CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumn_TRecord,TField_.md#CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumn_TRecord,TField_.TRecord 'CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumn<TRecord,TField>.TRecord')[&gt;](CloudyWing.SpreadsheetExporter.Templates.RecordSet.RecordContext_T_.md 'CloudyWing.SpreadsheetExporter.Templates.RecordSet.RecordContext<T>')

#### Returns
[CellStyle](CloudyWing.SpreadsheetExporter.CellStyle.md 'CloudyWing.SpreadsheetExporter.CellStyle')  
Returns the field style.

<a name='CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumn_TRecord,TField_.GetFieldValue(CloudyWing.SpreadsheetExporter.Templates.RecordSet.RecordContext_TRecord_)'></a>

## DataColumn<TRecord,TField>.GetFieldValue(RecordContext<TRecord>) Method

Gets the field value.

```csharp
public override object GetFieldValue(CloudyWing.SpreadsheetExporter.Templates.RecordSet.RecordContext<TRecord> recordContext);
```
#### Parameters

<a name='CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumn_TRecord,TField_.GetFieldValue(CloudyWing.SpreadsheetExporter.Templates.RecordSet.RecordContext_TRecord_).recordContext'></a>

`recordContext` [CloudyWing.SpreadsheetExporter.Templates.RecordSet.RecordContext&lt;](CloudyWing.SpreadsheetExporter.Templates.RecordSet.RecordContext_T_.md 'CloudyWing.SpreadsheetExporter.Templates.RecordSet.RecordContext<T>')[TRecord](CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumn_TRecord,TField_.md#CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumn_TRecord,TField_.TRecord 'CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumn<TRecord,TField>.TRecord')[&gt;](CloudyWing.SpreadsheetExporter.Templates.RecordSet.RecordContext_T_.md 'CloudyWing.SpreadsheetExporter.Templates.RecordSet.RecordContext<T>')

#### Returns
[System.Object](https://docs.microsoft.com/en-us/dotnet/api/System.Object 'System.Object')  
Returns the field value.