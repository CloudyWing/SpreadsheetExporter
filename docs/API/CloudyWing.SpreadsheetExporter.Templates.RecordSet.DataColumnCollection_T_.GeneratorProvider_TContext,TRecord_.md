#### [CloudyWing.SpreadsheetExporter](index.md 'index')
### [CloudyWing.SpreadsheetExporter.Templates.RecordSet](CloudyWing.SpreadsheetExporter.Templates.RecordSet.md 'CloudyWing.SpreadsheetExporter.Templates.RecordSet').[DataColumnCollection&lt;T&gt;](CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.md 'CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection<T>')

## DataColumnCollection<T>.GeneratorProvider<TContext,TRecord> Class

The generator provider.

```csharp
public sealed class DataColumnCollection<T>.GeneratorProvider<TContext,TRecord>
    where TContext : CloudyWing.SpreadsheetExporter.Templates.RecordSet.RecordContext<TRecord>
```
#### Type parameters

<a name='CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.GeneratorProvider_TContext,TRecord_.T'></a>

`T`

<a name='CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.GeneratorProvider_TContext,TRecord_.TContext'></a>

`TContext`

The type of the context.

<a name='CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.GeneratorProvider_TContext,TRecord_.TRecord'></a>

`TRecord`

The type of the record.

Inheritance [System.Object](https://docs.microsoft.com/en-us/dotnet/api/System.Object 'System.Object') &#129106; GeneratorProvider<TContext,TRecord>

### See Also
- [Collection{DataColumnBase}](https://docs.microsoft.com/en-us/dotnet/api/System.Collections.ObjectModel.Collection-1 'System.Collections.ObjectModel.Collection`1')
### Methods

<a name='CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.GeneratorProvider_TContext,TRecord_.UseFormula(System.Func_TContext,string_)'></a>

## DataColumnCollection<T>.GeneratorProvider<TContext,TRecord>.UseFormula(Func<TContext,string>) Method

Uses the formula.

```csharp
public void UseFormula(System.Func<TContext,string> generator);
```
#### Parameters

<a name='CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.GeneratorProvider_TContext,TRecord_.UseFormula(System.Func_TContext,string_).generator'></a>

`generator` [System.Func&lt;](https://docs.microsoft.com/en-us/dotnet/api/System.Func-2 'System.Func`2')[TContext](CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.GeneratorProvider_TContext,TRecord_.md#CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.GeneratorProvider_TContext,TRecord_.TContext 'CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection<T>.GeneratorProvider<TContext,TRecord>.TContext')[,](https://docs.microsoft.com/en-us/dotnet/api/System.Func-2 'System.Func`2')[System.String](https://docs.microsoft.com/en-us/dotnet/api/System.String 'System.String')[&gt;](https://docs.microsoft.com/en-us/dotnet/api/System.Func-2 'System.Func`2')

The generator.

<a name='CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.GeneratorProvider_TContext,TRecord_.UseValue(System.Func_TContext,object_)'></a>

## DataColumnCollection<T>.GeneratorProvider<TContext,TRecord>.UseValue(Func<TContext,object>) Method

Uses the value.

```csharp
public void UseValue(System.Func<TContext,object> generator);
```
#### Parameters

<a name='CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.GeneratorProvider_TContext,TRecord_.UseValue(System.Func_TContext,object_).generator'></a>

`generator` [System.Func&lt;](https://docs.microsoft.com/en-us/dotnet/api/System.Func-2 'System.Func`2')[TContext](CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.GeneratorProvider_TContext,TRecord_.md#CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.GeneratorProvider_TContext,TRecord_.TContext 'CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection<T>.GeneratorProvider<TContext,TRecord>.TContext')[,](https://docs.microsoft.com/en-us/dotnet/api/System.Func-2 'System.Func`2')[System.Object](https://docs.microsoft.com/en-us/dotnet/api/System.Object 'System.Object')[&gt;](https://docs.microsoft.com/en-us/dotnet/api/System.Func-2 'System.Func`2')

The generator.