#### [CloudyWing.SpreadsheetExporter](index.md 'index')
### [CloudyWing.SpreadsheetExporter.Templates.RecordSet](CloudyWing.SpreadsheetExporter.Templates.RecordSet.md 'CloudyWing.SpreadsheetExporter.Templates.RecordSet')

## DataColumnCollection<T> Class

The data column collection.

```csharp
public class DataColumnCollection<T> : System.Collections.ObjectModel.Collection<CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnBase<T>>
```
#### Type parameters

<a name='CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.T'></a>

`T`

The type of the record.

Inheritance [System.Object](https://docs.microsoft.com/en-us/dotnet/api/System.Object 'System.Object') &#129106; [System.Collections.ObjectModel.Collection&lt;](https://docs.microsoft.com/en-us/dotnet/api/System.Collections.ObjectModel.Collection-1 'System.Collections.ObjectModel.Collection`1')[CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnBase&lt;](CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnBase_T_.md 'CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnBase<T>')[T](CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.md#CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.T 'CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection<T>.T')[&gt;](CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnBase_T_.md 'CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnBase<T>')[&gt;](https://docs.microsoft.com/en-us/dotnet/api/System.Collections.ObjectModel.Collection-1 'System.Collections.ObjectModel.Collection`1') &#129106; DataColumnCollection<T>
### Properties

<a name='CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.ColumnSpan'></a>

## DataColumnCollection<T>.ColumnSpan Property

Gets the column span.

```csharp
public int ColumnSpan { get; }
```

#### Property Value
[System.Int32](https://docs.microsoft.com/en-us/dotnet/api/System.Int32 'System.Int32')  
The column span.

<a name='CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.DataSourceColumns'></a>

## DataColumnCollection<T>.DataSourceColumns Property

Get the column containing the properties of the data source.

```csharp
public System.Collections.Generic.IEnumerable<CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnBase<T>> DataSourceColumns { get; }
```

#### Property Value
[System.Collections.Generic.IEnumerable&lt;](https://docs.microsoft.com/en-us/dotnet/api/System.Collections.Generic.IEnumerable-1 'System.Collections.Generic.IEnumerable`1')[CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnBase&lt;](CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnBase_T_.md 'CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnBase<T>')[T](CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.md#CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.T 'CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection<T>.T')[&gt;](CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnBase_T_.md 'CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnBase<T>')[&gt;](https://docs.microsoft.com/en-us/dotnet/api/System.Collections.Generic.IEnumerable-1 'System.Collections.Generic.IEnumerable`1')  
The data columns.

<a name='CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.RootColumns'></a>

## DataColumnCollection<T>.RootColumns Property

Gets the root columns.

```csharp
public CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection<T> RootColumns { get; }
```

#### Property Value
[CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection&lt;](CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.md 'CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection<T>')[T](CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.md#CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.T 'CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection<T>.T')[&gt;](CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.md 'CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection<T>')  
The root columns.

<a name='CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.RowSpan'></a>

## DataColumnCollection<T>.RowSpan Property

Gets the row span.

```csharp
public int RowSpan { get; }
```

#### Property Value
[System.Int32](https://docs.microsoft.com/en-us/dotnet/api/System.Int32 'System.Int32')  
The row span.
### Methods

<a name='CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.Add(string,System.Action_CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.GeneratorProvider_T,CloudyWing.SpreadsheetExporter.Templates.RecordSet.RecordContext_T___,System.Nullable_CloudyWing.SpreadsheetExporter.CellStyle_,System.Func_CloudyWing.SpreadsheetExporter.Templates.RecordSet.RecordContext_T_,CloudyWing.SpreadsheetExporter.CellStyle_)'></a>

## DataColumnCollection<T>.Add(string, Action<GeneratorProvider<T,RecordContext<T>>>, Nullable<CellStyle>, Func<RecordContext<T>,CellStyle>) Method

Adds the data column to the end of the DataColumnCollection<T>.

```csharp
public void Add(string headerText, System.Action<CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection<T>.GeneratorProvider<T,CloudyWing.SpreadsheetExporter.Templates.RecordSet.RecordContext<T>>> providerSetter, System.Nullable<CloudyWing.SpreadsheetExporter.CellStyle> headerStyle=null, System.Func<CloudyWing.SpreadsheetExporter.Templates.RecordSet.RecordContext<T>,CloudyWing.SpreadsheetExporter.CellStyle> fieldStyleGenerator=null);
```
#### Parameters

<a name='CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.Add(string,System.Action_CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.GeneratorProvider_T,CloudyWing.SpreadsheetExporter.Templates.RecordSet.RecordContext_T___,System.Nullable_CloudyWing.SpreadsheetExporter.CellStyle_,System.Func_CloudyWing.SpreadsheetExporter.Templates.RecordSet.RecordContext_T_,CloudyWing.SpreadsheetExporter.CellStyle_).headerText'></a>

`headerText` [System.String](https://docs.microsoft.com/en-us/dotnet/api/System.String 'System.String')

The header text.

<a name='CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.Add(string,System.Action_CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.GeneratorProvider_T,CloudyWing.SpreadsheetExporter.Templates.RecordSet.RecordContext_T___,System.Nullable_CloudyWing.SpreadsheetExporter.CellStyle_,System.Func_CloudyWing.SpreadsheetExporter.Templates.RecordSet.RecordContext_T_,CloudyWing.SpreadsheetExporter.CellStyle_).providerSetter'></a>

`providerSetter` [System.Action&lt;](https://docs.microsoft.com/en-us/dotnet/api/System.Action-1 'System.Action`1')[CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection.GeneratorProvider&lt;](CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.GeneratorProvider_TRecord,TContext_.md 'CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection<T>.GeneratorProvider<TRecord,TContext>')[T](CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.md#CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.T 'CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection<T>.T')[,](CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.GeneratorProvider_TRecord,TContext_.md 'CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection<T>.GeneratorProvider<TRecord,TContext>')[T](CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.md#CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.T 'CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection<T>.T')[,](CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.GeneratorProvider_TRecord,TContext_.md 'CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection<T>.GeneratorProvider<TRecord,TContext>')[CloudyWing.SpreadsheetExporter.Templates.RecordSet.RecordContext&lt;](CloudyWing.SpreadsheetExporter.Templates.RecordSet.RecordContext_T_.md 'CloudyWing.SpreadsheetExporter.Templates.RecordSet.RecordContext<T>')[T](CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.md#CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.T 'CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection<T>.T')[&gt;](CloudyWing.SpreadsheetExporter.Templates.RecordSet.RecordContext_T_.md 'CloudyWing.SpreadsheetExporter.Templates.RecordSet.RecordContext<T>')[&gt;](CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.GeneratorProvider_TRecord,TContext_.md 'CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection<T>.GeneratorProvider<TRecord,TContext>')[&gt;](https://docs.microsoft.com/en-us/dotnet/api/System.Action-1 'System.Action`1')

The provider setter.

<a name='CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.Add(string,System.Action_CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.GeneratorProvider_T,CloudyWing.SpreadsheetExporter.Templates.RecordSet.RecordContext_T___,System.Nullable_CloudyWing.SpreadsheetExporter.CellStyle_,System.Func_CloudyWing.SpreadsheetExporter.Templates.RecordSet.RecordContext_T_,CloudyWing.SpreadsheetExporter.CellStyle_).headerStyle'></a>

`headerStyle` [System.Nullable&lt;](https://docs.microsoft.com/en-us/dotnet/api/System.Nullable-1 'System.Nullable`1')[CellStyle](CloudyWing.SpreadsheetExporter.CellStyle.md 'CloudyWing.SpreadsheetExporter.CellStyle')[&gt;](https://docs.microsoft.com/en-us/dotnet/api/System.Nullable-1 'System.Nullable`1')

The header style. The dafault is `SpreadsheetManager.DefaultCellStyles.HeaderStyle`.

<a name='CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.Add(string,System.Action_CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.GeneratorProvider_T,CloudyWing.SpreadsheetExporter.Templates.RecordSet.RecordContext_T___,System.Nullable_CloudyWing.SpreadsheetExporter.CellStyle_,System.Func_CloudyWing.SpreadsheetExporter.Templates.RecordSet.RecordContext_T_,CloudyWing.SpreadsheetExporter.CellStyle_).fieldStyleGenerator'></a>

`fieldStyleGenerator` [System.Func&lt;](https://docs.microsoft.com/en-us/dotnet/api/System.Func-2 'System.Func`2')[CloudyWing.SpreadsheetExporter.Templates.RecordSet.RecordContext&lt;](CloudyWing.SpreadsheetExporter.Templates.RecordSet.RecordContext_T_.md 'CloudyWing.SpreadsheetExporter.Templates.RecordSet.RecordContext<T>')[T](CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.md#CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.T 'CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection<T>.T')[&gt;](CloudyWing.SpreadsheetExporter.Templates.RecordSet.RecordContext_T_.md 'CloudyWing.SpreadsheetExporter.Templates.RecordSet.RecordContext<T>')[,](https://docs.microsoft.com/en-us/dotnet/api/System.Func-2 'System.Func`2')[CellStyle](CloudyWing.SpreadsheetExporter.CellStyle.md 'CloudyWing.SpreadsheetExporter.CellStyle')[&gt;](https://docs.microsoft.com/en-us/dotnet/api/System.Func-2 'System.Func`2')

The field style generator. The dafault is `(context) => SpreadsheetManager.DefaultCellStyles.FieldStyle`.

<a name='CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.Add_TField_(string,string,System.Action_CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.GeneratorProvider_T,CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_T,TField___,System.Nullable_CloudyWing.SpreadsheetExporter.CellStyle_,System.Func_CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_T,TField_,CloudyWing.SpreadsheetExporter.CellStyle_)'></a>

## DataColumnCollection<T>.Add<TField>(string, string, Action<GeneratorProvider<T,FieldContext<T,TField>>>, Nullable<CellStyle>, Func<FieldContext<T,TField>,CellStyle>) Method

Adds the data column to the end of the DataColumnCollection<T>.

```csharp
public void Add<TField>(string headerText, string fieldKey, System.Action<CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection<T>.GeneratorProvider<T,CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext<T,TField>>> providerSetter, System.Nullable<CloudyWing.SpreadsheetExporter.CellStyle> headerStyle=null, System.Func<CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext<T,TField>,CloudyWing.SpreadsheetExporter.CellStyle> fieldStyleGenerator=null);
```
#### Type parameters

<a name='CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.Add_TField_(string,string,System.Action_CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.GeneratorProvider_T,CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_T,TField___,System.Nullable_CloudyWing.SpreadsheetExporter.CellStyle_,System.Func_CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_T,TField_,CloudyWing.SpreadsheetExporter.CellStyle_).TField'></a>

`TField`

The type of the field.
#### Parameters

<a name='CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.Add_TField_(string,string,System.Action_CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.GeneratorProvider_T,CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_T,TField___,System.Nullable_CloudyWing.SpreadsheetExporter.CellStyle_,System.Func_CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_T,TField_,CloudyWing.SpreadsheetExporter.CellStyle_).headerText'></a>

`headerText` [System.String](https://docs.microsoft.com/en-us/dotnet/api/System.String 'System.String')

The header text.

<a name='CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.Add_TField_(string,string,System.Action_CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.GeneratorProvider_T,CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_T,TField___,System.Nullable_CloudyWing.SpreadsheetExporter.CellStyle_,System.Func_CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_T,TField_,CloudyWing.SpreadsheetExporter.CellStyle_).fieldKey'></a>

`fieldKey` [System.String](https://docs.microsoft.com/en-us/dotnet/api/System.String 'System.String')

The field key.

<a name='CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.Add_TField_(string,string,System.Action_CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.GeneratorProvider_T,CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_T,TField___,System.Nullable_CloudyWing.SpreadsheetExporter.CellStyle_,System.Func_CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_T,TField_,CloudyWing.SpreadsheetExporter.CellStyle_).providerSetter'></a>

`providerSetter` [System.Action&lt;](https://docs.microsoft.com/en-us/dotnet/api/System.Action-1 'System.Action`1')[CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection.GeneratorProvider&lt;](CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.GeneratorProvider_TRecord,TContext_.md 'CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection<T>.GeneratorProvider<TRecord,TContext>')[T](CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.md#CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.T 'CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection<T>.T')[,](CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.GeneratorProvider_TRecord,TContext_.md 'CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection<T>.GeneratorProvider<TRecord,TContext>')[T](CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.md#CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.T 'CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection<T>.T')[,](CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.GeneratorProvider_TRecord,TContext_.md 'CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection<T>.GeneratorProvider<TRecord,TContext>')[CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext&lt;](CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_TRecoed,TField_.md 'CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext<TRecoed,TField>')[T](CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.md#CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.T 'CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection<T>.T')[,](CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_TRecoed,TField_.md 'CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext<TRecoed,TField>')[TField](CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.md#CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.Add_TField_(string,string,System.Action_CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.GeneratorProvider_T,CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_T,TField___,System.Nullable_CloudyWing.SpreadsheetExporter.CellStyle_,System.Func_CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_T,TField_,CloudyWing.SpreadsheetExporter.CellStyle_).TField 'CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection<T>.Add<TField>(string, string, System.Action<CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection<T>.GeneratorProvider<T,CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext<T,TField>>>, System.Nullable<CloudyWing.SpreadsheetExporter.CellStyle>, System.Func<CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext<T,TField>,CloudyWing.SpreadsheetExporter.CellStyle>).TField')[&gt;](CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_TRecoed,TField_.md 'CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext<TRecoed,TField>')[&gt;](CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.GeneratorProvider_TRecord,TContext_.md 'CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection<T>.GeneratorProvider<TRecord,TContext>')[&gt;](https://docs.microsoft.com/en-us/dotnet/api/System.Action-1 'System.Action`1')

The provider setter.

<a name='CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.Add_TField_(string,string,System.Action_CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.GeneratorProvider_T,CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_T,TField___,System.Nullable_CloudyWing.SpreadsheetExporter.CellStyle_,System.Func_CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_T,TField_,CloudyWing.SpreadsheetExporter.CellStyle_).headerStyle'></a>

`headerStyle` [System.Nullable&lt;](https://docs.microsoft.com/en-us/dotnet/api/System.Nullable-1 'System.Nullable`1')[CellStyle](CloudyWing.SpreadsheetExporter.CellStyle.md 'CloudyWing.SpreadsheetExporter.CellStyle')[&gt;](https://docs.microsoft.com/en-us/dotnet/api/System.Nullable-1 'System.Nullable`1')

The header style. The dafault is `SpreadsheetManager.DefaultCellStyles.HeaderStyle`.

<a name='CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.Add_TField_(string,string,System.Action_CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.GeneratorProvider_T,CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_T,TField___,System.Nullable_CloudyWing.SpreadsheetExporter.CellStyle_,System.Func_CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_T,TField_,CloudyWing.SpreadsheetExporter.CellStyle_).fieldStyleGenerator'></a>

`fieldStyleGenerator` [System.Func&lt;](https://docs.microsoft.com/en-us/dotnet/api/System.Func-2 'System.Func`2')[CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext&lt;](CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_TRecoed,TField_.md 'CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext<TRecoed,TField>')[T](CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.md#CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.T 'CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection<T>.T')[,](CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_TRecoed,TField_.md 'CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext<TRecoed,TField>')[TField](CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.md#CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.Add_TField_(string,string,System.Action_CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.GeneratorProvider_T,CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_T,TField___,System.Nullable_CloudyWing.SpreadsheetExporter.CellStyle_,System.Func_CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_T,TField_,CloudyWing.SpreadsheetExporter.CellStyle_).TField 'CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection<T>.Add<TField>(string, string, System.Action<CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection<T>.GeneratorProvider<T,CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext<T,TField>>>, System.Nullable<CloudyWing.SpreadsheetExporter.CellStyle>, System.Func<CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext<T,TField>,CloudyWing.SpreadsheetExporter.CellStyle>).TField')[&gt;](CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_TRecoed,TField_.md 'CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext<TRecoed,TField>')[,](https://docs.microsoft.com/en-us/dotnet/api/System.Func-2 'System.Func`2')[CellStyle](CloudyWing.SpreadsheetExporter.CellStyle.md 'CloudyWing.SpreadsheetExporter.CellStyle')[&gt;](https://docs.microsoft.com/en-us/dotnet/api/System.Func-2 'System.Func`2')

The field style generator. The dafault is `(context) => SpreadsheetManager.DefaultCellStyles.FieldStyle`.

<a name='CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.Add_TField_(string,string,System.Nullable_CloudyWing.SpreadsheetExporter.CellStyle_,System.Func_CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_T,TField_,CloudyWing.SpreadsheetExporter.CellStyle_)'></a>

## DataColumnCollection<T>.Add<TField>(string, string, Nullable<CellStyle>, Func<FieldContext<T,TField>,CellStyle>) Method

Adds the data column to the end of the DataColumnCollection<T>.

```csharp
public void Add<TField>(string headerText, string fieldKey, System.Nullable<CloudyWing.SpreadsheetExporter.CellStyle> headerStyle=null, System.Func<CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext<T,TField>,CloudyWing.SpreadsheetExporter.CellStyle> fieldStyleGenerator=null);
```
#### Type parameters

<a name='CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.Add_TField_(string,string,System.Nullable_CloudyWing.SpreadsheetExporter.CellStyle_,System.Func_CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_T,TField_,CloudyWing.SpreadsheetExporter.CellStyle_).TField'></a>

`TField`

The type of the field.
#### Parameters

<a name='CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.Add_TField_(string,string,System.Nullable_CloudyWing.SpreadsheetExporter.CellStyle_,System.Func_CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_T,TField_,CloudyWing.SpreadsheetExporter.CellStyle_).headerText'></a>

`headerText` [System.String](https://docs.microsoft.com/en-us/dotnet/api/System.String 'System.String')

The header text.

<a name='CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.Add_TField_(string,string,System.Nullable_CloudyWing.SpreadsheetExporter.CellStyle_,System.Func_CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_T,TField_,CloudyWing.SpreadsheetExporter.CellStyle_).fieldKey'></a>

`fieldKey` [System.String](https://docs.microsoft.com/en-us/dotnet/api/System.String 'System.String')

The field key.

<a name='CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.Add_TField_(string,string,System.Nullable_CloudyWing.SpreadsheetExporter.CellStyle_,System.Func_CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_T,TField_,CloudyWing.SpreadsheetExporter.CellStyle_).headerStyle'></a>

`headerStyle` [System.Nullable&lt;](https://docs.microsoft.com/en-us/dotnet/api/System.Nullable-1 'System.Nullable`1')[CellStyle](CloudyWing.SpreadsheetExporter.CellStyle.md 'CloudyWing.SpreadsheetExporter.CellStyle')[&gt;](https://docs.microsoft.com/en-us/dotnet/api/System.Nullable-1 'System.Nullable`1')

The header style. The dafault is `SpreadsheetManager.DefaultCellStyles.HeaderStyle`.

<a name='CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.Add_TField_(string,string,System.Nullable_CloudyWing.SpreadsheetExporter.CellStyle_,System.Func_CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_T,TField_,CloudyWing.SpreadsheetExporter.CellStyle_).fieldStyleGenerator'></a>

`fieldStyleGenerator` [System.Func&lt;](https://docs.microsoft.com/en-us/dotnet/api/System.Func-2 'System.Func`2')[CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext&lt;](CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_TRecoed,TField_.md 'CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext<TRecoed,TField>')[T](CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.md#CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.T 'CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection<T>.T')[,](CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_TRecoed,TField_.md 'CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext<TRecoed,TField>')[TField](CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.md#CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.Add_TField_(string,string,System.Nullable_CloudyWing.SpreadsheetExporter.CellStyle_,System.Func_CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_T,TField_,CloudyWing.SpreadsheetExporter.CellStyle_).TField 'CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection<T>.Add<TField>(string, string, System.Nullable<CloudyWing.SpreadsheetExporter.CellStyle>, System.Func<CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext<T,TField>,CloudyWing.SpreadsheetExporter.CellStyle>).TField')[&gt;](CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_TRecoed,TField_.md 'CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext<TRecoed,TField>')[,](https://docs.microsoft.com/en-us/dotnet/api/System.Func-2 'System.Func`2')[CellStyle](CloudyWing.SpreadsheetExporter.CellStyle.md 'CloudyWing.SpreadsheetExporter.CellStyle')[&gt;](https://docs.microsoft.com/en-us/dotnet/api/System.Func-2 'System.Func`2')

The field style generator. The dafault is `(context) => SpreadsheetManager.DefaultCellStyles.FieldStyle`.

<a name='CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.Add_TField_(string,System.Linq.Expressions.Expression_System.Func_T,TField__,System.Action_CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.GeneratorProvider_T,CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_T,TField___,System.Nullable_CloudyWing.SpreadsheetExporter.CellStyle_,System.Func_CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_T,TField_,CloudyWing.SpreadsheetExporter.CellStyle_)'></a>

## DataColumnCollection<T>.Add<TField>(string, Expression<Func<T,TField>>, Action<GeneratorProvider<T,FieldContext<T,TField>>>, Nullable<CellStyle>, Func<FieldContext<T,TField>,CellStyle>) Method

Adds the data column to the end of the DataColumnCollection<T>.

```csharp
public void Add<TField>(string headerText, System.Linq.Expressions.Expression<System.Func<T,TField>> fieldKeyExpression, System.Action<CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection<T>.GeneratorProvider<T,CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext<T,TField>>> providerSetter, System.Nullable<CloudyWing.SpreadsheetExporter.CellStyle> headerStyle=null, System.Func<CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext<T,TField>,CloudyWing.SpreadsheetExporter.CellStyle> fieldStyleGenerator=null);
```
#### Type parameters

<a name='CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.Add_TField_(string,System.Linq.Expressions.Expression_System.Func_T,TField__,System.Action_CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.GeneratorProvider_T,CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_T,TField___,System.Nullable_CloudyWing.SpreadsheetExporter.CellStyle_,System.Func_CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_T,TField_,CloudyWing.SpreadsheetExporter.CellStyle_).TField'></a>

`TField`

The type of the field.
#### Parameters

<a name='CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.Add_TField_(string,System.Linq.Expressions.Expression_System.Func_T,TField__,System.Action_CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.GeneratorProvider_T,CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_T,TField___,System.Nullable_CloudyWing.SpreadsheetExporter.CellStyle_,System.Func_CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_T,TField_,CloudyWing.SpreadsheetExporter.CellStyle_).headerText'></a>

`headerText` [System.String](https://docs.microsoft.com/en-us/dotnet/api/System.String 'System.String')

The header text.

<a name='CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.Add_TField_(string,System.Linq.Expressions.Expression_System.Func_T,TField__,System.Action_CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.GeneratorProvider_T,CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_T,TField___,System.Nullable_CloudyWing.SpreadsheetExporter.CellStyle_,System.Func_CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_T,TField_,CloudyWing.SpreadsheetExporter.CellStyle_).fieldKeyExpression'></a>

`fieldKeyExpression` [System.Linq.Expressions.Expression&lt;](https://docs.microsoft.com/en-us/dotnet/api/System.Linq.Expressions.Expression-1 'System.Linq.Expressions.Expression`1')[System.Func&lt;](https://docs.microsoft.com/en-us/dotnet/api/System.Func-2 'System.Func`2')[T](CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.md#CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.T 'CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection<T>.T')[,](https://docs.microsoft.com/en-us/dotnet/api/System.Func-2 'System.Func`2')[TField](CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.md#CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.Add_TField_(string,System.Linq.Expressions.Expression_System.Func_T,TField__,System.Action_CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.GeneratorProvider_T,CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_T,TField___,System.Nullable_CloudyWing.SpreadsheetExporter.CellStyle_,System.Func_CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_T,TField_,CloudyWing.SpreadsheetExporter.CellStyle_).TField 'CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection<T>.Add<TField>(string, System.Linq.Expressions.Expression<System.Func<T,TField>>, System.Action<CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection<T>.GeneratorProvider<T,CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext<T,TField>>>, System.Nullable<CloudyWing.SpreadsheetExporter.CellStyle>, System.Func<CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext<T,TField>,CloudyWing.SpreadsheetExporter.CellStyle>).TField')[&gt;](https://docs.microsoft.com/en-us/dotnet/api/System.Func-2 'System.Func`2')[&gt;](https://docs.microsoft.com/en-us/dotnet/api/System.Linq.Expressions.Expression-1 'System.Linq.Expressions.Expression`1')

The field key expression.

<a name='CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.Add_TField_(string,System.Linq.Expressions.Expression_System.Func_T,TField__,System.Action_CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.GeneratorProvider_T,CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_T,TField___,System.Nullable_CloudyWing.SpreadsheetExporter.CellStyle_,System.Func_CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_T,TField_,CloudyWing.SpreadsheetExporter.CellStyle_).providerSetter'></a>

`providerSetter` [System.Action&lt;](https://docs.microsoft.com/en-us/dotnet/api/System.Action-1 'System.Action`1')[CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection.GeneratorProvider&lt;](CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.GeneratorProvider_TRecord,TContext_.md 'CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection<T>.GeneratorProvider<TRecord,TContext>')[T](CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.md#CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.T 'CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection<T>.T')[,](CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.GeneratorProvider_TRecord,TContext_.md 'CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection<T>.GeneratorProvider<TRecord,TContext>')[T](CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.md#CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.T 'CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection<T>.T')[,](CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.GeneratorProvider_TRecord,TContext_.md 'CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection<T>.GeneratorProvider<TRecord,TContext>')[CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext&lt;](CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_TRecoed,TField_.md 'CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext<TRecoed,TField>')[T](CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.md#CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.T 'CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection<T>.T')[,](CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_TRecoed,TField_.md 'CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext<TRecoed,TField>')[TField](CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.md#CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.Add_TField_(string,System.Linq.Expressions.Expression_System.Func_T,TField__,System.Action_CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.GeneratorProvider_T,CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_T,TField___,System.Nullable_CloudyWing.SpreadsheetExporter.CellStyle_,System.Func_CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_T,TField_,CloudyWing.SpreadsheetExporter.CellStyle_).TField 'CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection<T>.Add<TField>(string, System.Linq.Expressions.Expression<System.Func<T,TField>>, System.Action<CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection<T>.GeneratorProvider<T,CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext<T,TField>>>, System.Nullable<CloudyWing.SpreadsheetExporter.CellStyle>, System.Func<CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext<T,TField>,CloudyWing.SpreadsheetExporter.CellStyle>).TField')[&gt;](CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_TRecoed,TField_.md 'CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext<TRecoed,TField>')[&gt;](CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.GeneratorProvider_TRecord,TContext_.md 'CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection<T>.GeneratorProvider<TRecord,TContext>')[&gt;](https://docs.microsoft.com/en-us/dotnet/api/System.Action-1 'System.Action`1')

The provider setter.

<a name='CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.Add_TField_(string,System.Linq.Expressions.Expression_System.Func_T,TField__,System.Action_CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.GeneratorProvider_T,CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_T,TField___,System.Nullable_CloudyWing.SpreadsheetExporter.CellStyle_,System.Func_CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_T,TField_,CloudyWing.SpreadsheetExporter.CellStyle_).headerStyle'></a>

`headerStyle` [System.Nullable&lt;](https://docs.microsoft.com/en-us/dotnet/api/System.Nullable-1 'System.Nullable`1')[CellStyle](CloudyWing.SpreadsheetExporter.CellStyle.md 'CloudyWing.SpreadsheetExporter.CellStyle')[&gt;](https://docs.microsoft.com/en-us/dotnet/api/System.Nullable-1 'System.Nullable`1')

The header style. The dafault is `SpreadsheetManager.DefaultCellStyles.HeaderStyle`.

<a name='CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.Add_TField_(string,System.Linq.Expressions.Expression_System.Func_T,TField__,System.Action_CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.GeneratorProvider_T,CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_T,TField___,System.Nullable_CloudyWing.SpreadsheetExporter.CellStyle_,System.Func_CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_T,TField_,CloudyWing.SpreadsheetExporter.CellStyle_).fieldStyleGenerator'></a>

`fieldStyleGenerator` [System.Func&lt;](https://docs.microsoft.com/en-us/dotnet/api/System.Func-2 'System.Func`2')[CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext&lt;](CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_TRecoed,TField_.md 'CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext<TRecoed,TField>')[T](CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.md#CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.T 'CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection<T>.T')[,](CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_TRecoed,TField_.md 'CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext<TRecoed,TField>')[TField](CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.md#CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.Add_TField_(string,System.Linq.Expressions.Expression_System.Func_T,TField__,System.Action_CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.GeneratorProvider_T,CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_T,TField___,System.Nullable_CloudyWing.SpreadsheetExporter.CellStyle_,System.Func_CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_T,TField_,CloudyWing.SpreadsheetExporter.CellStyle_).TField 'CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection<T>.Add<TField>(string, System.Linq.Expressions.Expression<System.Func<T,TField>>, System.Action<CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection<T>.GeneratorProvider<T,CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext<T,TField>>>, System.Nullable<CloudyWing.SpreadsheetExporter.CellStyle>, System.Func<CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext<T,TField>,CloudyWing.SpreadsheetExporter.CellStyle>).TField')[&gt;](CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_TRecoed,TField_.md 'CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext<TRecoed,TField>')[,](https://docs.microsoft.com/en-us/dotnet/api/System.Func-2 'System.Func`2')[CellStyle](CloudyWing.SpreadsheetExporter.CellStyle.md 'CloudyWing.SpreadsheetExporter.CellStyle')[&gt;](https://docs.microsoft.com/en-us/dotnet/api/System.Func-2 'System.Func`2')

The field style generator. The dafault is `(context) => SpreadsheetManager.DefaultCellStyles.FieldStyle`.

<a name='CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.Add_TField_(string,System.Linq.Expressions.Expression_System.Func_T,TField__,System.Nullable_CloudyWing.SpreadsheetExporter.CellStyle_,System.Func_CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_T,TField_,CloudyWing.SpreadsheetExporter.CellStyle_)'></a>

## DataColumnCollection<T>.Add<TField>(string, Expression<Func<T,TField>>, Nullable<CellStyle>, Func<FieldContext<T,TField>,CellStyle>) Method

Adds the data column to the end of the DataColumnCollection<T>.

```csharp
public void Add<TField>(string headerText, System.Linq.Expressions.Expression<System.Func<T,TField>> fieldKeyExpression, System.Nullable<CloudyWing.SpreadsheetExporter.CellStyle> headerStyle=null, System.Func<CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext<T,TField>,CloudyWing.SpreadsheetExporter.CellStyle> fieldStyleGenerator=null);
```
#### Type parameters

<a name='CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.Add_TField_(string,System.Linq.Expressions.Expression_System.Func_T,TField__,System.Nullable_CloudyWing.SpreadsheetExporter.CellStyle_,System.Func_CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_T,TField_,CloudyWing.SpreadsheetExporter.CellStyle_).TField'></a>

`TField`

The type of the field.
#### Parameters

<a name='CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.Add_TField_(string,System.Linq.Expressions.Expression_System.Func_T,TField__,System.Nullable_CloudyWing.SpreadsheetExporter.CellStyle_,System.Func_CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_T,TField_,CloudyWing.SpreadsheetExporter.CellStyle_).headerText'></a>

`headerText` [System.String](https://docs.microsoft.com/en-us/dotnet/api/System.String 'System.String')

The header text.

<a name='CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.Add_TField_(string,System.Linq.Expressions.Expression_System.Func_T,TField__,System.Nullable_CloudyWing.SpreadsheetExporter.CellStyle_,System.Func_CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_T,TField_,CloudyWing.SpreadsheetExporter.CellStyle_).fieldKeyExpression'></a>

`fieldKeyExpression` [System.Linq.Expressions.Expression&lt;](https://docs.microsoft.com/en-us/dotnet/api/System.Linq.Expressions.Expression-1 'System.Linq.Expressions.Expression`1')[System.Func&lt;](https://docs.microsoft.com/en-us/dotnet/api/System.Func-2 'System.Func`2')[T](CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.md#CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.T 'CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection<T>.T')[,](https://docs.microsoft.com/en-us/dotnet/api/System.Func-2 'System.Func`2')[TField](CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.md#CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.Add_TField_(string,System.Linq.Expressions.Expression_System.Func_T,TField__,System.Nullable_CloudyWing.SpreadsheetExporter.CellStyle_,System.Func_CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_T,TField_,CloudyWing.SpreadsheetExporter.CellStyle_).TField 'CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection<T>.Add<TField>(string, System.Linq.Expressions.Expression<System.Func<T,TField>>, System.Nullable<CloudyWing.SpreadsheetExporter.CellStyle>, System.Func<CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext<T,TField>,CloudyWing.SpreadsheetExporter.CellStyle>).TField')[&gt;](https://docs.microsoft.com/en-us/dotnet/api/System.Func-2 'System.Func`2')[&gt;](https://docs.microsoft.com/en-us/dotnet/api/System.Linq.Expressions.Expression-1 'System.Linq.Expressions.Expression`1')

The field key expression.

<a name='CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.Add_TField_(string,System.Linq.Expressions.Expression_System.Func_T,TField__,System.Nullable_CloudyWing.SpreadsheetExporter.CellStyle_,System.Func_CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_T,TField_,CloudyWing.SpreadsheetExporter.CellStyle_).headerStyle'></a>

`headerStyle` [System.Nullable&lt;](https://docs.microsoft.com/en-us/dotnet/api/System.Nullable-1 'System.Nullable`1')[CellStyle](CloudyWing.SpreadsheetExporter.CellStyle.md 'CloudyWing.SpreadsheetExporter.CellStyle')[&gt;](https://docs.microsoft.com/en-us/dotnet/api/System.Nullable-1 'System.Nullable`1')

The header style. The dafault is `SpreadsheetManager.DefaultCellStyles.HeaderStyle`.

<a name='CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.Add_TField_(string,System.Linq.Expressions.Expression_System.Func_T,TField__,System.Nullable_CloudyWing.SpreadsheetExporter.CellStyle_,System.Func_CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_T,TField_,CloudyWing.SpreadsheetExporter.CellStyle_).fieldStyleGenerator'></a>

`fieldStyleGenerator` [System.Func&lt;](https://docs.microsoft.com/en-us/dotnet/api/System.Func-2 'System.Func`2')[CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext&lt;](CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_TRecoed,TField_.md 'CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext<TRecoed,TField>')[T](CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.md#CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.T 'CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection<T>.T')[,](CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_TRecoed,TField_.md 'CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext<TRecoed,TField>')[TField](CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.md#CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.Add_TField_(string,System.Linq.Expressions.Expression_System.Func_T,TField__,System.Nullable_CloudyWing.SpreadsheetExporter.CellStyle_,System.Func_CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_T,TField_,CloudyWing.SpreadsheetExporter.CellStyle_).TField 'CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection<T>.Add<TField>(string, System.Linq.Expressions.Expression<System.Func<T,TField>>, System.Nullable<CloudyWing.SpreadsheetExporter.CellStyle>, System.Func<CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext<T,TField>,CloudyWing.SpreadsheetExporter.CellStyle>).TField')[&gt;](CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_TRecoed,TField_.md 'CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext<TRecoed,TField>')[,](https://docs.microsoft.com/en-us/dotnet/api/System.Func-2 'System.Func`2')[CellStyle](CloudyWing.SpreadsheetExporter.CellStyle.md 'CloudyWing.SpreadsheetExporter.CellStyle')[&gt;](https://docs.microsoft.com/en-us/dotnet/api/System.Func-2 'System.Func`2')

The field style generator. The dafault is `(context) => SpreadsheetManager.DefaultCellStyles.FieldStyle`.

<a name='CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.AddChildToLast(CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnBase_T_)'></a>

## DataColumnCollection<T>.AddChildToLast(DataColumnBase<T>) Method

Adds the child data column at the end of the last data column.

```csharp
public void AddChildToLast(CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnBase<T> childColumn);
```
#### Parameters

<a name='CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.AddChildToLast(CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnBase_T_).childColumn'></a>

`childColumn` [CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnBase&lt;](CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnBase_T_.md 'CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnBase<T>')[T](CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.md#CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.T 'CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection<T>.T')[&gt;](CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnBase_T_.md 'CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnBase<T>')

The child data column.

#### Exceptions

[System.NullReferenceException](https://docs.microsoft.com/en-us/dotnet/api/System.NullReferenceException 'System.NullReferenceException')

<a name='CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.AddChildToLast(string,System.Action_CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.GeneratorProvider_T,CloudyWing.SpreadsheetExporter.Templates.RecordSet.RecordContext_T___,System.Nullable_CloudyWing.SpreadsheetExporter.CellStyle_,System.Func_CloudyWing.SpreadsheetExporter.Templates.RecordSet.RecordContext_T_,CloudyWing.SpreadsheetExporter.CellStyle_)'></a>

## DataColumnCollection<T>.AddChildToLast(string, Action<GeneratorProvider<T,RecordContext<T>>>, Nullable<CellStyle>, Func<RecordContext<T>,CellStyle>) Method

Adds the child data column at the end of the last data column.

```csharp
public void AddChildToLast(string headerText, System.Action<CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection<T>.GeneratorProvider<T,CloudyWing.SpreadsheetExporter.Templates.RecordSet.RecordContext<T>>> providerSetter, System.Nullable<CloudyWing.SpreadsheetExporter.CellStyle> headerStyle=null, System.Func<CloudyWing.SpreadsheetExporter.Templates.RecordSet.RecordContext<T>,CloudyWing.SpreadsheetExporter.CellStyle> fieldStyleGenerator=null);
```
#### Parameters

<a name='CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.AddChildToLast(string,System.Action_CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.GeneratorProvider_T,CloudyWing.SpreadsheetExporter.Templates.RecordSet.RecordContext_T___,System.Nullable_CloudyWing.SpreadsheetExporter.CellStyle_,System.Func_CloudyWing.SpreadsheetExporter.Templates.RecordSet.RecordContext_T_,CloudyWing.SpreadsheetExporter.CellStyle_).headerText'></a>

`headerText` [System.String](https://docs.microsoft.com/en-us/dotnet/api/System.String 'System.String')

The header text.

<a name='CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.AddChildToLast(string,System.Action_CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.GeneratorProvider_T,CloudyWing.SpreadsheetExporter.Templates.RecordSet.RecordContext_T___,System.Nullable_CloudyWing.SpreadsheetExporter.CellStyle_,System.Func_CloudyWing.SpreadsheetExporter.Templates.RecordSet.RecordContext_T_,CloudyWing.SpreadsheetExporter.CellStyle_).providerSetter'></a>

`providerSetter` [System.Action&lt;](https://docs.microsoft.com/en-us/dotnet/api/System.Action-1 'System.Action`1')[CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection.GeneratorProvider&lt;](CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.GeneratorProvider_TRecord,TContext_.md 'CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection<T>.GeneratorProvider<TRecord,TContext>')[T](CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.md#CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.T 'CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection<T>.T')[,](CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.GeneratorProvider_TRecord,TContext_.md 'CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection<T>.GeneratorProvider<TRecord,TContext>')[T](CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.md#CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.T 'CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection<T>.T')[,](CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.GeneratorProvider_TRecord,TContext_.md 'CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection<T>.GeneratorProvider<TRecord,TContext>')[CloudyWing.SpreadsheetExporter.Templates.RecordSet.RecordContext&lt;](CloudyWing.SpreadsheetExporter.Templates.RecordSet.RecordContext_T_.md 'CloudyWing.SpreadsheetExporter.Templates.RecordSet.RecordContext<T>')[T](CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.md#CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.T 'CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection<T>.T')[&gt;](CloudyWing.SpreadsheetExporter.Templates.RecordSet.RecordContext_T_.md 'CloudyWing.SpreadsheetExporter.Templates.RecordSet.RecordContext<T>')[&gt;](CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.GeneratorProvider_TRecord,TContext_.md 'CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection<T>.GeneratorProvider<TRecord,TContext>')[&gt;](https://docs.microsoft.com/en-us/dotnet/api/System.Action-1 'System.Action`1')

The provider setter.

<a name='CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.AddChildToLast(string,System.Action_CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.GeneratorProvider_T,CloudyWing.SpreadsheetExporter.Templates.RecordSet.RecordContext_T___,System.Nullable_CloudyWing.SpreadsheetExporter.CellStyle_,System.Func_CloudyWing.SpreadsheetExporter.Templates.RecordSet.RecordContext_T_,CloudyWing.SpreadsheetExporter.CellStyle_).headerStyle'></a>

`headerStyle` [System.Nullable&lt;](https://docs.microsoft.com/en-us/dotnet/api/System.Nullable-1 'System.Nullable`1')[CellStyle](CloudyWing.SpreadsheetExporter.CellStyle.md 'CloudyWing.SpreadsheetExporter.CellStyle')[&gt;](https://docs.microsoft.com/en-us/dotnet/api/System.Nullable-1 'System.Nullable`1')

The header style. The dafault is `SpreadsheetManager.DefaultCellStyles.HeaderStyle`.

<a name='CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.AddChildToLast(string,System.Action_CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.GeneratorProvider_T,CloudyWing.SpreadsheetExporter.Templates.RecordSet.RecordContext_T___,System.Nullable_CloudyWing.SpreadsheetExporter.CellStyle_,System.Func_CloudyWing.SpreadsheetExporter.Templates.RecordSet.RecordContext_T_,CloudyWing.SpreadsheetExporter.CellStyle_).fieldStyleGenerator'></a>

`fieldStyleGenerator` [System.Func&lt;](https://docs.microsoft.com/en-us/dotnet/api/System.Func-2 'System.Func`2')[CloudyWing.SpreadsheetExporter.Templates.RecordSet.RecordContext&lt;](CloudyWing.SpreadsheetExporter.Templates.RecordSet.RecordContext_T_.md 'CloudyWing.SpreadsheetExporter.Templates.RecordSet.RecordContext<T>')[T](CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.md#CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.T 'CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection<T>.T')[&gt;](CloudyWing.SpreadsheetExporter.Templates.RecordSet.RecordContext_T_.md 'CloudyWing.SpreadsheetExporter.Templates.RecordSet.RecordContext<T>')[,](https://docs.microsoft.com/en-us/dotnet/api/System.Func-2 'System.Func`2')[CellStyle](CloudyWing.SpreadsheetExporter.CellStyle.md 'CloudyWing.SpreadsheetExporter.CellStyle')[&gt;](https://docs.microsoft.com/en-us/dotnet/api/System.Func-2 'System.Func`2')

The field style generator. The dafault is `(context) => SpreadsheetManager.DefaultCellStyles.FieldStyle`.

<a name='CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.AddChildToLast_TField_(string,string,System.Action_CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.GeneratorProvider_T,CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_T,TField___,System.Nullable_CloudyWing.SpreadsheetExporter.CellStyle_,System.Func_CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_T,TField_,CloudyWing.SpreadsheetExporter.CellStyle_)'></a>

## DataColumnCollection<T>.AddChildToLast<TField>(string, string, Action<GeneratorProvider<T,FieldContext<T,TField>>>, Nullable<CellStyle>, Func<FieldContext<T,TField>,CellStyle>) Method

Adds the child data column at the end of the last data column.

```csharp
public void AddChildToLast<TField>(string headerText, string fieldKey, System.Action<CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection<T>.GeneratorProvider<T,CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext<T,TField>>> providerSetter, System.Nullable<CloudyWing.SpreadsheetExporter.CellStyle> headerStyle=null, System.Func<CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext<T,TField>,CloudyWing.SpreadsheetExporter.CellStyle> fieldStyleGenerator=null);
```
#### Type parameters

<a name='CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.AddChildToLast_TField_(string,string,System.Action_CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.GeneratorProvider_T,CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_T,TField___,System.Nullable_CloudyWing.SpreadsheetExporter.CellStyle_,System.Func_CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_T,TField_,CloudyWing.SpreadsheetExporter.CellStyle_).TField'></a>

`TField`

The type of the field.
#### Parameters

<a name='CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.AddChildToLast_TField_(string,string,System.Action_CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.GeneratorProvider_T,CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_T,TField___,System.Nullable_CloudyWing.SpreadsheetExporter.CellStyle_,System.Func_CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_T,TField_,CloudyWing.SpreadsheetExporter.CellStyle_).headerText'></a>

`headerText` [System.String](https://docs.microsoft.com/en-us/dotnet/api/System.String 'System.String')

The header text.

<a name='CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.AddChildToLast_TField_(string,string,System.Action_CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.GeneratorProvider_T,CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_T,TField___,System.Nullable_CloudyWing.SpreadsheetExporter.CellStyle_,System.Func_CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_T,TField_,CloudyWing.SpreadsheetExporter.CellStyle_).fieldKey'></a>

`fieldKey` [System.String](https://docs.microsoft.com/en-us/dotnet/api/System.String 'System.String')

The field key.

<a name='CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.AddChildToLast_TField_(string,string,System.Action_CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.GeneratorProvider_T,CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_T,TField___,System.Nullable_CloudyWing.SpreadsheetExporter.CellStyle_,System.Func_CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_T,TField_,CloudyWing.SpreadsheetExporter.CellStyle_).providerSetter'></a>

`providerSetter` [System.Action&lt;](https://docs.microsoft.com/en-us/dotnet/api/System.Action-1 'System.Action`1')[CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection.GeneratorProvider&lt;](CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.GeneratorProvider_TRecord,TContext_.md 'CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection<T>.GeneratorProvider<TRecord,TContext>')[T](CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.md#CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.T 'CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection<T>.T')[,](CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.GeneratorProvider_TRecord,TContext_.md 'CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection<T>.GeneratorProvider<TRecord,TContext>')[T](CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.md#CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.T 'CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection<T>.T')[,](CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.GeneratorProvider_TRecord,TContext_.md 'CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection<T>.GeneratorProvider<TRecord,TContext>')[CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext&lt;](CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_TRecoed,TField_.md 'CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext<TRecoed,TField>')[T](CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.md#CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.T 'CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection<T>.T')[,](CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_TRecoed,TField_.md 'CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext<TRecoed,TField>')[TField](CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.md#CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.AddChildToLast_TField_(string,string,System.Action_CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.GeneratorProvider_T,CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_T,TField___,System.Nullable_CloudyWing.SpreadsheetExporter.CellStyle_,System.Func_CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_T,TField_,CloudyWing.SpreadsheetExporter.CellStyle_).TField 'CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection<T>.AddChildToLast<TField>(string, string, System.Action<CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection<T>.GeneratorProvider<T,CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext<T,TField>>>, System.Nullable<CloudyWing.SpreadsheetExporter.CellStyle>, System.Func<CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext<T,TField>,CloudyWing.SpreadsheetExporter.CellStyle>).TField')[&gt;](CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_TRecoed,TField_.md 'CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext<TRecoed,TField>')[&gt;](CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.GeneratorProvider_TRecord,TContext_.md 'CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection<T>.GeneratorProvider<TRecord,TContext>')[&gt;](https://docs.microsoft.com/en-us/dotnet/api/System.Action-1 'System.Action`1')

The provider setter.

<a name='CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.AddChildToLast_TField_(string,string,System.Action_CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.GeneratorProvider_T,CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_T,TField___,System.Nullable_CloudyWing.SpreadsheetExporter.CellStyle_,System.Func_CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_T,TField_,CloudyWing.SpreadsheetExporter.CellStyle_).headerStyle'></a>

`headerStyle` [System.Nullable&lt;](https://docs.microsoft.com/en-us/dotnet/api/System.Nullable-1 'System.Nullable`1')[CellStyle](CloudyWing.SpreadsheetExporter.CellStyle.md 'CloudyWing.SpreadsheetExporter.CellStyle')[&gt;](https://docs.microsoft.com/en-us/dotnet/api/System.Nullable-1 'System.Nullable`1')

The header style. The dafault is `SpreadsheetManager.DefaultCellStyles.HeaderStyle`.

<a name='CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.AddChildToLast_TField_(string,string,System.Action_CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.GeneratorProvider_T,CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_T,TField___,System.Nullable_CloudyWing.SpreadsheetExporter.CellStyle_,System.Func_CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_T,TField_,CloudyWing.SpreadsheetExporter.CellStyle_).fieldStyleGenerator'></a>

`fieldStyleGenerator` [System.Func&lt;](https://docs.microsoft.com/en-us/dotnet/api/System.Func-2 'System.Func`2')[CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext&lt;](CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_TRecoed,TField_.md 'CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext<TRecoed,TField>')[T](CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.md#CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.T 'CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection<T>.T')[,](CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_TRecoed,TField_.md 'CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext<TRecoed,TField>')[TField](CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.md#CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.AddChildToLast_TField_(string,string,System.Action_CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.GeneratorProvider_T,CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_T,TField___,System.Nullable_CloudyWing.SpreadsheetExporter.CellStyle_,System.Func_CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_T,TField_,CloudyWing.SpreadsheetExporter.CellStyle_).TField 'CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection<T>.AddChildToLast<TField>(string, string, System.Action<CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection<T>.GeneratorProvider<T,CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext<T,TField>>>, System.Nullable<CloudyWing.SpreadsheetExporter.CellStyle>, System.Func<CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext<T,TField>,CloudyWing.SpreadsheetExporter.CellStyle>).TField')[&gt;](CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_TRecoed,TField_.md 'CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext<TRecoed,TField>')[,](https://docs.microsoft.com/en-us/dotnet/api/System.Func-2 'System.Func`2')[CellStyle](CloudyWing.SpreadsheetExporter.CellStyle.md 'CloudyWing.SpreadsheetExporter.CellStyle')[&gt;](https://docs.microsoft.com/en-us/dotnet/api/System.Func-2 'System.Func`2')

The field style generator. The dafault is `(context) => SpreadsheetManager.DefaultCellStyles.FieldStyle`.

<a name='CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.AddChildToLast_TField_(string,string,System.Nullable_CloudyWing.SpreadsheetExporter.CellStyle_,System.Func_CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_T,TField_,CloudyWing.SpreadsheetExporter.CellStyle_)'></a>

## DataColumnCollection<T>.AddChildToLast<TField>(string, string, Nullable<CellStyle>, Func<FieldContext<T,TField>,CellStyle>) Method

Adds the child data column at the end of the last data column.

```csharp
public void AddChildToLast<TField>(string headerText, string fieldKey, System.Nullable<CloudyWing.SpreadsheetExporter.CellStyle> headerStyle=null, System.Func<CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext<T,TField>,CloudyWing.SpreadsheetExporter.CellStyle> fieldStyleGenerator=null);
```
#### Type parameters

<a name='CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.AddChildToLast_TField_(string,string,System.Nullable_CloudyWing.SpreadsheetExporter.CellStyle_,System.Func_CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_T,TField_,CloudyWing.SpreadsheetExporter.CellStyle_).TField'></a>

`TField`

The type of the field.
#### Parameters

<a name='CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.AddChildToLast_TField_(string,string,System.Nullable_CloudyWing.SpreadsheetExporter.CellStyle_,System.Func_CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_T,TField_,CloudyWing.SpreadsheetExporter.CellStyle_).headerText'></a>

`headerText` [System.String](https://docs.microsoft.com/en-us/dotnet/api/System.String 'System.String')

The header text.

<a name='CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.AddChildToLast_TField_(string,string,System.Nullable_CloudyWing.SpreadsheetExporter.CellStyle_,System.Func_CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_T,TField_,CloudyWing.SpreadsheetExporter.CellStyle_).fieldKey'></a>

`fieldKey` [System.String](https://docs.microsoft.com/en-us/dotnet/api/System.String 'System.String')

The field key.

<a name='CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.AddChildToLast_TField_(string,string,System.Nullable_CloudyWing.SpreadsheetExporter.CellStyle_,System.Func_CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_T,TField_,CloudyWing.SpreadsheetExporter.CellStyle_).headerStyle'></a>

`headerStyle` [System.Nullable&lt;](https://docs.microsoft.com/en-us/dotnet/api/System.Nullable-1 'System.Nullable`1')[CellStyle](CloudyWing.SpreadsheetExporter.CellStyle.md 'CloudyWing.SpreadsheetExporter.CellStyle')[&gt;](https://docs.microsoft.com/en-us/dotnet/api/System.Nullable-1 'System.Nullable`1')

The header style. The dafault is `SpreadsheetManager.DefaultCellStyles.HeaderStyle`.

<a name='CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.AddChildToLast_TField_(string,string,System.Nullable_CloudyWing.SpreadsheetExporter.CellStyle_,System.Func_CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_T,TField_,CloudyWing.SpreadsheetExporter.CellStyle_).fieldStyleGenerator'></a>

`fieldStyleGenerator` [System.Func&lt;](https://docs.microsoft.com/en-us/dotnet/api/System.Func-2 'System.Func`2')[CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext&lt;](CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_TRecoed,TField_.md 'CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext<TRecoed,TField>')[T](CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.md#CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.T 'CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection<T>.T')[,](CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_TRecoed,TField_.md 'CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext<TRecoed,TField>')[TField](CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.md#CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.AddChildToLast_TField_(string,string,System.Nullable_CloudyWing.SpreadsheetExporter.CellStyle_,System.Func_CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_T,TField_,CloudyWing.SpreadsheetExporter.CellStyle_).TField 'CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection<T>.AddChildToLast<TField>(string, string, System.Nullable<CloudyWing.SpreadsheetExporter.CellStyle>, System.Func<CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext<T,TField>,CloudyWing.SpreadsheetExporter.CellStyle>).TField')[&gt;](CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_TRecoed,TField_.md 'CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext<TRecoed,TField>')[,](https://docs.microsoft.com/en-us/dotnet/api/System.Func-2 'System.Func`2')[CellStyle](CloudyWing.SpreadsheetExporter.CellStyle.md 'CloudyWing.SpreadsheetExporter.CellStyle')[&gt;](https://docs.microsoft.com/en-us/dotnet/api/System.Func-2 'System.Func`2')

The field style generator. The dafault is `(context) => SpreadsheetManager.DefaultCellStyles.FieldStyle`.

<a name='CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.AddChildToLast_TField_(string,System.Linq.Expressions.Expression_System.Func_T,TField__,System.Action_CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.GeneratorProvider_T,CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_T,TField___,System.Nullable_CloudyWing.SpreadsheetExporter.CellStyle_,System.Func_CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_T,TField_,CloudyWing.SpreadsheetExporter.CellStyle_)'></a>

## DataColumnCollection<T>.AddChildToLast<TField>(string, Expression<Func<T,TField>>, Action<GeneratorProvider<T,FieldContext<T,TField>>>, Nullable<CellStyle>, Func<FieldContext<T,TField>,CellStyle>) Method

Adds the child data column at the end of the last data column.

```csharp
public void AddChildToLast<TField>(string headerText, System.Linq.Expressions.Expression<System.Func<T,TField>> fieldKeyExpression, System.Action<CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection<T>.GeneratorProvider<T,CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext<T,TField>>> providerSetter, System.Nullable<CloudyWing.SpreadsheetExporter.CellStyle> headerStyle=null, System.Func<CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext<T,TField>,CloudyWing.SpreadsheetExporter.CellStyle> fieldStyleGenerator=null);
```
#### Type parameters

<a name='CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.AddChildToLast_TField_(string,System.Linq.Expressions.Expression_System.Func_T,TField__,System.Action_CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.GeneratorProvider_T,CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_T,TField___,System.Nullable_CloudyWing.SpreadsheetExporter.CellStyle_,System.Func_CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_T,TField_,CloudyWing.SpreadsheetExporter.CellStyle_).TField'></a>

`TField`

The type of the field.
#### Parameters

<a name='CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.AddChildToLast_TField_(string,System.Linq.Expressions.Expression_System.Func_T,TField__,System.Action_CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.GeneratorProvider_T,CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_T,TField___,System.Nullable_CloudyWing.SpreadsheetExporter.CellStyle_,System.Func_CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_T,TField_,CloudyWing.SpreadsheetExporter.CellStyle_).headerText'></a>

`headerText` [System.String](https://docs.microsoft.com/en-us/dotnet/api/System.String 'System.String')

The header text.

<a name='CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.AddChildToLast_TField_(string,System.Linq.Expressions.Expression_System.Func_T,TField__,System.Action_CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.GeneratorProvider_T,CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_T,TField___,System.Nullable_CloudyWing.SpreadsheetExporter.CellStyle_,System.Func_CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_T,TField_,CloudyWing.SpreadsheetExporter.CellStyle_).fieldKeyExpression'></a>

`fieldKeyExpression` [System.Linq.Expressions.Expression&lt;](https://docs.microsoft.com/en-us/dotnet/api/System.Linq.Expressions.Expression-1 'System.Linq.Expressions.Expression`1')[System.Func&lt;](https://docs.microsoft.com/en-us/dotnet/api/System.Func-2 'System.Func`2')[T](CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.md#CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.T 'CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection<T>.T')[,](https://docs.microsoft.com/en-us/dotnet/api/System.Func-2 'System.Func`2')[TField](CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.md#CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.AddChildToLast_TField_(string,System.Linq.Expressions.Expression_System.Func_T,TField__,System.Action_CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.GeneratorProvider_T,CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_T,TField___,System.Nullable_CloudyWing.SpreadsheetExporter.CellStyle_,System.Func_CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_T,TField_,CloudyWing.SpreadsheetExporter.CellStyle_).TField 'CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection<T>.AddChildToLast<TField>(string, System.Linq.Expressions.Expression<System.Func<T,TField>>, System.Action<CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection<T>.GeneratorProvider<T,CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext<T,TField>>>, System.Nullable<CloudyWing.SpreadsheetExporter.CellStyle>, System.Func<CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext<T,TField>,CloudyWing.SpreadsheetExporter.CellStyle>).TField')[&gt;](https://docs.microsoft.com/en-us/dotnet/api/System.Func-2 'System.Func`2')[&gt;](https://docs.microsoft.com/en-us/dotnet/api/System.Linq.Expressions.Expression-1 'System.Linq.Expressions.Expression`1')

The field key expression.

<a name='CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.AddChildToLast_TField_(string,System.Linq.Expressions.Expression_System.Func_T,TField__,System.Action_CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.GeneratorProvider_T,CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_T,TField___,System.Nullable_CloudyWing.SpreadsheetExporter.CellStyle_,System.Func_CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_T,TField_,CloudyWing.SpreadsheetExporter.CellStyle_).providerSetter'></a>

`providerSetter` [System.Action&lt;](https://docs.microsoft.com/en-us/dotnet/api/System.Action-1 'System.Action`1')[CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection.GeneratorProvider&lt;](CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.GeneratorProvider_TRecord,TContext_.md 'CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection<T>.GeneratorProvider<TRecord,TContext>')[T](CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.md#CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.T 'CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection<T>.T')[,](CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.GeneratorProvider_TRecord,TContext_.md 'CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection<T>.GeneratorProvider<TRecord,TContext>')[T](CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.md#CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.T 'CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection<T>.T')[,](CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.GeneratorProvider_TRecord,TContext_.md 'CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection<T>.GeneratorProvider<TRecord,TContext>')[CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext&lt;](CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_TRecoed,TField_.md 'CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext<TRecoed,TField>')[T](CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.md#CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.T 'CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection<T>.T')[,](CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_TRecoed,TField_.md 'CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext<TRecoed,TField>')[TField](CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.md#CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.AddChildToLast_TField_(string,System.Linq.Expressions.Expression_System.Func_T,TField__,System.Action_CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.GeneratorProvider_T,CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_T,TField___,System.Nullable_CloudyWing.SpreadsheetExporter.CellStyle_,System.Func_CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_T,TField_,CloudyWing.SpreadsheetExporter.CellStyle_).TField 'CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection<T>.AddChildToLast<TField>(string, System.Linq.Expressions.Expression<System.Func<T,TField>>, System.Action<CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection<T>.GeneratorProvider<T,CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext<T,TField>>>, System.Nullable<CloudyWing.SpreadsheetExporter.CellStyle>, System.Func<CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext<T,TField>,CloudyWing.SpreadsheetExporter.CellStyle>).TField')[&gt;](CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_TRecoed,TField_.md 'CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext<TRecoed,TField>')[&gt;](CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.GeneratorProvider_TRecord,TContext_.md 'CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection<T>.GeneratorProvider<TRecord,TContext>')[&gt;](https://docs.microsoft.com/en-us/dotnet/api/System.Action-1 'System.Action`1')

The provider setter.

<a name='CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.AddChildToLast_TField_(string,System.Linq.Expressions.Expression_System.Func_T,TField__,System.Action_CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.GeneratorProvider_T,CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_T,TField___,System.Nullable_CloudyWing.SpreadsheetExporter.CellStyle_,System.Func_CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_T,TField_,CloudyWing.SpreadsheetExporter.CellStyle_).headerStyle'></a>

`headerStyle` [System.Nullable&lt;](https://docs.microsoft.com/en-us/dotnet/api/System.Nullable-1 'System.Nullable`1')[CellStyle](CloudyWing.SpreadsheetExporter.CellStyle.md 'CloudyWing.SpreadsheetExporter.CellStyle')[&gt;](https://docs.microsoft.com/en-us/dotnet/api/System.Nullable-1 'System.Nullable`1')

The header style. The dafault is `SpreadsheetManager.DefaultCellStyles.HeaderStyle`.

<a name='CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.AddChildToLast_TField_(string,System.Linq.Expressions.Expression_System.Func_T,TField__,System.Action_CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.GeneratorProvider_T,CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_T,TField___,System.Nullable_CloudyWing.SpreadsheetExporter.CellStyle_,System.Func_CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_T,TField_,CloudyWing.SpreadsheetExporter.CellStyle_).fieldStyleGenerator'></a>

`fieldStyleGenerator` [System.Func&lt;](https://docs.microsoft.com/en-us/dotnet/api/System.Func-2 'System.Func`2')[CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext&lt;](CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_TRecoed,TField_.md 'CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext<TRecoed,TField>')[T](CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.md#CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.T 'CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection<T>.T')[,](CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_TRecoed,TField_.md 'CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext<TRecoed,TField>')[TField](CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.md#CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.AddChildToLast_TField_(string,System.Linq.Expressions.Expression_System.Func_T,TField__,System.Action_CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.GeneratorProvider_T,CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_T,TField___,System.Nullable_CloudyWing.SpreadsheetExporter.CellStyle_,System.Func_CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_T,TField_,CloudyWing.SpreadsheetExporter.CellStyle_).TField 'CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection<T>.AddChildToLast<TField>(string, System.Linq.Expressions.Expression<System.Func<T,TField>>, System.Action<CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection<T>.GeneratorProvider<T,CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext<T,TField>>>, System.Nullable<CloudyWing.SpreadsheetExporter.CellStyle>, System.Func<CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext<T,TField>,CloudyWing.SpreadsheetExporter.CellStyle>).TField')[&gt;](CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_TRecoed,TField_.md 'CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext<TRecoed,TField>')[,](https://docs.microsoft.com/en-us/dotnet/api/System.Func-2 'System.Func`2')[CellStyle](CloudyWing.SpreadsheetExporter.CellStyle.md 'CloudyWing.SpreadsheetExporter.CellStyle')[&gt;](https://docs.microsoft.com/en-us/dotnet/api/System.Func-2 'System.Func`2')

The field style generator. The dafault is `(context) => SpreadsheetManager.DefaultCellStyles.FieldStyle`.

<a name='CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.AddChildToLast_TField_(string,System.Linq.Expressions.Expression_System.Func_T,TField__,System.Nullable_CloudyWing.SpreadsheetExporter.CellStyle_,System.Func_CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_T,TField_,CloudyWing.SpreadsheetExporter.CellStyle_)'></a>

## DataColumnCollection<T>.AddChildToLast<TField>(string, Expression<Func<T,TField>>, Nullable<CellStyle>, Func<FieldContext<T,TField>,CellStyle>) Method

Adds the child data column at the end of the last data column.

```csharp
public void AddChildToLast<TField>(string headerText, System.Linq.Expressions.Expression<System.Func<T,TField>> fieldKeyExpression, System.Nullable<CloudyWing.SpreadsheetExporter.CellStyle> headerStyle=null, System.Func<CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext<T,TField>,CloudyWing.SpreadsheetExporter.CellStyle> fieldStyleGenerator=null);
```
#### Type parameters

<a name='CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.AddChildToLast_TField_(string,System.Linq.Expressions.Expression_System.Func_T,TField__,System.Nullable_CloudyWing.SpreadsheetExporter.CellStyle_,System.Func_CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_T,TField_,CloudyWing.SpreadsheetExporter.CellStyle_).TField'></a>

`TField`

The type of the field.
#### Parameters

<a name='CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.AddChildToLast_TField_(string,System.Linq.Expressions.Expression_System.Func_T,TField__,System.Nullable_CloudyWing.SpreadsheetExporter.CellStyle_,System.Func_CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_T,TField_,CloudyWing.SpreadsheetExporter.CellStyle_).headerText'></a>

`headerText` [System.String](https://docs.microsoft.com/en-us/dotnet/api/System.String 'System.String')

The header text.

<a name='CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.AddChildToLast_TField_(string,System.Linq.Expressions.Expression_System.Func_T,TField__,System.Nullable_CloudyWing.SpreadsheetExporter.CellStyle_,System.Func_CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_T,TField_,CloudyWing.SpreadsheetExporter.CellStyle_).fieldKeyExpression'></a>

`fieldKeyExpression` [System.Linq.Expressions.Expression&lt;](https://docs.microsoft.com/en-us/dotnet/api/System.Linq.Expressions.Expression-1 'System.Linq.Expressions.Expression`1')[System.Func&lt;](https://docs.microsoft.com/en-us/dotnet/api/System.Func-2 'System.Func`2')[T](CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.md#CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.T 'CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection<T>.T')[,](https://docs.microsoft.com/en-us/dotnet/api/System.Func-2 'System.Func`2')[TField](CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.md#CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.AddChildToLast_TField_(string,System.Linq.Expressions.Expression_System.Func_T,TField__,System.Nullable_CloudyWing.SpreadsheetExporter.CellStyle_,System.Func_CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_T,TField_,CloudyWing.SpreadsheetExporter.CellStyle_).TField 'CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection<T>.AddChildToLast<TField>(string, System.Linq.Expressions.Expression<System.Func<T,TField>>, System.Nullable<CloudyWing.SpreadsheetExporter.CellStyle>, System.Func<CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext<T,TField>,CloudyWing.SpreadsheetExporter.CellStyle>).TField')[&gt;](https://docs.microsoft.com/en-us/dotnet/api/System.Func-2 'System.Func`2')[&gt;](https://docs.microsoft.com/en-us/dotnet/api/System.Linq.Expressions.Expression-1 'System.Linq.Expressions.Expression`1')

The field key expression.

<a name='CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.AddChildToLast_TField_(string,System.Linq.Expressions.Expression_System.Func_T,TField__,System.Nullable_CloudyWing.SpreadsheetExporter.CellStyle_,System.Func_CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_T,TField_,CloudyWing.SpreadsheetExporter.CellStyle_).headerStyle'></a>

`headerStyle` [System.Nullable&lt;](https://docs.microsoft.com/en-us/dotnet/api/System.Nullable-1 'System.Nullable`1')[CellStyle](CloudyWing.SpreadsheetExporter.CellStyle.md 'CloudyWing.SpreadsheetExporter.CellStyle')[&gt;](https://docs.microsoft.com/en-us/dotnet/api/System.Nullable-1 'System.Nullable`1')

The header style. The dafault is `SpreadsheetManager.DefaultCellStyles.HeaderStyle`.

<a name='CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.AddChildToLast_TField_(string,System.Linq.Expressions.Expression_System.Func_T,TField__,System.Nullable_CloudyWing.SpreadsheetExporter.CellStyle_,System.Func_CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_T,TField_,CloudyWing.SpreadsheetExporter.CellStyle_).fieldStyleGenerator'></a>

`fieldStyleGenerator` [System.Func&lt;](https://docs.microsoft.com/en-us/dotnet/api/System.Func-2 'System.Func`2')[CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext&lt;](CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_TRecoed,TField_.md 'CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext<TRecoed,TField>')[T](CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.md#CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.T 'CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection<T>.T')[,](CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_TRecoed,TField_.md 'CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext<TRecoed,TField>')[TField](CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.md#CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.AddChildToLast_TField_(string,System.Linq.Expressions.Expression_System.Func_T,TField__,System.Nullable_CloudyWing.SpreadsheetExporter.CellStyle_,System.Func_CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_T,TField_,CloudyWing.SpreadsheetExporter.CellStyle_).TField 'CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection<T>.AddChildToLast<TField>(string, System.Linq.Expressions.Expression<System.Func<T,TField>>, System.Nullable<CloudyWing.SpreadsheetExporter.CellStyle>, System.Func<CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext<T,TField>,CloudyWing.SpreadsheetExporter.CellStyle>).TField')[&gt;](CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_TRecoed,TField_.md 'CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext<TRecoed,TField>')[,](https://docs.microsoft.com/en-us/dotnet/api/System.Func-2 'System.Func`2')[CellStyle](CloudyWing.SpreadsheetExporter.CellStyle.md 'CloudyWing.SpreadsheetExporter.CellStyle')[&gt;](https://docs.microsoft.com/en-us/dotnet/api/System.Func-2 'System.Func`2')

The field style generator. The dafault is `(context) => SpreadsheetManager.DefaultCellStyles.FieldStyle`.

<a name='CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.ResetColumnsPoint(System.Drawing.Point)'></a>

## DataColumnCollection<T>.ResetColumnsPoint(Point) Method

重設底下所有DataColumn的座標

```csharp
internal void ResetColumnsPoint(System.Drawing.Point point);
```
#### Parameters

<a name='CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.ResetColumnsPoint(System.Drawing.Point).point'></a>

`point` [System.Drawing.Point](https://docs.microsoft.com/en-us/dotnet/api/System.Drawing.Point 'System.Drawing.Point')

<a name='CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.ResetRootPoint()'></a>

## DataColumnCollection<T>.ResetRootPoint() Method

從根節點往下重設座標

```csharp
internal void ResetRootPoint();
```