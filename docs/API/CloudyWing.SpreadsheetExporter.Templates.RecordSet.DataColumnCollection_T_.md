#### [CloudyWing\.SpreadsheetExporter](index.md 'index')
### [CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet](CloudyWing.SpreadsheetExporter.Templates.RecordSet.md 'CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet')

## DataColumnCollection\<T\> Class

The data column collection\.

```csharp
public class DataColumnCollection<T> : System.Collections.ObjectModel.Collection<CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnBase<T>>
```
#### Type parameters

<a name='CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.T'></a>

`T`

The type of the record\.

Inheritance [System\.Object](https://learn.microsoft.com/en-us/dotnet/api/system.object 'System\.Object') &#129106; [System\.Collections\.ObjectModel\.Collection&lt;](https://learn.microsoft.com/en-us/dotnet/api/system.collections.objectmodel.collection-1 'System\.Collections\.ObjectModel\.Collection\`1')[CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.DataColumnBase&lt;](CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnBase_T_.md 'CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.DataColumnBase\<T\>')[T](CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.md#CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.T 'CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.DataColumnCollection\<T\>\.T')[&gt;](CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnBase_T_.md 'CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.DataColumnBase\<T\>')[&gt;](https://learn.microsoft.com/en-us/dotnet/api/system.collections.objectmodel.collection-1 'System\.Collections\.ObjectModel\.Collection\`1') &#129106; DataColumnCollection\<T\>
### Properties

<a name='CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.ColumnSpan'></a>

## DataColumnCollection\<T\>\.ColumnSpan Property

Gets the column span\.

```csharp
public int ColumnSpan { get; }
```

#### Property Value
[System\.Int32](https://learn.microsoft.com/en-us/dotnet/api/system.int32 'System\.Int32')  
The column span\.

<a name='CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.DataSourceColumns'></a>

## DataColumnCollection\<T\>\.DataSourceColumns Property

Get the column containing the properties of the data source\.

```csharp
public System.Collections.Generic.IEnumerable<CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnBase<T>> DataSourceColumns { get; }
```

#### Property Value
[System\.Collections\.Generic\.IEnumerable&lt;](https://learn.microsoft.com/en-us/dotnet/api/system.collections.generic.ienumerable-1 'System\.Collections\.Generic\.IEnumerable\`1')[CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.DataColumnBase&lt;](CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnBase_T_.md 'CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.DataColumnBase\<T\>')[T](CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.md#CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.T 'CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.DataColumnCollection\<T\>\.T')[&gt;](CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnBase_T_.md 'CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.DataColumnBase\<T\>')[&gt;](https://learn.microsoft.com/en-us/dotnet/api/system.collections.generic.ienumerable-1 'System\.Collections\.Generic\.IEnumerable\`1')  
The data columns\.

<a name='CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.RootColumns'></a>

## DataColumnCollection\<T\>\.RootColumns Property

Gets the root columns\.

```csharp
public CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection<T> RootColumns { get; }
```

#### Property Value
[CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.DataColumnCollection&lt;](CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.md 'CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.DataColumnCollection\<T\>')[T](CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.md#CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.T 'CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.DataColumnCollection\<T\>\.T')[&gt;](CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.md 'CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.DataColumnCollection\<T\>')  
The root columns\.

<a name='CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.RowSpan'></a>

## DataColumnCollection\<T\>\.RowSpan Property

Gets the row span\.

```csharp
public int RowSpan { get; }
```

#### Property Value
[System\.Int32](https://learn.microsoft.com/en-us/dotnet/api/system.int32 'System\.Int32')  
The row span\.
### Methods

<a name='CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.Add(CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnBase_T_)'></a>

## DataColumnCollection\<T\>\.Add\(DataColumnBase\<T\>\) Method

Adds the specified header text\.

```csharp
public CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection<T> Add(CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnBase<T> dataColumn);
```
#### Parameters

<a name='CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.Add(CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnBase_T_).dataColumn'></a>

`dataColumn` [CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.DataColumnBase&lt;](CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnBase_T_.md 'CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.DataColumnBase\<T\>')[T](CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.md#CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.T 'CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.DataColumnCollection\<T\>\.T')[&gt;](CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnBase_T_.md 'CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.DataColumnBase\<T\>')

The data column\.

#### Returns
[CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.DataColumnCollection&lt;](CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.md 'CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.DataColumnCollection\<T\>')[T](CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.md#CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.T 'CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.DataColumnCollection\<T\>\.T')[&gt;](CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.md 'CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.DataColumnCollection\<T\>')  
The self\.

<a name='CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.Add(string,System.Action_CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.GeneratorProvider_T,CloudyWing.SpreadsheetExporter.Templates.RecordSet.RecordContext_T___,System.Nullable_CloudyWing.SpreadsheetExporter.CellStyle_,System.Func_CloudyWing.SpreadsheetExporter.Templates.RecordSet.RecordContext_T_,CloudyWing.SpreadsheetExporter.CellStyle_)'></a>

## DataColumnCollection\<T\>\.Add\(string, Action\<GeneratorProvider\<T,RecordContext\<T\>\>\>, Nullable\<CellStyle\>, Func\<RecordContext\<T\>,CellStyle\>\) Method

Adds the data column to the end of the DataColumnCollection\<T\>\.

```csharp
public CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection<T> Add(string headerText, System.Action<CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection<T>.GeneratorProvider<T,CloudyWing.SpreadsheetExporter.Templates.RecordSet.RecordContext<T>>> providerSetter, System.Nullable<CloudyWing.SpreadsheetExporter.CellStyle> headerStyle=null, System.Func<CloudyWing.SpreadsheetExporter.Templates.RecordSet.RecordContext<T>,CloudyWing.SpreadsheetExporter.CellStyle> fieldStyleGenerator=null);
```
#### Parameters

<a name='CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.Add(string,System.Action_CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.GeneratorProvider_T,CloudyWing.SpreadsheetExporter.Templates.RecordSet.RecordContext_T___,System.Nullable_CloudyWing.SpreadsheetExporter.CellStyle_,System.Func_CloudyWing.SpreadsheetExporter.Templates.RecordSet.RecordContext_T_,CloudyWing.SpreadsheetExporter.CellStyle_).headerText'></a>

`headerText` [System\.String](https://learn.microsoft.com/en-us/dotnet/api/system.string 'System\.String')

The header text\.

<a name='CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.Add(string,System.Action_CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.GeneratorProvider_T,CloudyWing.SpreadsheetExporter.Templates.RecordSet.RecordContext_T___,System.Nullable_CloudyWing.SpreadsheetExporter.CellStyle_,System.Func_CloudyWing.SpreadsheetExporter.Templates.RecordSet.RecordContext_T_,CloudyWing.SpreadsheetExporter.CellStyle_).providerSetter'></a>

`providerSetter` [System\.Action&lt;](https://learn.microsoft.com/en-us/dotnet/api/system.action-1 'System\.Action\`1')[CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.DataColumnCollection\.GeneratorProvider&lt;](CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.GeneratorProvider_TRecord,TContext_.md 'CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.DataColumnCollection\<T\>\.GeneratorProvider\<TRecord,TContext\>')[T](CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.md#CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.T 'CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.DataColumnCollection\<T\>\.T')[,](CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.GeneratorProvider_TRecord,TContext_.md 'CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.DataColumnCollection\<T\>\.GeneratorProvider\<TRecord,TContext\>')[T](CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.md#CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.T 'CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.DataColumnCollection\<T\>\.T')[,](CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.GeneratorProvider_TRecord,TContext_.md 'CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.DataColumnCollection\<T\>\.GeneratorProvider\<TRecord,TContext\>')[CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.RecordContext&lt;](CloudyWing.SpreadsheetExporter.Templates.RecordSet.RecordContext_T_.md 'CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.RecordContext\<T\>')[T](CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.md#CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.T 'CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.DataColumnCollection\<T\>\.T')[&gt;](CloudyWing.SpreadsheetExporter.Templates.RecordSet.RecordContext_T_.md 'CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.RecordContext\<T\>')[&gt;](CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.GeneratorProvider_TRecord,TContext_.md 'CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.DataColumnCollection\<T\>\.GeneratorProvider\<TRecord,TContext\>')[&gt;](https://learn.microsoft.com/en-us/dotnet/api/system.action-1 'System\.Action\`1')

The provider setter\.

<a name='CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.Add(string,System.Action_CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.GeneratorProvider_T,CloudyWing.SpreadsheetExporter.Templates.RecordSet.RecordContext_T___,System.Nullable_CloudyWing.SpreadsheetExporter.CellStyle_,System.Func_CloudyWing.SpreadsheetExporter.Templates.RecordSet.RecordContext_T_,CloudyWing.SpreadsheetExporter.CellStyle_).headerStyle'></a>

`headerStyle` [System\.Nullable&lt;](https://learn.microsoft.com/en-us/dotnet/api/system.nullable-1 'System\.Nullable\`1')[CellStyle](CloudyWing.SpreadsheetExporter.CellStyle.md 'CloudyWing\.SpreadsheetExporter\.CellStyle')[&gt;](https://learn.microsoft.com/en-us/dotnet/api/system.nullable-1 'System\.Nullable\`1')

The header style\. The dafault is `SpreadsheetManager.DefaultCellStyles.HeaderStyle`\.

<a name='CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.Add(string,System.Action_CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.GeneratorProvider_T,CloudyWing.SpreadsheetExporter.Templates.RecordSet.RecordContext_T___,System.Nullable_CloudyWing.SpreadsheetExporter.CellStyle_,System.Func_CloudyWing.SpreadsheetExporter.Templates.RecordSet.RecordContext_T_,CloudyWing.SpreadsheetExporter.CellStyle_).fieldStyleGenerator'></a>

`fieldStyleGenerator` [System\.Func&lt;](https://learn.microsoft.com/en-us/dotnet/api/system.func-2 'System\.Func\`2')[CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.RecordContext&lt;](CloudyWing.SpreadsheetExporter.Templates.RecordSet.RecordContext_T_.md 'CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.RecordContext\<T\>')[T](CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.md#CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.T 'CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.DataColumnCollection\<T\>\.T')[&gt;](CloudyWing.SpreadsheetExporter.Templates.RecordSet.RecordContext_T_.md 'CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.RecordContext\<T\>')[,](https://learn.microsoft.com/en-us/dotnet/api/system.func-2 'System\.Func\`2')[CellStyle](CloudyWing.SpreadsheetExporter.CellStyle.md 'CloudyWing\.SpreadsheetExporter\.CellStyle')[&gt;](https://learn.microsoft.com/en-us/dotnet/api/system.func-2 'System\.Func\`2')

The field style generator\. The dafault is `(context) => SpreadsheetManager.DefaultCellStyles.FieldStyle`\.

#### Returns
[CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.DataColumnCollection&lt;](CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.md 'CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.DataColumnCollection\<T\>')[T](CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.md#CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.T 'CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.DataColumnCollection\<T\>\.T')[&gt;](CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.md 'CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.DataColumnCollection\<T\>')  
The self\.

<a name='CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.Add(string,System.Nullable_CloudyWing.SpreadsheetExporter.CellStyle_,System.Func_CloudyWing.SpreadsheetExporter.Templates.RecordSet.RecordContext_T_,CloudyWing.SpreadsheetExporter.CellStyle_)'></a>

## DataColumnCollection\<T\>\.Add\(string, Nullable\<CellStyle\>, Func\<RecordContext\<T\>,CellStyle\>\) Method

Adds the specified header text\.

```csharp
public CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection<T> Add(string headerText, System.Nullable<CloudyWing.SpreadsheetExporter.CellStyle> headerStyle=null, System.Func<CloudyWing.SpreadsheetExporter.Templates.RecordSet.RecordContext<T>,CloudyWing.SpreadsheetExporter.CellStyle> fieldStyleGenerator=null);
```
#### Parameters

<a name='CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.Add(string,System.Nullable_CloudyWing.SpreadsheetExporter.CellStyle_,System.Func_CloudyWing.SpreadsheetExporter.Templates.RecordSet.RecordContext_T_,CloudyWing.SpreadsheetExporter.CellStyle_).headerText'></a>

`headerText` [System\.String](https://learn.microsoft.com/en-us/dotnet/api/system.string 'System\.String')

The header text\.

<a name='CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.Add(string,System.Nullable_CloudyWing.SpreadsheetExporter.CellStyle_,System.Func_CloudyWing.SpreadsheetExporter.Templates.RecordSet.RecordContext_T_,CloudyWing.SpreadsheetExporter.CellStyle_).headerStyle'></a>

`headerStyle` [System\.Nullable&lt;](https://learn.microsoft.com/en-us/dotnet/api/system.nullable-1 'System\.Nullable\`1')[CellStyle](CloudyWing.SpreadsheetExporter.CellStyle.md 'CloudyWing\.SpreadsheetExporter\.CellStyle')[&gt;](https://learn.microsoft.com/en-us/dotnet/api/system.nullable-1 'System\.Nullable\`1')

The header style\.

<a name='CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.Add(string,System.Nullable_CloudyWing.SpreadsheetExporter.CellStyle_,System.Func_CloudyWing.SpreadsheetExporter.Templates.RecordSet.RecordContext_T_,CloudyWing.SpreadsheetExporter.CellStyle_).fieldStyleGenerator'></a>

`fieldStyleGenerator` [System\.Func&lt;](https://learn.microsoft.com/en-us/dotnet/api/system.func-2 'System\.Func\`2')[CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.RecordContext&lt;](CloudyWing.SpreadsheetExporter.Templates.RecordSet.RecordContext_T_.md 'CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.RecordContext\<T\>')[T](CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.md#CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.T 'CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.DataColumnCollection\<T\>\.T')[&gt;](CloudyWing.SpreadsheetExporter.Templates.RecordSet.RecordContext_T_.md 'CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.RecordContext\<T\>')[,](https://learn.microsoft.com/en-us/dotnet/api/system.func-2 'System\.Func\`2')[CellStyle](CloudyWing.SpreadsheetExporter.CellStyle.md 'CloudyWing\.SpreadsheetExporter\.CellStyle')[&gt;](https://learn.microsoft.com/en-us/dotnet/api/system.func-2 'System\.Func\`2')

The field style generator\.

#### Returns
[CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.DataColumnCollection&lt;](CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.md 'CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.DataColumnCollection\<T\>')[T](CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.md#CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.T 'CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.DataColumnCollection\<T\>\.T')[&gt;](CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.md 'CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.DataColumnCollection\<T\>')  
The self\.

<a name='CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.Add_TField_(string,string,System.Action_CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.GeneratorProvider_T,CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_T,TField___,System.Nullable_CloudyWing.SpreadsheetExporter.CellStyle_,System.Func_CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_T,TField_,CloudyWing.SpreadsheetExporter.CellStyle_)'></a>

## DataColumnCollection\<T\>\.Add\<TField\>\(string, string, Action\<GeneratorProvider\<T,FieldContext\<T,TField\>\>\>, Nullable\<CellStyle\>, Func\<FieldContext\<T,TField\>,CellStyle\>\) Method

Adds the data column to the end of the DataColumnCollection\<T\>\.

```csharp
public CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection<T> Add<TField>(string headerText, string fieldKey, System.Action<CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection<T>.GeneratorProvider<T,CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext<T,TField>>> providerSetter, System.Nullable<CloudyWing.SpreadsheetExporter.CellStyle> headerStyle=null, System.Func<CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext<T,TField>,CloudyWing.SpreadsheetExporter.CellStyle> fieldStyleGenerator=null);
```
#### Type parameters

<a name='CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.Add_TField_(string,string,System.Action_CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.GeneratorProvider_T,CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_T,TField___,System.Nullable_CloudyWing.SpreadsheetExporter.CellStyle_,System.Func_CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_T,TField_,CloudyWing.SpreadsheetExporter.CellStyle_).TField'></a>

`TField`

The type of the field\.
#### Parameters

<a name='CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.Add_TField_(string,string,System.Action_CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.GeneratorProvider_T,CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_T,TField___,System.Nullable_CloudyWing.SpreadsheetExporter.CellStyle_,System.Func_CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_T,TField_,CloudyWing.SpreadsheetExporter.CellStyle_).headerText'></a>

`headerText` [System\.String](https://learn.microsoft.com/en-us/dotnet/api/system.string 'System\.String')

The header text\.

<a name='CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.Add_TField_(string,string,System.Action_CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.GeneratorProvider_T,CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_T,TField___,System.Nullable_CloudyWing.SpreadsheetExporter.CellStyle_,System.Func_CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_T,TField_,CloudyWing.SpreadsheetExporter.CellStyle_).fieldKey'></a>

`fieldKey` [System\.String](https://learn.microsoft.com/en-us/dotnet/api/system.string 'System\.String')

The field key\.

<a name='CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.Add_TField_(string,string,System.Action_CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.GeneratorProvider_T,CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_T,TField___,System.Nullable_CloudyWing.SpreadsheetExporter.CellStyle_,System.Func_CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_T,TField_,CloudyWing.SpreadsheetExporter.CellStyle_).providerSetter'></a>

`providerSetter` [System\.Action&lt;](https://learn.microsoft.com/en-us/dotnet/api/system.action-1 'System\.Action\`1')[CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.DataColumnCollection\.GeneratorProvider&lt;](CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.GeneratorProvider_TRecord,TContext_.md 'CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.DataColumnCollection\<T\>\.GeneratorProvider\<TRecord,TContext\>')[T](CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.md#CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.T 'CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.DataColumnCollection\<T\>\.T')[,](CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.GeneratorProvider_TRecord,TContext_.md 'CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.DataColumnCollection\<T\>\.GeneratorProvider\<TRecord,TContext\>')[T](CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.md#CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.T 'CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.DataColumnCollection\<T\>\.T')[,](CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.GeneratorProvider_TRecord,TContext_.md 'CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.DataColumnCollection\<T\>\.GeneratorProvider\<TRecord,TContext\>')[CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.FieldContext&lt;](CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_TRecord,TField_.md 'CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.FieldContext\<TRecord,TField\>')[T](CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.md#CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.T 'CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.DataColumnCollection\<T\>\.T')[,](CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_TRecord,TField_.md 'CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.FieldContext\<TRecord,TField\>')[TField](CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.md#CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.Add_TField_(string,string,System.Action_CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.GeneratorProvider_T,CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_T,TField___,System.Nullable_CloudyWing.SpreadsheetExporter.CellStyle_,System.Func_CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_T,TField_,CloudyWing.SpreadsheetExporter.CellStyle_).TField 'CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.DataColumnCollection\<T\>\.Add\<TField\>\(string, string, System\.Action\<CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.DataColumnCollection\<T\>\.GeneratorProvider\<T,CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.FieldContext\<T,TField\>\>\>, System\.Nullable\<CloudyWing\.SpreadsheetExporter\.CellStyle\>, System\.Func\<CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.FieldContext\<T,TField\>,CloudyWing\.SpreadsheetExporter\.CellStyle\>\)\.TField')[&gt;](CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_TRecord,TField_.md 'CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.FieldContext\<TRecord,TField\>')[&gt;](CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.GeneratorProvider_TRecord,TContext_.md 'CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.DataColumnCollection\<T\>\.GeneratorProvider\<TRecord,TContext\>')[&gt;](https://learn.microsoft.com/en-us/dotnet/api/system.action-1 'System\.Action\`1')

The provider setter\.

<a name='CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.Add_TField_(string,string,System.Action_CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.GeneratorProvider_T,CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_T,TField___,System.Nullable_CloudyWing.SpreadsheetExporter.CellStyle_,System.Func_CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_T,TField_,CloudyWing.SpreadsheetExporter.CellStyle_).headerStyle'></a>

`headerStyle` [System\.Nullable&lt;](https://learn.microsoft.com/en-us/dotnet/api/system.nullable-1 'System\.Nullable\`1')[CellStyle](CloudyWing.SpreadsheetExporter.CellStyle.md 'CloudyWing\.SpreadsheetExporter\.CellStyle')[&gt;](https://learn.microsoft.com/en-us/dotnet/api/system.nullable-1 'System\.Nullable\`1')

The header style\. The dafault is `SpreadsheetManager.DefaultCellStyles.HeaderStyle`\.

<a name='CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.Add_TField_(string,string,System.Action_CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.GeneratorProvider_T,CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_T,TField___,System.Nullable_CloudyWing.SpreadsheetExporter.CellStyle_,System.Func_CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_T,TField_,CloudyWing.SpreadsheetExporter.CellStyle_).fieldStyleGenerator'></a>

`fieldStyleGenerator` [System\.Func&lt;](https://learn.microsoft.com/en-us/dotnet/api/system.func-2 'System\.Func\`2')[CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.FieldContext&lt;](CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_TRecord,TField_.md 'CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.FieldContext\<TRecord,TField\>')[T](CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.md#CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.T 'CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.DataColumnCollection\<T\>\.T')[,](CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_TRecord,TField_.md 'CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.FieldContext\<TRecord,TField\>')[TField](CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.md#CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.Add_TField_(string,string,System.Action_CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.GeneratorProvider_T,CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_T,TField___,System.Nullable_CloudyWing.SpreadsheetExporter.CellStyle_,System.Func_CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_T,TField_,CloudyWing.SpreadsheetExporter.CellStyle_).TField 'CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.DataColumnCollection\<T\>\.Add\<TField\>\(string, string, System\.Action\<CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.DataColumnCollection\<T\>\.GeneratorProvider\<T,CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.FieldContext\<T,TField\>\>\>, System\.Nullable\<CloudyWing\.SpreadsheetExporter\.CellStyle\>, System\.Func\<CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.FieldContext\<T,TField\>,CloudyWing\.SpreadsheetExporter\.CellStyle\>\)\.TField')[&gt;](CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_TRecord,TField_.md 'CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.FieldContext\<TRecord,TField\>')[,](https://learn.microsoft.com/en-us/dotnet/api/system.func-2 'System\.Func\`2')[CellStyle](CloudyWing.SpreadsheetExporter.CellStyle.md 'CloudyWing\.SpreadsheetExporter\.CellStyle')[&gt;](https://learn.microsoft.com/en-us/dotnet/api/system.func-2 'System\.Func\`2')

The field style generator\. The dafault is `(context) => SpreadsheetManager.DefaultCellStyles.FieldStyle`\.

#### Returns
[CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.DataColumnCollection&lt;](CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.md 'CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.DataColumnCollection\<T\>')[T](CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.md#CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.T 'CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.DataColumnCollection\<T\>\.T')[&gt;](CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.md 'CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.DataColumnCollection\<T\>')  
The self\.

<a name='CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.Add_TField_(string,string,System.Nullable_CloudyWing.SpreadsheetExporter.CellStyle_,System.Func_CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_T,TField_,CloudyWing.SpreadsheetExporter.CellStyle_)'></a>

## DataColumnCollection\<T\>\.Add\<TField\>\(string, string, Nullable\<CellStyle\>, Func\<FieldContext\<T,TField\>,CellStyle\>\) Method

Adds the data column to the end of the DataColumnCollection\<T\>\.

```csharp
public CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection<T> Add<TField>(string headerText, string fieldKey, System.Nullable<CloudyWing.SpreadsheetExporter.CellStyle> headerStyle=null, System.Func<CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext<T,TField>,CloudyWing.SpreadsheetExporter.CellStyle> fieldStyleGenerator=null);
```
#### Type parameters

<a name='CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.Add_TField_(string,string,System.Nullable_CloudyWing.SpreadsheetExporter.CellStyle_,System.Func_CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_T,TField_,CloudyWing.SpreadsheetExporter.CellStyle_).TField'></a>

`TField`

The type of the field\.
#### Parameters

<a name='CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.Add_TField_(string,string,System.Nullable_CloudyWing.SpreadsheetExporter.CellStyle_,System.Func_CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_T,TField_,CloudyWing.SpreadsheetExporter.CellStyle_).headerText'></a>

`headerText` [System\.String](https://learn.microsoft.com/en-us/dotnet/api/system.string 'System\.String')

The header text\.

<a name='CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.Add_TField_(string,string,System.Nullable_CloudyWing.SpreadsheetExporter.CellStyle_,System.Func_CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_T,TField_,CloudyWing.SpreadsheetExporter.CellStyle_).fieldKey'></a>

`fieldKey` [System\.String](https://learn.microsoft.com/en-us/dotnet/api/system.string 'System\.String')

The field key\.

<a name='CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.Add_TField_(string,string,System.Nullable_CloudyWing.SpreadsheetExporter.CellStyle_,System.Func_CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_T,TField_,CloudyWing.SpreadsheetExporter.CellStyle_).headerStyle'></a>

`headerStyle` [System\.Nullable&lt;](https://learn.microsoft.com/en-us/dotnet/api/system.nullable-1 'System\.Nullable\`1')[CellStyle](CloudyWing.SpreadsheetExporter.CellStyle.md 'CloudyWing\.SpreadsheetExporter\.CellStyle')[&gt;](https://learn.microsoft.com/en-us/dotnet/api/system.nullable-1 'System\.Nullable\`1')

The header style\. The dafault is `SpreadsheetManager.DefaultCellStyles.HeaderStyle`\.

<a name='CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.Add_TField_(string,string,System.Nullable_CloudyWing.SpreadsheetExporter.CellStyle_,System.Func_CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_T,TField_,CloudyWing.SpreadsheetExporter.CellStyle_).fieldStyleGenerator'></a>

`fieldStyleGenerator` [System\.Func&lt;](https://learn.microsoft.com/en-us/dotnet/api/system.func-2 'System\.Func\`2')[CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.FieldContext&lt;](CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_TRecord,TField_.md 'CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.FieldContext\<TRecord,TField\>')[T](CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.md#CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.T 'CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.DataColumnCollection\<T\>\.T')[,](CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_TRecord,TField_.md 'CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.FieldContext\<TRecord,TField\>')[TField](CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.md#CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.Add_TField_(string,string,System.Nullable_CloudyWing.SpreadsheetExporter.CellStyle_,System.Func_CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_T,TField_,CloudyWing.SpreadsheetExporter.CellStyle_).TField 'CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.DataColumnCollection\<T\>\.Add\<TField\>\(string, string, System\.Nullable\<CloudyWing\.SpreadsheetExporter\.CellStyle\>, System\.Func\<CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.FieldContext\<T,TField\>,CloudyWing\.SpreadsheetExporter\.CellStyle\>\)\.TField')[&gt;](CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_TRecord,TField_.md 'CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.FieldContext\<TRecord,TField\>')[,](https://learn.microsoft.com/en-us/dotnet/api/system.func-2 'System\.Func\`2')[CellStyle](CloudyWing.SpreadsheetExporter.CellStyle.md 'CloudyWing\.SpreadsheetExporter\.CellStyle')[&gt;](https://learn.microsoft.com/en-us/dotnet/api/system.func-2 'System\.Func\`2')

The field style generator\. The dafault is `(context) => SpreadsheetManager.DefaultCellStyles.FieldStyle`\.

#### Returns
[CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.DataColumnCollection&lt;](CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.md 'CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.DataColumnCollection\<T\>')[T](CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.md#CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.T 'CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.DataColumnCollection\<T\>\.T')[&gt;](CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.md 'CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.DataColumnCollection\<T\>')  
The self\.

<a name='CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.Add_TField_(string,System.Linq.Expressions.Expression_System.Func_T,TField__,System.Action_CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.GeneratorProvider_T,CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_T,TField___,System.Nullable_CloudyWing.SpreadsheetExporter.CellStyle_,System.Func_CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_T,TField_,CloudyWing.SpreadsheetExporter.CellStyle_)'></a>

## DataColumnCollection\<T\>\.Add\<TField\>\(string, Expression\<Func\<T,TField\>\>, Action\<GeneratorProvider\<T,FieldContext\<T,TField\>\>\>, Nullable\<CellStyle\>, Func\<FieldContext\<T,TField\>,CellStyle\>\) Method

Adds the data column to the end of the DataColumnCollection\<T\>\.

```csharp
public CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection<T> Add<TField>(string headerText, System.Linq.Expressions.Expression<System.Func<T,TField>> fieldKeyExpression, System.Action<CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection<T>.GeneratorProvider<T,CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext<T,TField>>> providerSetter, System.Nullable<CloudyWing.SpreadsheetExporter.CellStyle> headerStyle=null, System.Func<CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext<T,TField>,CloudyWing.SpreadsheetExporter.CellStyle> fieldStyleGenerator=null);
```
#### Type parameters

<a name='CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.Add_TField_(string,System.Linq.Expressions.Expression_System.Func_T,TField__,System.Action_CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.GeneratorProvider_T,CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_T,TField___,System.Nullable_CloudyWing.SpreadsheetExporter.CellStyle_,System.Func_CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_T,TField_,CloudyWing.SpreadsheetExporter.CellStyle_).TField'></a>

`TField`

The type of the field\.
#### Parameters

<a name='CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.Add_TField_(string,System.Linq.Expressions.Expression_System.Func_T,TField__,System.Action_CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.GeneratorProvider_T,CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_T,TField___,System.Nullable_CloudyWing.SpreadsheetExporter.CellStyle_,System.Func_CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_T,TField_,CloudyWing.SpreadsheetExporter.CellStyle_).headerText'></a>

`headerText` [System\.String](https://learn.microsoft.com/en-us/dotnet/api/system.string 'System\.String')

The header text\.

<a name='CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.Add_TField_(string,System.Linq.Expressions.Expression_System.Func_T,TField__,System.Action_CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.GeneratorProvider_T,CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_T,TField___,System.Nullable_CloudyWing.SpreadsheetExporter.CellStyle_,System.Func_CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_T,TField_,CloudyWing.SpreadsheetExporter.CellStyle_).fieldKeyExpression'></a>

`fieldKeyExpression` [System\.Linq\.Expressions\.Expression&lt;](https://learn.microsoft.com/en-us/dotnet/api/system.linq.expressions.expression-1 'System\.Linq\.Expressions\.Expression\`1')[System\.Func&lt;](https://learn.microsoft.com/en-us/dotnet/api/system.func-2 'System\.Func\`2')[T](CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.md#CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.T 'CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.DataColumnCollection\<T\>\.T')[,](https://learn.microsoft.com/en-us/dotnet/api/system.func-2 'System\.Func\`2')[TField](CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.md#CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.Add_TField_(string,System.Linq.Expressions.Expression_System.Func_T,TField__,System.Action_CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.GeneratorProvider_T,CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_T,TField___,System.Nullable_CloudyWing.SpreadsheetExporter.CellStyle_,System.Func_CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_T,TField_,CloudyWing.SpreadsheetExporter.CellStyle_).TField 'CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.DataColumnCollection\<T\>\.Add\<TField\>\(string, System\.Linq\.Expressions\.Expression\<System\.Func\<T,TField\>\>, System\.Action\<CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.DataColumnCollection\<T\>\.GeneratorProvider\<T,CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.FieldContext\<T,TField\>\>\>, System\.Nullable\<CloudyWing\.SpreadsheetExporter\.CellStyle\>, System\.Func\<CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.FieldContext\<T,TField\>,CloudyWing\.SpreadsheetExporter\.CellStyle\>\)\.TField')[&gt;](https://learn.microsoft.com/en-us/dotnet/api/system.func-2 'System\.Func\`2')[&gt;](https://learn.microsoft.com/en-us/dotnet/api/system.linq.expressions.expression-1 'System\.Linq\.Expressions\.Expression\`1')

The field key expression\.

<a name='CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.Add_TField_(string,System.Linq.Expressions.Expression_System.Func_T,TField__,System.Action_CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.GeneratorProvider_T,CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_T,TField___,System.Nullable_CloudyWing.SpreadsheetExporter.CellStyle_,System.Func_CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_T,TField_,CloudyWing.SpreadsheetExporter.CellStyle_).providerSetter'></a>

`providerSetter` [System\.Action&lt;](https://learn.microsoft.com/en-us/dotnet/api/system.action-1 'System\.Action\`1')[CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.DataColumnCollection\.GeneratorProvider&lt;](CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.GeneratorProvider_TRecord,TContext_.md 'CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.DataColumnCollection\<T\>\.GeneratorProvider\<TRecord,TContext\>')[T](CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.md#CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.T 'CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.DataColumnCollection\<T\>\.T')[,](CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.GeneratorProvider_TRecord,TContext_.md 'CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.DataColumnCollection\<T\>\.GeneratorProvider\<TRecord,TContext\>')[T](CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.md#CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.T 'CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.DataColumnCollection\<T\>\.T')[,](CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.GeneratorProvider_TRecord,TContext_.md 'CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.DataColumnCollection\<T\>\.GeneratorProvider\<TRecord,TContext\>')[CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.FieldContext&lt;](CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_TRecord,TField_.md 'CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.FieldContext\<TRecord,TField\>')[T](CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.md#CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.T 'CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.DataColumnCollection\<T\>\.T')[,](CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_TRecord,TField_.md 'CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.FieldContext\<TRecord,TField\>')[TField](CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.md#CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.Add_TField_(string,System.Linq.Expressions.Expression_System.Func_T,TField__,System.Action_CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.GeneratorProvider_T,CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_T,TField___,System.Nullable_CloudyWing.SpreadsheetExporter.CellStyle_,System.Func_CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_T,TField_,CloudyWing.SpreadsheetExporter.CellStyle_).TField 'CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.DataColumnCollection\<T\>\.Add\<TField\>\(string, System\.Linq\.Expressions\.Expression\<System\.Func\<T,TField\>\>, System\.Action\<CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.DataColumnCollection\<T\>\.GeneratorProvider\<T,CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.FieldContext\<T,TField\>\>\>, System\.Nullable\<CloudyWing\.SpreadsheetExporter\.CellStyle\>, System\.Func\<CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.FieldContext\<T,TField\>,CloudyWing\.SpreadsheetExporter\.CellStyle\>\)\.TField')[&gt;](CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_TRecord,TField_.md 'CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.FieldContext\<TRecord,TField\>')[&gt;](CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.GeneratorProvider_TRecord,TContext_.md 'CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.DataColumnCollection\<T\>\.GeneratorProvider\<TRecord,TContext\>')[&gt;](https://learn.microsoft.com/en-us/dotnet/api/system.action-1 'System\.Action\`1')

The provider setter\.

<a name='CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.Add_TField_(string,System.Linq.Expressions.Expression_System.Func_T,TField__,System.Action_CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.GeneratorProvider_T,CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_T,TField___,System.Nullable_CloudyWing.SpreadsheetExporter.CellStyle_,System.Func_CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_T,TField_,CloudyWing.SpreadsheetExporter.CellStyle_).headerStyle'></a>

`headerStyle` [System\.Nullable&lt;](https://learn.microsoft.com/en-us/dotnet/api/system.nullable-1 'System\.Nullable\`1')[CellStyle](CloudyWing.SpreadsheetExporter.CellStyle.md 'CloudyWing\.SpreadsheetExporter\.CellStyle')[&gt;](https://learn.microsoft.com/en-us/dotnet/api/system.nullable-1 'System\.Nullable\`1')

The header style\. The dafault is `SpreadsheetManager.DefaultCellStyles.HeaderStyle`\.

<a name='CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.Add_TField_(string,System.Linq.Expressions.Expression_System.Func_T,TField__,System.Action_CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.GeneratorProvider_T,CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_T,TField___,System.Nullable_CloudyWing.SpreadsheetExporter.CellStyle_,System.Func_CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_T,TField_,CloudyWing.SpreadsheetExporter.CellStyle_).fieldStyleGenerator'></a>

`fieldStyleGenerator` [System\.Func&lt;](https://learn.microsoft.com/en-us/dotnet/api/system.func-2 'System\.Func\`2')[CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.FieldContext&lt;](CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_TRecord,TField_.md 'CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.FieldContext\<TRecord,TField\>')[T](CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.md#CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.T 'CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.DataColumnCollection\<T\>\.T')[,](CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_TRecord,TField_.md 'CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.FieldContext\<TRecord,TField\>')[TField](CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.md#CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.Add_TField_(string,System.Linq.Expressions.Expression_System.Func_T,TField__,System.Action_CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.GeneratorProvider_T,CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_T,TField___,System.Nullable_CloudyWing.SpreadsheetExporter.CellStyle_,System.Func_CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_T,TField_,CloudyWing.SpreadsheetExporter.CellStyle_).TField 'CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.DataColumnCollection\<T\>\.Add\<TField\>\(string, System\.Linq\.Expressions\.Expression\<System\.Func\<T,TField\>\>, System\.Action\<CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.DataColumnCollection\<T\>\.GeneratorProvider\<T,CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.FieldContext\<T,TField\>\>\>, System\.Nullable\<CloudyWing\.SpreadsheetExporter\.CellStyle\>, System\.Func\<CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.FieldContext\<T,TField\>,CloudyWing\.SpreadsheetExporter\.CellStyle\>\)\.TField')[&gt;](CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_TRecord,TField_.md 'CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.FieldContext\<TRecord,TField\>')[,](https://learn.microsoft.com/en-us/dotnet/api/system.func-2 'System\.Func\`2')[CellStyle](CloudyWing.SpreadsheetExporter.CellStyle.md 'CloudyWing\.SpreadsheetExporter\.CellStyle')[&gt;](https://learn.microsoft.com/en-us/dotnet/api/system.func-2 'System\.Func\`2')

The field style generator\. The dafault is `(context) => SpreadsheetManager.DefaultCellStyles.FieldStyle`\.

#### Returns
[CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.DataColumnCollection&lt;](CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.md 'CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.DataColumnCollection\<T\>')[T](CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.md#CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.T 'CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.DataColumnCollection\<T\>\.T')[&gt;](CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.md 'CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.DataColumnCollection\<T\>')  
The self\.

<a name='CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.Add_TField_(string,System.Linq.Expressions.Expression_System.Func_T,TField__,System.Nullable_CloudyWing.SpreadsheetExporter.CellStyle_,System.Func_CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_T,TField_,CloudyWing.SpreadsheetExporter.CellStyle_)'></a>

## DataColumnCollection\<T\>\.Add\<TField\>\(string, Expression\<Func\<T,TField\>\>, Nullable\<CellStyle\>, Func\<FieldContext\<T,TField\>,CellStyle\>\) Method

Adds the data column to the end of the DataColumnCollection\<T\>\.

```csharp
public CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection<T> Add<TField>(string headerText, System.Linq.Expressions.Expression<System.Func<T,TField>> fieldKeyExpression, System.Nullable<CloudyWing.SpreadsheetExporter.CellStyle> headerStyle=null, System.Func<CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext<T,TField>,CloudyWing.SpreadsheetExporter.CellStyle> fieldStyleGenerator=null);
```
#### Type parameters

<a name='CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.Add_TField_(string,System.Linq.Expressions.Expression_System.Func_T,TField__,System.Nullable_CloudyWing.SpreadsheetExporter.CellStyle_,System.Func_CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_T,TField_,CloudyWing.SpreadsheetExporter.CellStyle_).TField'></a>

`TField`

The type of the field\.
#### Parameters

<a name='CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.Add_TField_(string,System.Linq.Expressions.Expression_System.Func_T,TField__,System.Nullable_CloudyWing.SpreadsheetExporter.CellStyle_,System.Func_CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_T,TField_,CloudyWing.SpreadsheetExporter.CellStyle_).headerText'></a>

`headerText` [System\.String](https://learn.microsoft.com/en-us/dotnet/api/system.string 'System\.String')

The header text\.

<a name='CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.Add_TField_(string,System.Linq.Expressions.Expression_System.Func_T,TField__,System.Nullable_CloudyWing.SpreadsheetExporter.CellStyle_,System.Func_CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_T,TField_,CloudyWing.SpreadsheetExporter.CellStyle_).fieldKeyExpression'></a>

`fieldKeyExpression` [System\.Linq\.Expressions\.Expression&lt;](https://learn.microsoft.com/en-us/dotnet/api/system.linq.expressions.expression-1 'System\.Linq\.Expressions\.Expression\`1')[System\.Func&lt;](https://learn.microsoft.com/en-us/dotnet/api/system.func-2 'System\.Func\`2')[T](CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.md#CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.T 'CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.DataColumnCollection\<T\>\.T')[,](https://learn.microsoft.com/en-us/dotnet/api/system.func-2 'System\.Func\`2')[TField](CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.md#CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.Add_TField_(string,System.Linq.Expressions.Expression_System.Func_T,TField__,System.Nullable_CloudyWing.SpreadsheetExporter.CellStyle_,System.Func_CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_T,TField_,CloudyWing.SpreadsheetExporter.CellStyle_).TField 'CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.DataColumnCollection\<T\>\.Add\<TField\>\(string, System\.Linq\.Expressions\.Expression\<System\.Func\<T,TField\>\>, System\.Nullable\<CloudyWing\.SpreadsheetExporter\.CellStyle\>, System\.Func\<CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.FieldContext\<T,TField\>,CloudyWing\.SpreadsheetExporter\.CellStyle\>\)\.TField')[&gt;](https://learn.microsoft.com/en-us/dotnet/api/system.func-2 'System\.Func\`2')[&gt;](https://learn.microsoft.com/en-us/dotnet/api/system.linq.expressions.expression-1 'System\.Linq\.Expressions\.Expression\`1')

The field key expression\.

<a name='CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.Add_TField_(string,System.Linq.Expressions.Expression_System.Func_T,TField__,System.Nullable_CloudyWing.SpreadsheetExporter.CellStyle_,System.Func_CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_T,TField_,CloudyWing.SpreadsheetExporter.CellStyle_).headerStyle'></a>

`headerStyle` [System\.Nullable&lt;](https://learn.microsoft.com/en-us/dotnet/api/system.nullable-1 'System\.Nullable\`1')[CellStyle](CloudyWing.SpreadsheetExporter.CellStyle.md 'CloudyWing\.SpreadsheetExporter\.CellStyle')[&gt;](https://learn.microsoft.com/en-us/dotnet/api/system.nullable-1 'System\.Nullable\`1')

The header style\. The dafault is `SpreadsheetManager.DefaultCellStyles.HeaderStyle`\.

<a name='CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.Add_TField_(string,System.Linq.Expressions.Expression_System.Func_T,TField__,System.Nullable_CloudyWing.SpreadsheetExporter.CellStyle_,System.Func_CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_T,TField_,CloudyWing.SpreadsheetExporter.CellStyle_).fieldStyleGenerator'></a>

`fieldStyleGenerator` [System\.Func&lt;](https://learn.microsoft.com/en-us/dotnet/api/system.func-2 'System\.Func\`2')[CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.FieldContext&lt;](CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_TRecord,TField_.md 'CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.FieldContext\<TRecord,TField\>')[T](CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.md#CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.T 'CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.DataColumnCollection\<T\>\.T')[,](CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_TRecord,TField_.md 'CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.FieldContext\<TRecord,TField\>')[TField](CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.md#CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.Add_TField_(string,System.Linq.Expressions.Expression_System.Func_T,TField__,System.Nullable_CloudyWing.SpreadsheetExporter.CellStyle_,System.Func_CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_T,TField_,CloudyWing.SpreadsheetExporter.CellStyle_).TField 'CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.DataColumnCollection\<T\>\.Add\<TField\>\(string, System\.Linq\.Expressions\.Expression\<System\.Func\<T,TField\>\>, System\.Nullable\<CloudyWing\.SpreadsheetExporter\.CellStyle\>, System\.Func\<CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.FieldContext\<T,TField\>,CloudyWing\.SpreadsheetExporter\.CellStyle\>\)\.TField')[&gt;](CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_TRecord,TField_.md 'CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.FieldContext\<TRecord,TField\>')[,](https://learn.microsoft.com/en-us/dotnet/api/system.func-2 'System\.Func\`2')[CellStyle](CloudyWing.SpreadsheetExporter.CellStyle.md 'CloudyWing\.SpreadsheetExporter\.CellStyle')[&gt;](https://learn.microsoft.com/en-us/dotnet/api/system.func-2 'System\.Func\`2')

The field style generator\. The dafault is `(context) => SpreadsheetManager.DefaultCellStyles.FieldStyle`\.

#### Returns
[CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.DataColumnCollection&lt;](CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.md 'CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.DataColumnCollection\<T\>')[T](CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.md#CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.T 'CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.DataColumnCollection\<T\>\.T')[&gt;](CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.md 'CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.DataColumnCollection\<T\>')  
The self\.

<a name='CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.AddChildToLast(CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnBase_T_)'></a>

## DataColumnCollection\<T\>\.AddChildToLast\(DataColumnBase\<T\>\) Method

Adds the child data column at the end of the last data column\.

```csharp
public CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection<T> AddChildToLast(CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnBase<T> childColumn);
```
#### Parameters

<a name='CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.AddChildToLast(CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnBase_T_).childColumn'></a>

`childColumn` [CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.DataColumnBase&lt;](CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnBase_T_.md 'CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.DataColumnBase\<T\>')[T](CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.md#CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.T 'CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.DataColumnCollection\<T\>\.T')[&gt;](CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnBase_T_.md 'CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.DataColumnBase\<T\>')

The child data column\.

#### Returns
[CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.DataColumnCollection&lt;](CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.md 'CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.DataColumnCollection\<T\>')[T](CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.md#CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.T 'CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.DataColumnCollection\<T\>\.T')[&gt;](CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.md 'CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.DataColumnCollection\<T\>')  
The self\.

#### Exceptions

[System\.InvalidOperationException](https://learn.microsoft.com/en-us/dotnet/api/system.invalidoperationexception 'System\.InvalidOperationException')  
No columns available to add child to\.

<a name='CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.AddChildToLast(string,System.Action_CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.GeneratorProvider_T,CloudyWing.SpreadsheetExporter.Templates.RecordSet.RecordContext_T___,System.Nullable_CloudyWing.SpreadsheetExporter.CellStyle_,System.Func_CloudyWing.SpreadsheetExporter.Templates.RecordSet.RecordContext_T_,CloudyWing.SpreadsheetExporter.CellStyle_)'></a>

## DataColumnCollection\<T\>\.AddChildToLast\(string, Action\<GeneratorProvider\<T,RecordContext\<T\>\>\>, Nullable\<CellStyle\>, Func\<RecordContext\<T\>,CellStyle\>\) Method

Adds the child data column at the end of the last data column\.

```csharp
public CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection<T> AddChildToLast(string headerText, System.Action<CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection<T>.GeneratorProvider<T,CloudyWing.SpreadsheetExporter.Templates.RecordSet.RecordContext<T>>> providerSetter, System.Nullable<CloudyWing.SpreadsheetExporter.CellStyle> headerStyle=null, System.Func<CloudyWing.SpreadsheetExporter.Templates.RecordSet.RecordContext<T>,CloudyWing.SpreadsheetExporter.CellStyle> fieldStyleGenerator=null);
```
#### Parameters

<a name='CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.AddChildToLast(string,System.Action_CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.GeneratorProvider_T,CloudyWing.SpreadsheetExporter.Templates.RecordSet.RecordContext_T___,System.Nullable_CloudyWing.SpreadsheetExporter.CellStyle_,System.Func_CloudyWing.SpreadsheetExporter.Templates.RecordSet.RecordContext_T_,CloudyWing.SpreadsheetExporter.CellStyle_).headerText'></a>

`headerText` [System\.String](https://learn.microsoft.com/en-us/dotnet/api/system.string 'System\.String')

The header text\.

<a name='CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.AddChildToLast(string,System.Action_CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.GeneratorProvider_T,CloudyWing.SpreadsheetExporter.Templates.RecordSet.RecordContext_T___,System.Nullable_CloudyWing.SpreadsheetExporter.CellStyle_,System.Func_CloudyWing.SpreadsheetExporter.Templates.RecordSet.RecordContext_T_,CloudyWing.SpreadsheetExporter.CellStyle_).providerSetter'></a>

`providerSetter` [System\.Action&lt;](https://learn.microsoft.com/en-us/dotnet/api/system.action-1 'System\.Action\`1')[CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.DataColumnCollection\.GeneratorProvider&lt;](CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.GeneratorProvider_TRecord,TContext_.md 'CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.DataColumnCollection\<T\>\.GeneratorProvider\<TRecord,TContext\>')[T](CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.md#CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.T 'CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.DataColumnCollection\<T\>\.T')[,](CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.GeneratorProvider_TRecord,TContext_.md 'CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.DataColumnCollection\<T\>\.GeneratorProvider\<TRecord,TContext\>')[T](CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.md#CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.T 'CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.DataColumnCollection\<T\>\.T')[,](CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.GeneratorProvider_TRecord,TContext_.md 'CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.DataColumnCollection\<T\>\.GeneratorProvider\<TRecord,TContext\>')[CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.RecordContext&lt;](CloudyWing.SpreadsheetExporter.Templates.RecordSet.RecordContext_T_.md 'CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.RecordContext\<T\>')[T](CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.md#CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.T 'CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.DataColumnCollection\<T\>\.T')[&gt;](CloudyWing.SpreadsheetExporter.Templates.RecordSet.RecordContext_T_.md 'CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.RecordContext\<T\>')[&gt;](CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.GeneratorProvider_TRecord,TContext_.md 'CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.DataColumnCollection\<T\>\.GeneratorProvider\<TRecord,TContext\>')[&gt;](https://learn.microsoft.com/en-us/dotnet/api/system.action-1 'System\.Action\`1')

The provider setter\.

<a name='CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.AddChildToLast(string,System.Action_CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.GeneratorProvider_T,CloudyWing.SpreadsheetExporter.Templates.RecordSet.RecordContext_T___,System.Nullable_CloudyWing.SpreadsheetExporter.CellStyle_,System.Func_CloudyWing.SpreadsheetExporter.Templates.RecordSet.RecordContext_T_,CloudyWing.SpreadsheetExporter.CellStyle_).headerStyle'></a>

`headerStyle` [System\.Nullable&lt;](https://learn.microsoft.com/en-us/dotnet/api/system.nullable-1 'System\.Nullable\`1')[CellStyle](CloudyWing.SpreadsheetExporter.CellStyle.md 'CloudyWing\.SpreadsheetExporter\.CellStyle')[&gt;](https://learn.microsoft.com/en-us/dotnet/api/system.nullable-1 'System\.Nullable\`1')

The header style\. The dafault is `SpreadsheetManager.DefaultCellStyles.HeaderStyle`\.

<a name='CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.AddChildToLast(string,System.Action_CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.GeneratorProvider_T,CloudyWing.SpreadsheetExporter.Templates.RecordSet.RecordContext_T___,System.Nullable_CloudyWing.SpreadsheetExporter.CellStyle_,System.Func_CloudyWing.SpreadsheetExporter.Templates.RecordSet.RecordContext_T_,CloudyWing.SpreadsheetExporter.CellStyle_).fieldStyleGenerator'></a>

`fieldStyleGenerator` [System\.Func&lt;](https://learn.microsoft.com/en-us/dotnet/api/system.func-2 'System\.Func\`2')[CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.RecordContext&lt;](CloudyWing.SpreadsheetExporter.Templates.RecordSet.RecordContext_T_.md 'CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.RecordContext\<T\>')[T](CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.md#CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.T 'CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.DataColumnCollection\<T\>\.T')[&gt;](CloudyWing.SpreadsheetExporter.Templates.RecordSet.RecordContext_T_.md 'CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.RecordContext\<T\>')[,](https://learn.microsoft.com/en-us/dotnet/api/system.func-2 'System\.Func\`2')[CellStyle](CloudyWing.SpreadsheetExporter.CellStyle.md 'CloudyWing\.SpreadsheetExporter\.CellStyle')[&gt;](https://learn.microsoft.com/en-us/dotnet/api/system.func-2 'System\.Func\`2')

The field style generator\. The dafault is `(context) => SpreadsheetManager.DefaultCellStyles.FieldStyle`\.

#### Returns
[CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.DataColumnCollection&lt;](CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.md 'CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.DataColumnCollection\<T\>')[T](CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.md#CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.T 'CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.DataColumnCollection\<T\>\.T')[&gt;](CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.md 'CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.DataColumnCollection\<T\>')  
The self\.

<a name='CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.AddChildToLast(string,System.Nullable_CloudyWing.SpreadsheetExporter.CellStyle_,System.Func_CloudyWing.SpreadsheetExporter.Templates.RecordSet.RecordContext_T_,CloudyWing.SpreadsheetExporter.CellStyle_)'></a>

## DataColumnCollection\<T\>\.AddChildToLast\(string, Nullable\<CellStyle\>, Func\<RecordContext\<T\>,CellStyle\>\) Method

Adds the child data column at the end of the last data column\.

```csharp
public CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection<T> AddChildToLast(string headerText, System.Nullable<CloudyWing.SpreadsheetExporter.CellStyle> headerStyle=null, System.Func<CloudyWing.SpreadsheetExporter.Templates.RecordSet.RecordContext<T>,CloudyWing.SpreadsheetExporter.CellStyle> fieldStyleGenerator=null);
```
#### Parameters

<a name='CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.AddChildToLast(string,System.Nullable_CloudyWing.SpreadsheetExporter.CellStyle_,System.Func_CloudyWing.SpreadsheetExporter.Templates.RecordSet.RecordContext_T_,CloudyWing.SpreadsheetExporter.CellStyle_).headerText'></a>

`headerText` [System\.String](https://learn.microsoft.com/en-us/dotnet/api/system.string 'System\.String')

The header text\.

<a name='CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.AddChildToLast(string,System.Nullable_CloudyWing.SpreadsheetExporter.CellStyle_,System.Func_CloudyWing.SpreadsheetExporter.Templates.RecordSet.RecordContext_T_,CloudyWing.SpreadsheetExporter.CellStyle_).headerStyle'></a>

`headerStyle` [System\.Nullable&lt;](https://learn.microsoft.com/en-us/dotnet/api/system.nullable-1 'System\.Nullable\`1')[CellStyle](CloudyWing.SpreadsheetExporter.CellStyle.md 'CloudyWing\.SpreadsheetExporter\.CellStyle')[&gt;](https://learn.microsoft.com/en-us/dotnet/api/system.nullable-1 'System\.Nullable\`1')

The header style\.

<a name='CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.AddChildToLast(string,System.Nullable_CloudyWing.SpreadsheetExporter.CellStyle_,System.Func_CloudyWing.SpreadsheetExporter.Templates.RecordSet.RecordContext_T_,CloudyWing.SpreadsheetExporter.CellStyle_).fieldStyleGenerator'></a>

`fieldStyleGenerator` [System\.Func&lt;](https://learn.microsoft.com/en-us/dotnet/api/system.func-2 'System\.Func\`2')[CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.RecordContext&lt;](CloudyWing.SpreadsheetExporter.Templates.RecordSet.RecordContext_T_.md 'CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.RecordContext\<T\>')[T](CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.md#CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.T 'CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.DataColumnCollection\<T\>\.T')[&gt;](CloudyWing.SpreadsheetExporter.Templates.RecordSet.RecordContext_T_.md 'CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.RecordContext\<T\>')[,](https://learn.microsoft.com/en-us/dotnet/api/system.func-2 'System\.Func\`2')[CellStyle](CloudyWing.SpreadsheetExporter.CellStyle.md 'CloudyWing\.SpreadsheetExporter\.CellStyle')[&gt;](https://learn.microsoft.com/en-us/dotnet/api/system.func-2 'System\.Func\`2')

The field style generator\.

#### Returns
[CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.DataColumnCollection&lt;](CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.md 'CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.DataColumnCollection\<T\>')[T](CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.md#CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.T 'CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.DataColumnCollection\<T\>\.T')[&gt;](CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.md 'CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.DataColumnCollection\<T\>')  
The self\.

<a name='CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.AddChildToLast_TField_(string,string,System.Action_CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.GeneratorProvider_T,CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_T,TField___,System.Nullable_CloudyWing.SpreadsheetExporter.CellStyle_,System.Func_CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_T,TField_,CloudyWing.SpreadsheetExporter.CellStyle_)'></a>

## DataColumnCollection\<T\>\.AddChildToLast\<TField\>\(string, string, Action\<GeneratorProvider\<T,FieldContext\<T,TField\>\>\>, Nullable\<CellStyle\>, Func\<FieldContext\<T,TField\>,CellStyle\>\) Method

Adds the child data column at the end of the last data column\.

```csharp
public CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection<T> AddChildToLast<TField>(string headerText, string fieldKey, System.Action<CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection<T>.GeneratorProvider<T,CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext<T,TField>>> providerSetter, System.Nullable<CloudyWing.SpreadsheetExporter.CellStyle> headerStyle=null, System.Func<CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext<T,TField>,CloudyWing.SpreadsheetExporter.CellStyle> fieldStyleGenerator=null);
```
#### Type parameters

<a name='CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.AddChildToLast_TField_(string,string,System.Action_CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.GeneratorProvider_T,CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_T,TField___,System.Nullable_CloudyWing.SpreadsheetExporter.CellStyle_,System.Func_CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_T,TField_,CloudyWing.SpreadsheetExporter.CellStyle_).TField'></a>

`TField`

The type of the field\.
#### Parameters

<a name='CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.AddChildToLast_TField_(string,string,System.Action_CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.GeneratorProvider_T,CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_T,TField___,System.Nullable_CloudyWing.SpreadsheetExporter.CellStyle_,System.Func_CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_T,TField_,CloudyWing.SpreadsheetExporter.CellStyle_).headerText'></a>

`headerText` [System\.String](https://learn.microsoft.com/en-us/dotnet/api/system.string 'System\.String')

The header text\.

<a name='CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.AddChildToLast_TField_(string,string,System.Action_CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.GeneratorProvider_T,CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_T,TField___,System.Nullable_CloudyWing.SpreadsheetExporter.CellStyle_,System.Func_CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_T,TField_,CloudyWing.SpreadsheetExporter.CellStyle_).fieldKey'></a>

`fieldKey` [System\.String](https://learn.microsoft.com/en-us/dotnet/api/system.string 'System\.String')

The field key\.

<a name='CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.AddChildToLast_TField_(string,string,System.Action_CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.GeneratorProvider_T,CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_T,TField___,System.Nullable_CloudyWing.SpreadsheetExporter.CellStyle_,System.Func_CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_T,TField_,CloudyWing.SpreadsheetExporter.CellStyle_).providerSetter'></a>

`providerSetter` [System\.Action&lt;](https://learn.microsoft.com/en-us/dotnet/api/system.action-1 'System\.Action\`1')[CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.DataColumnCollection\.GeneratorProvider&lt;](CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.GeneratorProvider_TRecord,TContext_.md 'CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.DataColumnCollection\<T\>\.GeneratorProvider\<TRecord,TContext\>')[T](CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.md#CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.T 'CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.DataColumnCollection\<T\>\.T')[,](CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.GeneratorProvider_TRecord,TContext_.md 'CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.DataColumnCollection\<T\>\.GeneratorProvider\<TRecord,TContext\>')[T](CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.md#CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.T 'CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.DataColumnCollection\<T\>\.T')[,](CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.GeneratorProvider_TRecord,TContext_.md 'CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.DataColumnCollection\<T\>\.GeneratorProvider\<TRecord,TContext\>')[CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.FieldContext&lt;](CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_TRecord,TField_.md 'CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.FieldContext\<TRecord,TField\>')[T](CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.md#CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.T 'CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.DataColumnCollection\<T\>\.T')[,](CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_TRecord,TField_.md 'CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.FieldContext\<TRecord,TField\>')[TField](CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.md#CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.AddChildToLast_TField_(string,string,System.Action_CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.GeneratorProvider_T,CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_T,TField___,System.Nullable_CloudyWing.SpreadsheetExporter.CellStyle_,System.Func_CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_T,TField_,CloudyWing.SpreadsheetExporter.CellStyle_).TField 'CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.DataColumnCollection\<T\>\.AddChildToLast\<TField\>\(string, string, System\.Action\<CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.DataColumnCollection\<T\>\.GeneratorProvider\<T,CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.FieldContext\<T,TField\>\>\>, System\.Nullable\<CloudyWing\.SpreadsheetExporter\.CellStyle\>, System\.Func\<CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.FieldContext\<T,TField\>,CloudyWing\.SpreadsheetExporter\.CellStyle\>\)\.TField')[&gt;](CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_TRecord,TField_.md 'CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.FieldContext\<TRecord,TField\>')[&gt;](CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.GeneratorProvider_TRecord,TContext_.md 'CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.DataColumnCollection\<T\>\.GeneratorProvider\<TRecord,TContext\>')[&gt;](https://learn.microsoft.com/en-us/dotnet/api/system.action-1 'System\.Action\`1')

The provider setter\.

<a name='CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.AddChildToLast_TField_(string,string,System.Action_CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.GeneratorProvider_T,CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_T,TField___,System.Nullable_CloudyWing.SpreadsheetExporter.CellStyle_,System.Func_CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_T,TField_,CloudyWing.SpreadsheetExporter.CellStyle_).headerStyle'></a>

`headerStyle` [System\.Nullable&lt;](https://learn.microsoft.com/en-us/dotnet/api/system.nullable-1 'System\.Nullable\`1')[CellStyle](CloudyWing.SpreadsheetExporter.CellStyle.md 'CloudyWing\.SpreadsheetExporter\.CellStyle')[&gt;](https://learn.microsoft.com/en-us/dotnet/api/system.nullable-1 'System\.Nullable\`1')

The header style\. The dafault is `SpreadsheetManager.DefaultCellStyles.HeaderStyle`\.

<a name='CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.AddChildToLast_TField_(string,string,System.Action_CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.GeneratorProvider_T,CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_T,TField___,System.Nullable_CloudyWing.SpreadsheetExporter.CellStyle_,System.Func_CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_T,TField_,CloudyWing.SpreadsheetExporter.CellStyle_).fieldStyleGenerator'></a>

`fieldStyleGenerator` [System\.Func&lt;](https://learn.microsoft.com/en-us/dotnet/api/system.func-2 'System\.Func\`2')[CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.FieldContext&lt;](CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_TRecord,TField_.md 'CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.FieldContext\<TRecord,TField\>')[T](CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.md#CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.T 'CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.DataColumnCollection\<T\>\.T')[,](CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_TRecord,TField_.md 'CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.FieldContext\<TRecord,TField\>')[TField](CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.md#CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.AddChildToLast_TField_(string,string,System.Action_CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.GeneratorProvider_T,CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_T,TField___,System.Nullable_CloudyWing.SpreadsheetExporter.CellStyle_,System.Func_CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_T,TField_,CloudyWing.SpreadsheetExporter.CellStyle_).TField 'CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.DataColumnCollection\<T\>\.AddChildToLast\<TField\>\(string, string, System\.Action\<CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.DataColumnCollection\<T\>\.GeneratorProvider\<T,CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.FieldContext\<T,TField\>\>\>, System\.Nullable\<CloudyWing\.SpreadsheetExporter\.CellStyle\>, System\.Func\<CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.FieldContext\<T,TField\>,CloudyWing\.SpreadsheetExporter\.CellStyle\>\)\.TField')[&gt;](CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_TRecord,TField_.md 'CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.FieldContext\<TRecord,TField\>')[,](https://learn.microsoft.com/en-us/dotnet/api/system.func-2 'System\.Func\`2')[CellStyle](CloudyWing.SpreadsheetExporter.CellStyle.md 'CloudyWing\.SpreadsheetExporter\.CellStyle')[&gt;](https://learn.microsoft.com/en-us/dotnet/api/system.func-2 'System\.Func\`2')

The field style generator\. The dafault is `(context) => SpreadsheetManager.DefaultCellStyles.FieldStyle`\.

#### Returns
[CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.DataColumnCollection&lt;](CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.md 'CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.DataColumnCollection\<T\>')[T](CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.md#CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.T 'CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.DataColumnCollection\<T\>\.T')[&gt;](CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.md 'CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.DataColumnCollection\<T\>')  
The self\.

<a name='CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.AddChildToLast_TField_(string,string,System.Nullable_CloudyWing.SpreadsheetExporter.CellStyle_,System.Func_CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_T,TField_,CloudyWing.SpreadsheetExporter.CellStyle_)'></a>

## DataColumnCollection\<T\>\.AddChildToLast\<TField\>\(string, string, Nullable\<CellStyle\>, Func\<FieldContext\<T,TField\>,CellStyle\>\) Method

Adds the child data column at the end of the last data column\.

```csharp
public CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection<T> AddChildToLast<TField>(string headerText, string fieldKey, System.Nullable<CloudyWing.SpreadsheetExporter.CellStyle> headerStyle=null, System.Func<CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext<T,TField>,CloudyWing.SpreadsheetExporter.CellStyle> fieldStyleGenerator=null);
```
#### Type parameters

<a name='CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.AddChildToLast_TField_(string,string,System.Nullable_CloudyWing.SpreadsheetExporter.CellStyle_,System.Func_CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_T,TField_,CloudyWing.SpreadsheetExporter.CellStyle_).TField'></a>

`TField`

The type of the field\.
#### Parameters

<a name='CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.AddChildToLast_TField_(string,string,System.Nullable_CloudyWing.SpreadsheetExporter.CellStyle_,System.Func_CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_T,TField_,CloudyWing.SpreadsheetExporter.CellStyle_).headerText'></a>

`headerText` [System\.String](https://learn.microsoft.com/en-us/dotnet/api/system.string 'System\.String')

The header text\.

<a name='CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.AddChildToLast_TField_(string,string,System.Nullable_CloudyWing.SpreadsheetExporter.CellStyle_,System.Func_CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_T,TField_,CloudyWing.SpreadsheetExporter.CellStyle_).fieldKey'></a>

`fieldKey` [System\.String](https://learn.microsoft.com/en-us/dotnet/api/system.string 'System\.String')

The field key\.

<a name='CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.AddChildToLast_TField_(string,string,System.Nullable_CloudyWing.SpreadsheetExporter.CellStyle_,System.Func_CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_T,TField_,CloudyWing.SpreadsheetExporter.CellStyle_).headerStyle'></a>

`headerStyle` [System\.Nullable&lt;](https://learn.microsoft.com/en-us/dotnet/api/system.nullable-1 'System\.Nullable\`1')[CellStyle](CloudyWing.SpreadsheetExporter.CellStyle.md 'CloudyWing\.SpreadsheetExporter\.CellStyle')[&gt;](https://learn.microsoft.com/en-us/dotnet/api/system.nullable-1 'System\.Nullable\`1')

The header style\. The dafault is `SpreadsheetManager.DefaultCellStyles.HeaderStyle`\.

<a name='CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.AddChildToLast_TField_(string,string,System.Nullable_CloudyWing.SpreadsheetExporter.CellStyle_,System.Func_CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_T,TField_,CloudyWing.SpreadsheetExporter.CellStyle_).fieldStyleGenerator'></a>

`fieldStyleGenerator` [System\.Func&lt;](https://learn.microsoft.com/en-us/dotnet/api/system.func-2 'System\.Func\`2')[CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.FieldContext&lt;](CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_TRecord,TField_.md 'CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.FieldContext\<TRecord,TField\>')[T](CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.md#CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.T 'CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.DataColumnCollection\<T\>\.T')[,](CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_TRecord,TField_.md 'CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.FieldContext\<TRecord,TField\>')[TField](CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.md#CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.AddChildToLast_TField_(string,string,System.Nullable_CloudyWing.SpreadsheetExporter.CellStyle_,System.Func_CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_T,TField_,CloudyWing.SpreadsheetExporter.CellStyle_).TField 'CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.DataColumnCollection\<T\>\.AddChildToLast\<TField\>\(string, string, System\.Nullable\<CloudyWing\.SpreadsheetExporter\.CellStyle\>, System\.Func\<CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.FieldContext\<T,TField\>,CloudyWing\.SpreadsheetExporter\.CellStyle\>\)\.TField')[&gt;](CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_TRecord,TField_.md 'CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.FieldContext\<TRecord,TField\>')[,](https://learn.microsoft.com/en-us/dotnet/api/system.func-2 'System\.Func\`2')[CellStyle](CloudyWing.SpreadsheetExporter.CellStyle.md 'CloudyWing\.SpreadsheetExporter\.CellStyle')[&gt;](https://learn.microsoft.com/en-us/dotnet/api/system.func-2 'System\.Func\`2')

The field style generator\. The dafault is `(context) => SpreadsheetManager.DefaultCellStyles.FieldStyle`\.

#### Returns
[CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.DataColumnCollection&lt;](CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.md 'CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.DataColumnCollection\<T\>')[T](CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.md#CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.T 'CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.DataColumnCollection\<T\>\.T')[&gt;](CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.md 'CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.DataColumnCollection\<T\>')  
The self\.

<a name='CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.AddChildToLast_TField_(string,System.Linq.Expressions.Expression_System.Func_T,TField__,System.Action_CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.GeneratorProvider_T,CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_T,TField___,System.Nullable_CloudyWing.SpreadsheetExporter.CellStyle_,System.Func_CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_T,TField_,CloudyWing.SpreadsheetExporter.CellStyle_)'></a>

## DataColumnCollection\<T\>\.AddChildToLast\<TField\>\(string, Expression\<Func\<T,TField\>\>, Action\<GeneratorProvider\<T,FieldContext\<T,TField\>\>\>, Nullable\<CellStyle\>, Func\<FieldContext\<T,TField\>,CellStyle\>\) Method

Adds the child data column at the end of the last data column\.

```csharp
public CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection<T> AddChildToLast<TField>(string headerText, System.Linq.Expressions.Expression<System.Func<T,TField>> fieldKeyExpression, System.Action<CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection<T>.GeneratorProvider<T,CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext<T,TField>>> providerSetter, System.Nullable<CloudyWing.SpreadsheetExporter.CellStyle> headerStyle=null, System.Func<CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext<T,TField>,CloudyWing.SpreadsheetExporter.CellStyle> fieldStyleGenerator=null);
```
#### Type parameters

<a name='CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.AddChildToLast_TField_(string,System.Linq.Expressions.Expression_System.Func_T,TField__,System.Action_CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.GeneratorProvider_T,CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_T,TField___,System.Nullable_CloudyWing.SpreadsheetExporter.CellStyle_,System.Func_CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_T,TField_,CloudyWing.SpreadsheetExporter.CellStyle_).TField'></a>

`TField`

The type of the field\.
#### Parameters

<a name='CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.AddChildToLast_TField_(string,System.Linq.Expressions.Expression_System.Func_T,TField__,System.Action_CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.GeneratorProvider_T,CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_T,TField___,System.Nullable_CloudyWing.SpreadsheetExporter.CellStyle_,System.Func_CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_T,TField_,CloudyWing.SpreadsheetExporter.CellStyle_).headerText'></a>

`headerText` [System\.String](https://learn.microsoft.com/en-us/dotnet/api/system.string 'System\.String')

The header text\.

<a name='CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.AddChildToLast_TField_(string,System.Linq.Expressions.Expression_System.Func_T,TField__,System.Action_CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.GeneratorProvider_T,CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_T,TField___,System.Nullable_CloudyWing.SpreadsheetExporter.CellStyle_,System.Func_CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_T,TField_,CloudyWing.SpreadsheetExporter.CellStyle_).fieldKeyExpression'></a>

`fieldKeyExpression` [System\.Linq\.Expressions\.Expression&lt;](https://learn.microsoft.com/en-us/dotnet/api/system.linq.expressions.expression-1 'System\.Linq\.Expressions\.Expression\`1')[System\.Func&lt;](https://learn.microsoft.com/en-us/dotnet/api/system.func-2 'System\.Func\`2')[T](CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.md#CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.T 'CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.DataColumnCollection\<T\>\.T')[,](https://learn.microsoft.com/en-us/dotnet/api/system.func-2 'System\.Func\`2')[TField](CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.md#CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.AddChildToLast_TField_(string,System.Linq.Expressions.Expression_System.Func_T,TField__,System.Action_CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.GeneratorProvider_T,CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_T,TField___,System.Nullable_CloudyWing.SpreadsheetExporter.CellStyle_,System.Func_CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_T,TField_,CloudyWing.SpreadsheetExporter.CellStyle_).TField 'CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.DataColumnCollection\<T\>\.AddChildToLast\<TField\>\(string, System\.Linq\.Expressions\.Expression\<System\.Func\<T,TField\>\>, System\.Action\<CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.DataColumnCollection\<T\>\.GeneratorProvider\<T,CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.FieldContext\<T,TField\>\>\>, System\.Nullable\<CloudyWing\.SpreadsheetExporter\.CellStyle\>, System\.Func\<CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.FieldContext\<T,TField\>,CloudyWing\.SpreadsheetExporter\.CellStyle\>\)\.TField')[&gt;](https://learn.microsoft.com/en-us/dotnet/api/system.func-2 'System\.Func\`2')[&gt;](https://learn.microsoft.com/en-us/dotnet/api/system.linq.expressions.expression-1 'System\.Linq\.Expressions\.Expression\`1')

The field key expression\.

<a name='CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.AddChildToLast_TField_(string,System.Linq.Expressions.Expression_System.Func_T,TField__,System.Action_CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.GeneratorProvider_T,CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_T,TField___,System.Nullable_CloudyWing.SpreadsheetExporter.CellStyle_,System.Func_CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_T,TField_,CloudyWing.SpreadsheetExporter.CellStyle_).providerSetter'></a>

`providerSetter` [System\.Action&lt;](https://learn.microsoft.com/en-us/dotnet/api/system.action-1 'System\.Action\`1')[CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.DataColumnCollection\.GeneratorProvider&lt;](CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.GeneratorProvider_TRecord,TContext_.md 'CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.DataColumnCollection\<T\>\.GeneratorProvider\<TRecord,TContext\>')[T](CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.md#CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.T 'CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.DataColumnCollection\<T\>\.T')[,](CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.GeneratorProvider_TRecord,TContext_.md 'CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.DataColumnCollection\<T\>\.GeneratorProvider\<TRecord,TContext\>')[T](CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.md#CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.T 'CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.DataColumnCollection\<T\>\.T')[,](CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.GeneratorProvider_TRecord,TContext_.md 'CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.DataColumnCollection\<T\>\.GeneratorProvider\<TRecord,TContext\>')[CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.FieldContext&lt;](CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_TRecord,TField_.md 'CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.FieldContext\<TRecord,TField\>')[T](CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.md#CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.T 'CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.DataColumnCollection\<T\>\.T')[,](CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_TRecord,TField_.md 'CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.FieldContext\<TRecord,TField\>')[TField](CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.md#CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.AddChildToLast_TField_(string,System.Linq.Expressions.Expression_System.Func_T,TField__,System.Action_CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.GeneratorProvider_T,CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_T,TField___,System.Nullable_CloudyWing.SpreadsheetExporter.CellStyle_,System.Func_CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_T,TField_,CloudyWing.SpreadsheetExporter.CellStyle_).TField 'CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.DataColumnCollection\<T\>\.AddChildToLast\<TField\>\(string, System\.Linq\.Expressions\.Expression\<System\.Func\<T,TField\>\>, System\.Action\<CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.DataColumnCollection\<T\>\.GeneratorProvider\<T,CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.FieldContext\<T,TField\>\>\>, System\.Nullable\<CloudyWing\.SpreadsheetExporter\.CellStyle\>, System\.Func\<CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.FieldContext\<T,TField\>,CloudyWing\.SpreadsheetExporter\.CellStyle\>\)\.TField')[&gt;](CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_TRecord,TField_.md 'CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.FieldContext\<TRecord,TField\>')[&gt;](CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.GeneratorProvider_TRecord,TContext_.md 'CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.DataColumnCollection\<T\>\.GeneratorProvider\<TRecord,TContext\>')[&gt;](https://learn.microsoft.com/en-us/dotnet/api/system.action-1 'System\.Action\`1')

The provider setter\.

<a name='CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.AddChildToLast_TField_(string,System.Linq.Expressions.Expression_System.Func_T,TField__,System.Action_CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.GeneratorProvider_T,CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_T,TField___,System.Nullable_CloudyWing.SpreadsheetExporter.CellStyle_,System.Func_CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_T,TField_,CloudyWing.SpreadsheetExporter.CellStyle_).headerStyle'></a>

`headerStyle` [System\.Nullable&lt;](https://learn.microsoft.com/en-us/dotnet/api/system.nullable-1 'System\.Nullable\`1')[CellStyle](CloudyWing.SpreadsheetExporter.CellStyle.md 'CloudyWing\.SpreadsheetExporter\.CellStyle')[&gt;](https://learn.microsoft.com/en-us/dotnet/api/system.nullable-1 'System\.Nullable\`1')

The header style\. The dafault is `SpreadsheetManager.DefaultCellStyles.HeaderStyle`\.

<a name='CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.AddChildToLast_TField_(string,System.Linq.Expressions.Expression_System.Func_T,TField__,System.Action_CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.GeneratorProvider_T,CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_T,TField___,System.Nullable_CloudyWing.SpreadsheetExporter.CellStyle_,System.Func_CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_T,TField_,CloudyWing.SpreadsheetExporter.CellStyle_).fieldStyleGenerator'></a>

`fieldStyleGenerator` [System\.Func&lt;](https://learn.microsoft.com/en-us/dotnet/api/system.func-2 'System\.Func\`2')[CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.FieldContext&lt;](CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_TRecord,TField_.md 'CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.FieldContext\<TRecord,TField\>')[T](CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.md#CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.T 'CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.DataColumnCollection\<T\>\.T')[,](CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_TRecord,TField_.md 'CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.FieldContext\<TRecord,TField\>')[TField](CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.md#CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.AddChildToLast_TField_(string,System.Linq.Expressions.Expression_System.Func_T,TField__,System.Action_CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.GeneratorProvider_T,CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_T,TField___,System.Nullable_CloudyWing.SpreadsheetExporter.CellStyle_,System.Func_CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_T,TField_,CloudyWing.SpreadsheetExporter.CellStyle_).TField 'CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.DataColumnCollection\<T\>\.AddChildToLast\<TField\>\(string, System\.Linq\.Expressions\.Expression\<System\.Func\<T,TField\>\>, System\.Action\<CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.DataColumnCollection\<T\>\.GeneratorProvider\<T,CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.FieldContext\<T,TField\>\>\>, System\.Nullable\<CloudyWing\.SpreadsheetExporter\.CellStyle\>, System\.Func\<CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.FieldContext\<T,TField\>,CloudyWing\.SpreadsheetExporter\.CellStyle\>\)\.TField')[&gt;](CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_TRecord,TField_.md 'CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.FieldContext\<TRecord,TField\>')[,](https://learn.microsoft.com/en-us/dotnet/api/system.func-2 'System\.Func\`2')[CellStyle](CloudyWing.SpreadsheetExporter.CellStyle.md 'CloudyWing\.SpreadsheetExporter\.CellStyle')[&gt;](https://learn.microsoft.com/en-us/dotnet/api/system.func-2 'System\.Func\`2')

The field style generator\. The dafault is `(context) => SpreadsheetManager.DefaultCellStyles.FieldStyle`\.

#### Returns
[CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.DataColumnCollection&lt;](CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.md 'CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.DataColumnCollection\<T\>')[T](CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.md#CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.T 'CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.DataColumnCollection\<T\>\.T')[&gt;](CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.md 'CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.DataColumnCollection\<T\>')  
The self\.

<a name='CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.AddChildToLast_TField_(string,System.Linq.Expressions.Expression_System.Func_T,TField__,System.Nullable_CloudyWing.SpreadsheetExporter.CellStyle_,System.Func_CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_T,TField_,CloudyWing.SpreadsheetExporter.CellStyle_)'></a>

## DataColumnCollection\<T\>\.AddChildToLast\<TField\>\(string, Expression\<Func\<T,TField\>\>, Nullable\<CellStyle\>, Func\<FieldContext\<T,TField\>,CellStyle\>\) Method

Adds the child data column at the end of the last data column\.

```csharp
public CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection<T> AddChildToLast<TField>(string headerText, System.Linq.Expressions.Expression<System.Func<T,TField>> fieldKeyExpression, System.Nullable<CloudyWing.SpreadsheetExporter.CellStyle> headerStyle=null, System.Func<CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext<T,TField>,CloudyWing.SpreadsheetExporter.CellStyle> fieldStyleGenerator=null);
```
#### Type parameters

<a name='CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.AddChildToLast_TField_(string,System.Linq.Expressions.Expression_System.Func_T,TField__,System.Nullable_CloudyWing.SpreadsheetExporter.CellStyle_,System.Func_CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_T,TField_,CloudyWing.SpreadsheetExporter.CellStyle_).TField'></a>

`TField`

The type of the field\.
#### Parameters

<a name='CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.AddChildToLast_TField_(string,System.Linq.Expressions.Expression_System.Func_T,TField__,System.Nullable_CloudyWing.SpreadsheetExporter.CellStyle_,System.Func_CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_T,TField_,CloudyWing.SpreadsheetExporter.CellStyle_).headerText'></a>

`headerText` [System\.String](https://learn.microsoft.com/en-us/dotnet/api/system.string 'System\.String')

The header text\.

<a name='CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.AddChildToLast_TField_(string,System.Linq.Expressions.Expression_System.Func_T,TField__,System.Nullable_CloudyWing.SpreadsheetExporter.CellStyle_,System.Func_CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_T,TField_,CloudyWing.SpreadsheetExporter.CellStyle_).fieldKeyExpression'></a>

`fieldKeyExpression` [System\.Linq\.Expressions\.Expression&lt;](https://learn.microsoft.com/en-us/dotnet/api/system.linq.expressions.expression-1 'System\.Linq\.Expressions\.Expression\`1')[System\.Func&lt;](https://learn.microsoft.com/en-us/dotnet/api/system.func-2 'System\.Func\`2')[T](CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.md#CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.T 'CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.DataColumnCollection\<T\>\.T')[,](https://learn.microsoft.com/en-us/dotnet/api/system.func-2 'System\.Func\`2')[TField](CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.md#CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.AddChildToLast_TField_(string,System.Linq.Expressions.Expression_System.Func_T,TField__,System.Nullable_CloudyWing.SpreadsheetExporter.CellStyle_,System.Func_CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_T,TField_,CloudyWing.SpreadsheetExporter.CellStyle_).TField 'CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.DataColumnCollection\<T\>\.AddChildToLast\<TField\>\(string, System\.Linq\.Expressions\.Expression\<System\.Func\<T,TField\>\>, System\.Nullable\<CloudyWing\.SpreadsheetExporter\.CellStyle\>, System\.Func\<CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.FieldContext\<T,TField\>,CloudyWing\.SpreadsheetExporter\.CellStyle\>\)\.TField')[&gt;](https://learn.microsoft.com/en-us/dotnet/api/system.func-2 'System\.Func\`2')[&gt;](https://learn.microsoft.com/en-us/dotnet/api/system.linq.expressions.expression-1 'System\.Linq\.Expressions\.Expression\`1')

The field key expression\.

<a name='CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.AddChildToLast_TField_(string,System.Linq.Expressions.Expression_System.Func_T,TField__,System.Nullable_CloudyWing.SpreadsheetExporter.CellStyle_,System.Func_CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_T,TField_,CloudyWing.SpreadsheetExporter.CellStyle_).headerStyle'></a>

`headerStyle` [System\.Nullable&lt;](https://learn.microsoft.com/en-us/dotnet/api/system.nullable-1 'System\.Nullable\`1')[CellStyle](CloudyWing.SpreadsheetExporter.CellStyle.md 'CloudyWing\.SpreadsheetExporter\.CellStyle')[&gt;](https://learn.microsoft.com/en-us/dotnet/api/system.nullable-1 'System\.Nullable\`1')

The header style\. The dafault is `SpreadsheetManager.DefaultCellStyles.HeaderStyle`\.

<a name='CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.AddChildToLast_TField_(string,System.Linq.Expressions.Expression_System.Func_T,TField__,System.Nullable_CloudyWing.SpreadsheetExporter.CellStyle_,System.Func_CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_T,TField_,CloudyWing.SpreadsheetExporter.CellStyle_).fieldStyleGenerator'></a>

`fieldStyleGenerator` [System\.Func&lt;](https://learn.microsoft.com/en-us/dotnet/api/system.func-2 'System\.Func\`2')[CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.FieldContext&lt;](CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_TRecord,TField_.md 'CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.FieldContext\<TRecord,TField\>')[T](CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.md#CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.T 'CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.DataColumnCollection\<T\>\.T')[,](CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_TRecord,TField_.md 'CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.FieldContext\<TRecord,TField\>')[TField](CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.md#CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.AddChildToLast_TField_(string,System.Linq.Expressions.Expression_System.Func_T,TField__,System.Nullable_CloudyWing.SpreadsheetExporter.CellStyle_,System.Func_CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_T,TField_,CloudyWing.SpreadsheetExporter.CellStyle_).TField 'CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.DataColumnCollection\<T\>\.AddChildToLast\<TField\>\(string, System\.Linq\.Expressions\.Expression\<System\.Func\<T,TField\>\>, System\.Nullable\<CloudyWing\.SpreadsheetExporter\.CellStyle\>, System\.Func\<CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.FieldContext\<T,TField\>,CloudyWing\.SpreadsheetExporter\.CellStyle\>\)\.TField')[&gt;](CloudyWing.SpreadsheetExporter.Templates.RecordSet.FieldContext_TRecord,TField_.md 'CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.FieldContext\<TRecord,TField\>')[,](https://learn.microsoft.com/en-us/dotnet/api/system.func-2 'System\.Func\`2')[CellStyle](CloudyWing.SpreadsheetExporter.CellStyle.md 'CloudyWing\.SpreadsheetExporter\.CellStyle')[&gt;](https://learn.microsoft.com/en-us/dotnet/api/system.func-2 'System\.Func\`2')

The field style generator\. The dafault is `(context) => SpreadsheetManager.DefaultCellStyles.FieldStyle`\.

#### Returns
[CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.DataColumnCollection&lt;](CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.md 'CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.DataColumnCollection\<T\>')[T](CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.md#CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.T 'CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.DataColumnCollection\<T\>\.T')[&gt;](CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.md 'CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.DataColumnCollection\<T\>')  
The self\.

<a name='CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.CalculatePoints(System.Drawing.Point)'></a>

## DataColumnCollection\<T\>\.CalculatePoints\(Point\) Method

Calculates the points\.

```csharp
internal void CalculatePoints(System.Drawing.Point point);
```
#### Parameters

<a name='CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.CalculatePoints(System.Drawing.Point).point'></a>

`point` [System\.Drawing\.Point](https://learn.microsoft.com/en-us/dotnet/api/system.drawing.point 'System\.Drawing\.Point')

<a name='CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.ClearItems()'></a>

## DataColumnCollection\<T\>\.ClearItems\(\) Method

Removes all elements from the [System\.Collections\.ObjectModel\.Collection&lt;&gt;](https://learn.microsoft.com/en-us/dotnet/api/system.collections.objectmodel.collection-1 'System\.Collections\.ObjectModel\.Collection\`1')\.

```csharp
protected override void ClearItems();
```

<a name='CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.RemoveItem(int)'></a>

## DataColumnCollection\<T\>\.RemoveItem\(int\) Method

Removes the element at the specified index of the [System\.Collections\.ObjectModel\.Collection&lt;&gt;](https://learn.microsoft.com/en-us/dotnet/api/system.collections.objectmodel.collection-1 'System\.Collections\.ObjectModel\.Collection\`1')\.

```csharp
protected override void RemoveItem(int index);
```
#### Parameters

<a name='CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.RemoveItem(int).index'></a>

`index` [System\.Int32](https://learn.microsoft.com/en-us/dotnet/api/system.int32 'System\.Int32')

The zero\-based index of the element to remove\.

#### Exceptions

[System\.ArgumentOutOfRangeException](https://learn.microsoft.com/en-us/dotnet/api/system.argumentoutofrangeexception 'System\.ArgumentOutOfRangeException')  
[index](CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.md#CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.RemoveItem(int).index 'CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.DataColumnCollection\<T\>\.RemoveItem\(int\)\.index') is less than zero\.\-or\-[index](CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.md#CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.RemoveItem(int).index 'CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.DataColumnCollection\<T\>\.RemoveItem\(int\)\.index') is equal to or greater than [System\.Collections\.ObjectModel\.Collection&lt;&gt;\.Count](https://learn.microsoft.com/en-us/dotnet/api/system.collections.objectmodel.collection-1.count 'System\.Collections\.ObjectModel\.Collection\`1\.Count')\.