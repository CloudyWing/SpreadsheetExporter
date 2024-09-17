#### [CloudyWing.SpreadsheetExporter](index.md 'index')
### [CloudyWing.SpreadsheetExporter.Templates.RecordSet](CloudyWing.SpreadsheetExporter.Templates.RecordSet.md 'CloudyWing.SpreadsheetExporter.Templates.RecordSet')

## RecordSetTemplate<T> Class

The recordset template. Create cell information using set data source and data column.

```csharp
public class RecordSetTemplate<T> :
CloudyWing.SpreadsheetExporter.Templates.ITemplate
```
#### Type parameters

<a name='CloudyWing.SpreadsheetExporter.Templates.RecordSet.RecordSetTemplate_T_.T'></a>

`T`

The type of the record.

Inheritance [System.Object](https://docs.microsoft.com/en-us/dotnet/api/System.Object 'System.Object') &#129106; RecordSetTemplate<T>

Implements [ITemplate](CloudyWing.SpreadsheetExporter.Templates.ITemplate.md 'CloudyWing.SpreadsheetExporter.Templates.ITemplate')

#### Exceptions

[System.ArgumentNullException](https://docs.microsoft.com/en-us/dotnet/api/System.ArgumentNullException 'System.ArgumentNullException')  
dataSource

### See Also
- [ITemplate](CloudyWing.SpreadsheetExporter.Templates.ITemplate.md 'CloudyWing.SpreadsheetExporter.Templates.ITemplate')
### Constructors

<a name='CloudyWing.SpreadsheetExporter.Templates.RecordSet.RecordSetTemplate_T_.RecordSetTemplate(System.Collections.Generic.IEnumerable_T_)'></a>

## RecordSetTemplate(IEnumerable<T>) Constructor

The recordset template. Create cell information using set data source and data column.

```csharp
public RecordSetTemplate(System.Collections.Generic.IEnumerable<T> dataSource);
```
#### Parameters

<a name='CloudyWing.SpreadsheetExporter.Templates.RecordSet.RecordSetTemplate_T_.RecordSetTemplate(System.Collections.Generic.IEnumerable_T_).dataSource'></a>

`dataSource` [System.Collections.Generic.IEnumerable&lt;](https://docs.microsoft.com/en-us/dotnet/api/System.Collections.Generic.IEnumerable-1 'System.Collections.Generic.IEnumerable`1')[T](CloudyWing.SpreadsheetExporter.Templates.RecordSet.RecordSetTemplate_T_.md#CloudyWing.SpreadsheetExporter.Templates.RecordSet.RecordSetTemplate_T_.T 'CloudyWing.SpreadsheetExporter.Templates.RecordSet.RecordSetTemplate<T>.T')[&gt;](https://docs.microsoft.com/en-us/dotnet/api/System.Collections.Generic.IEnumerable-1 'System.Collections.Generic.IEnumerable`1')

The data source.

#### Exceptions

[System.ArgumentNullException](https://docs.microsoft.com/en-us/dotnet/api/System.ArgumentNullException 'System.ArgumentNullException')  
dataSource

### See Also
- [ITemplate](CloudyWing.SpreadsheetExporter.Templates.ITemplate.md 'CloudyWing.SpreadsheetExporter.Templates.ITemplate')
### Properties

<a name='CloudyWing.SpreadsheetExporter.Templates.RecordSet.RecordSetTemplate_T_.Columns'></a>

## RecordSetTemplate<T>.Columns Property

Gets the columns.

```csharp
public CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection<T> Columns { get; }
```

#### Property Value
[CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection&lt;](CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.md 'CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection<T>')[T](CloudyWing.SpreadsheetExporter.Templates.RecordSet.RecordSetTemplate_T_.md#CloudyWing.SpreadsheetExporter.Templates.RecordSet.RecordSetTemplate_T_.T 'CloudyWing.SpreadsheetExporter.Templates.RecordSet.RecordSetTemplate<T>.T')[&gt;](CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection_T_.md 'CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnCollection<T>')  
The columns.

<a name='CloudyWing.SpreadsheetExporter.Templates.RecordSet.RecordSetTemplate_T_.ColumnSpan'></a>

## RecordSetTemplate<T>.ColumnSpan Property

Gets the column span.

```csharp
public int ColumnSpan { get; }
```

#### Property Value
[System.Int32](https://docs.microsoft.com/en-us/dotnet/api/System.Int32 'System.Int32')  
The column span.

<a name='CloudyWing.SpreadsheetExporter.Templates.RecordSet.RecordSetTemplate_T_.DataSource'></a>

## RecordSetTemplate<T>.DataSource Property

Gets or sets the data source.

```csharp
public System.Collections.Generic.IEnumerable<T> DataSource { get; set; }
```

#### Property Value
[System.Collections.Generic.IEnumerable&lt;](https://docs.microsoft.com/en-us/dotnet/api/System.Collections.Generic.IEnumerable-1 'System.Collections.Generic.IEnumerable`1')[T](CloudyWing.SpreadsheetExporter.Templates.RecordSet.RecordSetTemplate_T_.md#CloudyWing.SpreadsheetExporter.Templates.RecordSet.RecordSetTemplate_T_.T 'CloudyWing.SpreadsheetExporter.Templates.RecordSet.RecordSetTemplate<T>.T')[&gt;](https://docs.microsoft.com/en-us/dotnet/api/System.Collections.Generic.IEnumerable-1 'System.Collections.Generic.IEnumerable`1')  
The data source.

<a name='CloudyWing.SpreadsheetExporter.Templates.RecordSet.RecordSetTemplate_T_.HeaderHeight'></a>

## RecordSetTemplate<T>.HeaderHeight Property

Gets or sets the height of the header.

```csharp
public System.Nullable<double> HeaderHeight { get; set; }
```

#### Property Value
[System.Nullable&lt;](https://docs.microsoft.com/en-us/dotnet/api/System.Nullable-1 'System.Nullable`1')[System.Double](https://docs.microsoft.com/en-us/dotnet/api/System.Double 'System.Double')[&gt;](https://docs.microsoft.com/en-us/dotnet/api/System.Nullable-1 'System.Nullable`1')  
The height of the header.

<a name='CloudyWing.SpreadsheetExporter.Templates.RecordSet.RecordSetTemplate_T_.RecordHeight'></a>

## RecordSetTemplate<T>.RecordHeight Property

Gets or sets the height of the record.

```csharp
public System.Nullable<double> RecordHeight { get; set; }
```

#### Property Value
[System.Nullable&lt;](https://docs.microsoft.com/en-us/dotnet/api/System.Nullable-1 'System.Nullable`1')[System.Double](https://docs.microsoft.com/en-us/dotnet/api/System.Double 'System.Double')[&gt;](https://docs.microsoft.com/en-us/dotnet/api/System.Nullable-1 'System.Nullable`1')  
The height of the record.

<a name='CloudyWing.SpreadsheetExporter.Templates.RecordSet.RecordSetTemplate_T_.RowSpan'></a>

## RecordSetTemplate<T>.RowSpan Property

Gets the row span.

```csharp
public int RowSpan { get; }
```

#### Property Value
[System.Int32](https://docs.microsoft.com/en-us/dotnet/api/System.Int32 'System.Int32')  
The row span.
### Methods

<a name='CloudyWing.SpreadsheetExporter.Templates.RecordSet.RecordSetTemplate_T_.GetContext()'></a>

## RecordSetTemplate<T>.GetContext() Method

Gets the context.

```csharp
public CloudyWing.SpreadsheetExporter.Templates.TemplateContext GetContext();
```

Implements [GetContext()](CloudyWing.SpreadsheetExporter.Templates.ITemplate.md#CloudyWing.SpreadsheetExporter.Templates.ITemplate.GetContext() 'CloudyWing.SpreadsheetExporter.Templates.ITemplate.GetContext()')

#### Returns
[TemplateContext](CloudyWing.SpreadsheetExporter.Templates.TemplateContext.md 'CloudyWing.SpreadsheetExporter.Templates.TemplateContext')  
The template context.