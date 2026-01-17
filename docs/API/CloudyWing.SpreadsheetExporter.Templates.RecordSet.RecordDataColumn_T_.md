#### [CloudyWing\.SpreadsheetExporter](index.md 'index')
### [CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet](CloudyWing.SpreadsheetExporter.Templates.RecordSet.md 'CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet')

## RecordDataColumn\<T\> Class

The simple data column\.

```csharp
internal class RecordDataColumn<T> : CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnBase<T>
```
#### Type parameters

<a name='CloudyWing.SpreadsheetExporter.Templates.RecordSet.RecordDataColumn_T_.T'></a>

`T`

The type of the record\.

Inheritance [System\.Object](https://learn.microsoft.com/en-us/dotnet/api/system.object 'System\.Object') &#129106; [CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.DataColumnBase&lt;](CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnBase_T_.md 'CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.DataColumnBase\<T\>')[T](CloudyWing.SpreadsheetExporter.Templates.RecordSet.RecordDataColumn_T_.md#CloudyWing.SpreadsheetExporter.Templates.RecordSet.RecordDataColumn_T_.T 'CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.RecordDataColumn\<T\>\.T')[&gt;](CloudyWing.SpreadsheetExporter.Templates.RecordSet.DataColumnBase_T_.md 'CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.DataColumnBase\<T\>') &#129106; RecordDataColumn\<T\>
### Properties

<a name='CloudyWing.SpreadsheetExporter.Templates.RecordSet.RecordDataColumn_T_.FieldFormulaGenerator'></a>

## RecordDataColumn\<T\>\.FieldFormulaGenerator Property

Gets the field formula generator\.

```csharp
public System.Func<CloudyWing.SpreadsheetExporter.Templates.RecordSet.RecordContext<T>,string> FieldFormulaGenerator { get; set; }
```

#### Property Value
[System\.Func&lt;](https://learn.microsoft.com/en-us/dotnet/api/system.func-2 'System\.Func\`2')[CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.RecordContext&lt;](CloudyWing.SpreadsheetExporter.Templates.RecordSet.RecordContext_T_.md 'CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.RecordContext\<T\>')[T](CloudyWing.SpreadsheetExporter.Templates.RecordSet.RecordDataColumn_T_.md#CloudyWing.SpreadsheetExporter.Templates.RecordSet.RecordDataColumn_T_.T 'CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.RecordDataColumn\<T\>\.T')[&gt;](CloudyWing.SpreadsheetExporter.Templates.RecordSet.RecordContext_T_.md 'CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.RecordContext\<T\>')[,](https://learn.microsoft.com/en-us/dotnet/api/system.func-2 'System\.Func\`2')[System\.String](https://learn.microsoft.com/en-us/dotnet/api/system.string 'System\.String')[&gt;](https://learn.microsoft.com/en-us/dotnet/api/system.func-2 'System\.Func\`2')  
The field formula generator\.

<a name='CloudyWing.SpreadsheetExporter.Templates.RecordSet.RecordDataColumn_T_.FieldStyleGenerator'></a>

## RecordDataColumn\<T\>\.FieldStyleGenerator Property

Gets or sets the field style generator\.

```csharp
public System.Func<CloudyWing.SpreadsheetExporter.Templates.RecordSet.RecordContext<T>,CloudyWing.SpreadsheetExporter.CellStyle> FieldStyleGenerator { get; set; }
```

#### Property Value
[System\.Func&lt;](https://learn.microsoft.com/en-us/dotnet/api/system.func-2 'System\.Func\`2')[CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.RecordContext&lt;](CloudyWing.SpreadsheetExporter.Templates.RecordSet.RecordContext_T_.md 'CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.RecordContext\<T\>')[T](CloudyWing.SpreadsheetExporter.Templates.RecordSet.RecordDataColumn_T_.md#CloudyWing.SpreadsheetExporter.Templates.RecordSet.RecordDataColumn_T_.T 'CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.RecordDataColumn\<T\>\.T')[&gt;](CloudyWing.SpreadsheetExporter.Templates.RecordSet.RecordContext_T_.md 'CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.RecordContext\<T\>')[,](https://learn.microsoft.com/en-us/dotnet/api/system.func-2 'System\.Func\`2')[CellStyle](CloudyWing.SpreadsheetExporter.CellStyle.md 'CloudyWing\.SpreadsheetExporter\.CellStyle')[&gt;](https://learn.microsoft.com/en-us/dotnet/api/system.func-2 'System\.Func\`2')  
The field style generator\.

<a name='CloudyWing.SpreadsheetExporter.Templates.RecordSet.RecordDataColumn_T_.FieldValueGenerator'></a>

## RecordDataColumn\<T\>\.FieldValueGenerator Property

Gets the field value generator\.

```csharp
public System.Func<CloudyWing.SpreadsheetExporter.Templates.RecordSet.RecordContext<T>,object> FieldValueGenerator { get; set; }
```

#### Property Value
[System\.Func&lt;](https://learn.microsoft.com/en-us/dotnet/api/system.func-2 'System\.Func\`2')[CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.RecordContext&lt;](CloudyWing.SpreadsheetExporter.Templates.RecordSet.RecordContext_T_.md 'CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.RecordContext\<T\>')[T](CloudyWing.SpreadsheetExporter.Templates.RecordSet.RecordDataColumn_T_.md#CloudyWing.SpreadsheetExporter.Templates.RecordSet.RecordDataColumn_T_.T 'CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.RecordDataColumn\<T\>\.T')[&gt;](CloudyWing.SpreadsheetExporter.Templates.RecordSet.RecordContext_T_.md 'CloudyWing\.SpreadsheetExporter\.Templates\.RecordSet\.RecordContext\<T\>')[,](https://learn.microsoft.com/en-us/dotnet/api/system.func-2 'System\.Func\`2')[System\.Object](https://learn.microsoft.com/en-us/dotnet/api/system.object 'System\.Object')[&gt;](https://learn.microsoft.com/en-us/dotnet/api/system.func-2 'System\.Func\`2')  
The field value generator\.