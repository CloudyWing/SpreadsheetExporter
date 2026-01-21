#### [CloudyWing\.SpreadsheetExporter](index.md 'index')
### [CloudyWing\.SpreadsheetExporter](CloudyWing.SpreadsheetExporter.md 'CloudyWing\.SpreadsheetExporter')

## Sheeter Class

The sheeter\.

```csharp
public class Sheeter
```

Inheritance [System\.Object](https://learn.microsoft.com/en-us/dotnet/api/system.object 'System\.Object') &#129106; Sheeter
### Properties

<a name='CloudyWing.SpreadsheetExporter.Sheeter.ColumnWidths'></a>

## Sheeter\.ColumnWidths Property

Gets the width of the columns\.

```csharp
public System.Collections.Generic.IReadOnlyDictionary<int,double> ColumnWidths { get; }
```

#### Property Value
[System\.Collections\.Generic\.IReadOnlyDictionary&lt;](https://learn.microsoft.com/en-us/dotnet/api/system.collections.generic.ireadonlydictionary-2 'System\.Collections\.Generic\.IReadOnlyDictionary\`2')[System\.Int32](https://learn.microsoft.com/en-us/dotnet/api/system.int32 'System\.Int32')[,](https://learn.microsoft.com/en-us/dotnet/api/system.collections.generic.ireadonlydictionary-2 'System\.Collections\.Generic\.IReadOnlyDictionary\`2')[System\.Double](https://learn.microsoft.com/en-us/dotnet/api/system.double 'System\.Double')[&gt;](https://learn.microsoft.com/en-us/dotnet/api/system.collections.generic.ireadonlydictionary-2 'System\.Collections\.Generic\.IReadOnlyDictionary\`2')  
The width of columns\.

<a name='CloudyWing.SpreadsheetExporter.Sheeter.DefaultRowHeight'></a>

## Sheeter\.DefaultRowHeight Property

Gets or sets the default height of the row\.

```csharp
public System.Nullable<double> DefaultRowHeight { get; set; }
```

#### Property Value
[System\.Nullable&lt;](https://learn.microsoft.com/en-us/dotnet/api/system.nullable-1 'System\.Nullable\`1')[System\.Double](https://learn.microsoft.com/en-us/dotnet/api/system.double 'System\.Double')[&gt;](https://learn.microsoft.com/en-us/dotnet/api/system.nullable-1 'System\.Nullable\`1')  
The default height of the row\.

<a name='CloudyWing.SpreadsheetExporter.Sheeter.PageSettings'></a>

## Sheeter\.PageSettings Property

Gets the page settings\.

```csharp
public CloudyWing.SpreadsheetExporter.PageSettings PageSettings { get; }
```

#### Property Value
[PageSettings](CloudyWing.SpreadsheetExporter.PageSettings.md 'CloudyWing\.SpreadsheetExporter\.PageSettings')  
The page settings\.

<a name='CloudyWing.SpreadsheetExporter.Sheeter.Password'></a>

## Sheeter\.Password Property

Gets or sets the password\.

```csharp
public string Password { get; set; }
```

#### Property Value
[System\.String](https://learn.microsoft.com/en-us/dotnet/api/system.string 'System\.String')  
The password\.

<a name='CloudyWing.SpreadsheetExporter.Sheeter.SheetName'></a>

## Sheeter\.SheetName Property

Gets or sets the name of the sheet\.

```csharp
public string SheetName { get; set; }
```

#### Property Value
[System\.String](https://learn.microsoft.com/en-us/dotnet/api/system.string 'System\.String')  
The name of the sheet\.

<a name='CloudyWing.SpreadsheetExporter.Sheeter.Templates'></a>

## Sheeter\.Templates Property

Gets the templates\.

```csharp
public System.Collections.Generic.IReadOnlyList<CloudyWing.SpreadsheetExporter.Templates.ITemplate> Templates { get; }
```

#### Property Value
[System\.Collections\.Generic\.IReadOnlyList&lt;](https://learn.microsoft.com/en-us/dotnet/api/system.collections.generic.ireadonlylist-1 'System\.Collections\.Generic\.IReadOnlyList\`1')[ITemplate](CloudyWing.SpreadsheetExporter.Templates.ITemplate.md 'CloudyWing\.SpreadsheetExporter\.Templates\.ITemplate')[&gt;](https://learn.microsoft.com/en-us/dotnet/api/system.collections.generic.ireadonlylist-1 'System\.Collections\.Generic\.IReadOnlyList\`1')  
The templates\.
### Methods

<a name='CloudyWing.SpreadsheetExporter.Sheeter.AddTemplate(CloudyWing.SpreadsheetExporter.Templates.ITemplate)'></a>

## Sheeter\.AddTemplate\(ITemplate\) Method

Adds the template\.

```csharp
public void AddTemplate(CloudyWing.SpreadsheetExporter.Templates.ITemplate template);
```
#### Parameters

<a name='CloudyWing.SpreadsheetExporter.Sheeter.AddTemplate(CloudyWing.SpreadsheetExporter.Templates.ITemplate).template'></a>

`template` [ITemplate](CloudyWing.SpreadsheetExporter.Templates.ITemplate.md 'CloudyWing\.SpreadsheetExporter\.Templates\.ITemplate')

The template\.

<a name='CloudyWing.SpreadsheetExporter.Sheeter.AddTemplates(CloudyWing.SpreadsheetExporter.Templates.ITemplate[])'></a>

## Sheeter\.AddTemplates\(ITemplate\[\]\) Method

Adds the templates\.

```csharp
public void AddTemplates(params CloudyWing.SpreadsheetExporter.Templates.ITemplate[] templates);
```
#### Parameters

<a name='CloudyWing.SpreadsheetExporter.Sheeter.AddTemplates(CloudyWing.SpreadsheetExporter.Templates.ITemplate[]).templates'></a>

`templates` [ITemplate](CloudyWing.SpreadsheetExporter.Templates.ITemplate.md 'CloudyWing\.SpreadsheetExporter\.Templates\.ITemplate')[\[\]](https://learn.microsoft.com/en-us/dotnet/api/system.array 'System\.Array')

The templates\.

<a name='CloudyWing.SpreadsheetExporter.Sheeter.AddTemplates(System.Collections.Generic.IEnumerable_CloudyWing.SpreadsheetExporter.Templates.ITemplate_)'></a>

## Sheeter\.AddTemplates\(IEnumerable\<ITemplate\>\) Method

Adds the templates\.

```csharp
public void AddTemplates(System.Collections.Generic.IEnumerable<CloudyWing.SpreadsheetExporter.Templates.ITemplate> templates);
```
#### Parameters

<a name='CloudyWing.SpreadsheetExporter.Sheeter.AddTemplates(System.Collections.Generic.IEnumerable_CloudyWing.SpreadsheetExporter.Templates.ITemplate_).templates'></a>

`templates` [System\.Collections\.Generic\.IEnumerable&lt;](https://learn.microsoft.com/en-us/dotnet/api/system.collections.generic.ienumerable-1 'System\.Collections\.Generic\.IEnumerable\`1')[ITemplate](CloudyWing.SpreadsheetExporter.Templates.ITemplate.md 'CloudyWing\.SpreadsheetExporter\.Templates\.ITemplate')[&gt;](https://learn.microsoft.com/en-us/dotnet/api/system.collections.generic.ienumerable-1 'System\.Collections\.Generic\.IEnumerable\`1')

The templates\.

<a name='CloudyWing.SpreadsheetExporter.Sheeter.SetColumnWidth(int,double)'></a>

## Sheeter\.SetColumnWidth\(int, double\) Method

Sets the width of the column\.

```csharp
public void SetColumnWidth(int index, double width);
```
#### Parameters

<a name='CloudyWing.SpreadsheetExporter.Sheeter.SetColumnWidth(int,double).index'></a>

`index` [System\.Int32](https://learn.microsoft.com/en-us/dotnet/api/system.int32 'System\.Int32')

The zero\-based column index\.

<a name='CloudyWing.SpreadsheetExporter.Sheeter.SetColumnWidth(int,double).width'></a>

`width` [System\.Double](https://learn.microsoft.com/en-us/dotnet/api/system.double 'System\.Double')

The column width in characters\.
            Use [HiddenColumn](CloudyWing.SpreadsheetExporter.Constants.md#CloudyWing.SpreadsheetExporter.Constants.HiddenColumn 'CloudyWing\.SpreadsheetExporter\.Constants\.HiddenColumn') \(0\) to hide the column\.
            Use [AutoFitColumnWidth](CloudyWing.SpreadsheetExporter.Constants.md#CloudyWing.SpreadsheetExporter.Constants.AutoFitColumnWidth 'CloudyWing\.SpreadsheetExporter\.Constants\.AutoFitColumnWidth') \(\-1\) to auto\-fit the column width\.
            Positive values specify the exact width in character units\.