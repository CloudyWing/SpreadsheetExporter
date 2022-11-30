#### [CloudyWing.SpreadsheetExporter](index.md 'index')
### [CloudyWing.SpreadsheetExporter](CloudyWing.SpreadsheetExporter.md 'CloudyWing.SpreadsheetExporter')

## Sheeter Class

The sheeter.

```csharp
public class Sheeter
```

Inheritance [System.Object](https://docs.microsoft.com/en-us/dotnet/api/System.Object 'System.Object') &#129106; Sheeter
### Properties

<a name='CloudyWing.SpreadsheetExporter.Sheeter.ColumnWidths'></a>

## Sheeter.ColumnWidths Property

Gets the width of the columns.

```csharp
public System.Collections.Generic.IReadOnlyDictionary<int,double> ColumnWidths { get; }
```

#### Property Value
[System.Collections.Generic.IReadOnlyDictionary&lt;](https://docs.microsoft.com/en-us/dotnet/api/System.Collections.Generic.IReadOnlyDictionary-2 'System.Collections.Generic.IReadOnlyDictionary`2')[System.Int32](https://docs.microsoft.com/en-us/dotnet/api/System.Int32 'System.Int32')[,](https://docs.microsoft.com/en-us/dotnet/api/System.Collections.Generic.IReadOnlyDictionary-2 'System.Collections.Generic.IReadOnlyDictionary`2')[System.Double](https://docs.microsoft.com/en-us/dotnet/api/System.Double 'System.Double')[&gt;](https://docs.microsoft.com/en-us/dotnet/api/System.Collections.Generic.IReadOnlyDictionary-2 'System.Collections.Generic.IReadOnlyDictionary`2')  
The width of columns.

<a name='CloudyWing.SpreadsheetExporter.Sheeter.PageSettings'></a>

## Sheeter.PageSettings Property

Gets the page settings.

```csharp
public CloudyWing.SpreadsheetExporter.PageSettings PageSettings { get; }
```

#### Property Value
[PageSettings](CloudyWing.SpreadsheetExporter.PageSettings.md 'CloudyWing.SpreadsheetExporter.PageSettings')  
The page settings.

<a name='CloudyWing.SpreadsheetExporter.Sheeter.Password'></a>

## Sheeter.Password Property

Gets or sets the password.

```csharp
public string Password { get; set; }
```

#### Property Value
[System.String](https://docs.microsoft.com/en-us/dotnet/api/System.String 'System.String')  
The password.

<a name='CloudyWing.SpreadsheetExporter.Sheeter.SheetName'></a>

## Sheeter.SheetName Property

Gets or sets the name of the sheet.

```csharp
public string SheetName { get; set; }
```

#### Property Value
[System.String](https://docs.microsoft.com/en-us/dotnet/api/System.String 'System.String')  
The name of the sheet.

<a name='CloudyWing.SpreadsheetExporter.Sheeter.Templates'></a>

## Sheeter.Templates Property

Gets the templates.

```csharp
public System.Collections.Generic.IReadOnlyList<CloudyWing.SpreadsheetExporter.Templates.ITemplate> Templates { get; }
```

#### Property Value
[System.Collections.Generic.IReadOnlyList&lt;](https://docs.microsoft.com/en-us/dotnet/api/System.Collections.Generic.IReadOnlyList-1 'System.Collections.Generic.IReadOnlyList`1')[ITemplate](CloudyWing.SpreadsheetExporter.Templates.ITemplate.md 'CloudyWing.SpreadsheetExporter.Templates.ITemplate')[&gt;](https://docs.microsoft.com/en-us/dotnet/api/System.Collections.Generic.IReadOnlyList-1 'System.Collections.Generic.IReadOnlyList`1')  
The templates.

<a name='CloudyWing.SpreadsheetExporter.Sheeter.Watermark'></a>

## Sheeter.Watermark Property

Gets or sets the watermark.

```csharp
public System.Drawing.Image Watermark { get; set; }
```

#### Property Value
[System.Drawing.Image](https://docs.microsoft.com/en-us/dotnet/api/System.Drawing.Image 'System.Drawing.Image')  
The watermark.
### Methods

<a name='CloudyWing.SpreadsheetExporter.Sheeter.AddTemplate(CloudyWing.SpreadsheetExporter.Templates.ITemplate)'></a>

## Sheeter.AddTemplate(ITemplate) Method

Adds the template.

```csharp
public void AddTemplate(CloudyWing.SpreadsheetExporter.Templates.ITemplate template);
```
#### Parameters

<a name='CloudyWing.SpreadsheetExporter.Sheeter.AddTemplate(CloudyWing.SpreadsheetExporter.Templates.ITemplate).template'></a>

`template` [ITemplate](CloudyWing.SpreadsheetExporter.Templates.ITemplate.md 'CloudyWing.SpreadsheetExporter.Templates.ITemplate')

The template.

<a name='CloudyWing.SpreadsheetExporter.Sheeter.AddTemplates(CloudyWing.SpreadsheetExporter.Templates.ITemplate[])'></a>

## Sheeter.AddTemplates(ITemplate[]) Method

Adds the templates.

```csharp
public void AddTemplates(params CloudyWing.SpreadsheetExporter.Templates.ITemplate[] templates);
```
#### Parameters

<a name='CloudyWing.SpreadsheetExporter.Sheeter.AddTemplates(CloudyWing.SpreadsheetExporter.Templates.ITemplate[]).templates'></a>

`templates` [ITemplate](CloudyWing.SpreadsheetExporter.Templates.ITemplate.md 'CloudyWing.SpreadsheetExporter.Templates.ITemplate')[[]](https://docs.microsoft.com/en-us/dotnet/api/System.Array 'System.Array')

The templates.

<a name='CloudyWing.SpreadsheetExporter.Sheeter.AddTemplates(System.Collections.Generic.IEnumerable_CloudyWing.SpreadsheetExporter.Templates.ITemplate_)'></a>

## Sheeter.AddTemplates(IEnumerable<ITemplate>) Method

Adds the templates.

```csharp
public void AddTemplates(System.Collections.Generic.IEnumerable<CloudyWing.SpreadsheetExporter.Templates.ITemplate> templates);
```
#### Parameters

<a name='CloudyWing.SpreadsheetExporter.Sheeter.AddTemplates(System.Collections.Generic.IEnumerable_CloudyWing.SpreadsheetExporter.Templates.ITemplate_).templates'></a>

`templates` [System.Collections.Generic.IEnumerable&lt;](https://docs.microsoft.com/en-us/dotnet/api/System.Collections.Generic.IEnumerable-1 'System.Collections.Generic.IEnumerable`1')[ITemplate](CloudyWing.SpreadsheetExporter.Templates.ITemplate.md 'CloudyWing.SpreadsheetExporter.Templates.ITemplate')[&gt;](https://docs.microsoft.com/en-us/dotnet/api/System.Collections.Generic.IEnumerable-1 'System.Collections.Generic.IEnumerable`1')

The templates.

<a name='CloudyWing.SpreadsheetExporter.Sheeter.SetColumnWidth(int,double)'></a>

## Sheeter.SetColumnWidth(int, double) Method

Sets the width of the column.

```csharp
public void SetColumnWidth(int index, double width);
```
#### Parameters

<a name='CloudyWing.SpreadsheetExporter.Sheeter.SetColumnWidth(int,double).index'></a>

`index` [System.Int32](https://docs.microsoft.com/en-us/dotnet/api/System.Int32 'System.Int32')

The index.

<a name='CloudyWing.SpreadsheetExporter.Sheeter.SetColumnWidth(int,double).width'></a>

`width` [System.Double](https://docs.microsoft.com/en-us/dotnet/api/System.Double 'System.Double')

The width. If width is `0`, hide width. if the width is `-1`, the column width will be adjusted automatically.