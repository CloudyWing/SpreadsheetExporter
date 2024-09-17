#### [CloudyWing.SpreadsheetExporter](index.md 'index')
### [CloudyWing.SpreadsheetExporter.Templates](CloudyWing.SpreadsheetExporter.Templates.md 'CloudyWing.SpreadsheetExporter.Templates')

## MergedTemplate Class

The merged template. Merge templates into new template.

```csharp
public class MergedTemplate :
CloudyWing.SpreadsheetExporter.Templates.ITemplate
```

Inheritance [System.Object](https://docs.microsoft.com/en-us/dotnet/api/System.Object 'System.Object') &#129106; MergedTemplate

Implements [ITemplate](CloudyWing.SpreadsheetExporter.Templates.ITemplate.md 'CloudyWing.SpreadsheetExporter.Templates.ITemplate')

#### Exceptions

[System.ArgumentNullException](https://docs.microsoft.com/en-us/dotnet/api/System.ArgumentNullException 'System.ArgumentNullException')  
templates

### See Also
- [ITemplate](CloudyWing.SpreadsheetExporter.Templates.ITemplate.md 'CloudyWing.SpreadsheetExporter.Templates.ITemplate')
### Constructors

<a name='CloudyWing.SpreadsheetExporter.Templates.MergedTemplate.MergedTemplate(CloudyWing.SpreadsheetExporter.Templates.ITemplate[])'></a>

## MergedTemplate(ITemplate[]) Constructor

Initializes a new instance of the [MergedTemplate](CloudyWing.SpreadsheetExporter.Templates.MergedTemplate.md 'CloudyWing.SpreadsheetExporter.Templates.MergedTemplate') class.

```csharp
public MergedTemplate(params CloudyWing.SpreadsheetExporter.Templates.ITemplate[] templates);
```
#### Parameters

<a name='CloudyWing.SpreadsheetExporter.Templates.MergedTemplate.MergedTemplate(CloudyWing.SpreadsheetExporter.Templates.ITemplate[]).templates'></a>

`templates` [ITemplate](CloudyWing.SpreadsheetExporter.Templates.ITemplate.md 'CloudyWing.SpreadsheetExporter.Templates.ITemplate')[[]](https://docs.microsoft.com/en-us/dotnet/api/System.Array 'System.Array')

The templates.

<a name='CloudyWing.SpreadsheetExporter.Templates.MergedTemplate.MergedTemplate(System.Collections.Generic.IEnumerable_CloudyWing.SpreadsheetExporter.Templates.ITemplate_)'></a>

## MergedTemplate(IEnumerable<ITemplate>) Constructor

The merged template. Merge templates into new template.

```csharp
public MergedTemplate(System.Collections.Generic.IEnumerable<CloudyWing.SpreadsheetExporter.Templates.ITemplate> templates);
```
#### Parameters

<a name='CloudyWing.SpreadsheetExporter.Templates.MergedTemplate.MergedTemplate(System.Collections.Generic.IEnumerable_CloudyWing.SpreadsheetExporter.Templates.ITemplate_).templates'></a>

`templates` [System.Collections.Generic.IEnumerable&lt;](https://docs.microsoft.com/en-us/dotnet/api/System.Collections.Generic.IEnumerable-1 'System.Collections.Generic.IEnumerable`1')[ITemplate](CloudyWing.SpreadsheetExporter.Templates.ITemplate.md 'CloudyWing.SpreadsheetExporter.Templates.ITemplate')[&gt;](https://docs.microsoft.com/en-us/dotnet/api/System.Collections.Generic.IEnumerable-1 'System.Collections.Generic.IEnumerable`1')

The templates.

#### Exceptions

[System.ArgumentNullException](https://docs.microsoft.com/en-us/dotnet/api/System.ArgumentNullException 'System.ArgumentNullException')  
templates

### See Also
- [ITemplate](CloudyWing.SpreadsheetExporter.Templates.ITemplate.md 'CloudyWing.SpreadsheetExporter.Templates.ITemplate')
### Methods

<a name='CloudyWing.SpreadsheetExporter.Templates.MergedTemplate.GetContext()'></a>

## MergedTemplate.GetContext() Method

Gets the context.

```csharp
public CloudyWing.SpreadsheetExporter.Templates.TemplateContext GetContext();
```

Implements [GetContext()](CloudyWing.SpreadsheetExporter.Templates.ITemplate.md#CloudyWing.SpreadsheetExporter.Templates.ITemplate.GetContext() 'CloudyWing.SpreadsheetExporter.Templates.ITemplate.GetContext()')

#### Returns
[TemplateContext](CloudyWing.SpreadsheetExporter.Templates.TemplateContext.md 'CloudyWing.SpreadsheetExporter.Templates.TemplateContext')  
The template context.