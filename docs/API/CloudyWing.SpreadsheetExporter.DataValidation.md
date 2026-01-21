#### [CloudyWing\.SpreadsheetExporter](index.md 'index')
### [CloudyWing\.SpreadsheetExporter](CloudyWing.SpreadsheetExporter.md 'CloudyWing\.SpreadsheetExporter')

## DataValidation Class

Represents data validation rules that can be applied to spreadsheet cells\.
Data validation restricts the type of data or values that users can enter into a cell\.

```csharp
public class DataValidation
```

Inheritance [System\.Object](https://learn.microsoft.com/en-us/dotnet/api/system.object 'System\.Object') &#129106; DataValidation
### Properties

<a name='CloudyWing.SpreadsheetExporter.DataValidation.ErrorMessage'></a>

## DataValidation\.ErrorMessage Property

Gets or sets the error message shown when invalid data is entered\.

```csharp
public string ErrorMessage { get; set; }
```

#### Property Value
[System\.String](https://learn.microsoft.com/en-us/dotnet/api/system.string 'System\.String')  
The error message\.

<a name='CloudyWing.SpreadsheetExporter.DataValidation.ErrorTitle'></a>

## DataValidation\.ErrorTitle Property

Gets or sets the title of the error message shown when invalid data is entered\.

```csharp
public string ErrorTitle { get; set; }
```

#### Property Value
[System\.String](https://learn.microsoft.com/en-us/dotnet/api/system.string 'System\.String')  
The error title\.

<a name='CloudyWing.SpreadsheetExporter.DataValidation.Formula'></a>

## DataValidation\.Formula Property

Gets or sets the formula for validation\.
Used with Custom validation type \(required\), and can also be used with Date, Time, Integer, Decimal, and TextLength validation types as an alternative to Value1/Value2\.
When using formulas with Date/Time/Numeric validations, the formula will be automatically prefixed with '=' if not already present\.

```csharp
public string Formula { get; set; }
```

#### Property Value
[System\.String](https://learn.microsoft.com/en-us/dotnet/api/system.string 'System\.String')  
The validation formula \(e\.g\., "TODAY\(\)", "=A1\+10"\)\.

<a name='CloudyWing.SpreadsheetExporter.DataValidation.IsBlankAllowed'></a>

## DataValidation\.IsBlankAllowed Property

Gets or sets a value indicating whether blank/empty values are allowed\.

```csharp
public bool IsBlankAllowed { get; set; }
```

#### Property Value
[System\.Boolean](https://learn.microsoft.com/en-us/dotnet/api/system.boolean 'System\.Boolean')  
`true` if blank values are allowed; otherwise, `false`\. Default is `true`\.

<a name='CloudyWing.SpreadsheetExporter.DataValidation.IsDropdownShown'></a>

## DataValidation\.IsDropdownShown Property

Gets or sets a value indicating whether the dropdown button is shown for list validation\.
Only applicable when ValidationType is List\.

```csharp
public bool IsDropdownShown { get; set; }
```

#### Property Value
[System\.Boolean](https://learn.microsoft.com/en-us/dotnet/api/system.boolean 'System\.Boolean')  
`true` if dropdown is shown; otherwise, `false`\. Default is `true`\.

<a name='CloudyWing.SpreadsheetExporter.DataValidation.IsErrorAlertShown'></a>

## DataValidation\.IsErrorAlertShown Property

Gets or sets a value indicating whether the error alert is shown when invalid data is entered\.

```csharp
public bool IsErrorAlertShown { get; set; }
```

#### Property Value
[System\.Boolean](https://learn.microsoft.com/en-us/dotnet/api/system.boolean 'System\.Boolean')  
`true` if error alert is shown; otherwise, `false`\. Default is `true`\.

<a name='CloudyWing.SpreadsheetExporter.DataValidation.IsInputPromptShown'></a>

## DataValidation\.IsInputPromptShown Property

Gets or sets a value indicating whether the input prompt is shown when the cell is selected\.

```csharp
public bool IsInputPromptShown { get; set; }
```

#### Property Value
[System\.Boolean](https://learn.microsoft.com/en-us/dotnet/api/system.boolean 'System\.Boolean')  
`true` if input prompt is shown; otherwise, `false`\. Default is `false`\.

<a name='CloudyWing.SpreadsheetExporter.DataValidation.ListItems'></a>

## DataValidation\.ListItems Property

Gets or sets the list of valid items for List validation type\.
Only used when ValidationType is List\.

```csharp
public System.Collections.Generic.IEnumerable<string> ListItems { get; set; }
```

#### Property Value
[System\.Collections\.Generic\.IEnumerable&lt;](https://learn.microsoft.com/en-us/dotnet/api/system.collections.generic.ienumerable-1 'System\.Collections\.Generic\.IEnumerable\`1')[System\.String](https://learn.microsoft.com/en-us/dotnet/api/system.string 'System\.String')[&gt;](https://learn.microsoft.com/en-us/dotnet/api/system.collections.generic.ienumerable-1 'System\.Collections\.Generic\.IEnumerable\`1')  
The list items\.

<a name='CloudyWing.SpreadsheetExporter.DataValidation.Operator'></a>

## DataValidation\.Operator Property

Gets or sets the comparison operator for the validation\.
Not required for List or Custom validation types\.

```csharp
public System.Nullable<CloudyWing.SpreadsheetExporter.DataValidationOperator> Operator { get; set; }
```

#### Property Value
[System\.Nullable&lt;](https://learn.microsoft.com/en-us/dotnet/api/system.nullable-1 'System\.Nullable\`1')[DataValidationOperator](CloudyWing.SpreadsheetExporter.DataValidationOperator.md 'CloudyWing\.SpreadsheetExporter\.DataValidationOperator')[&gt;](https://learn.microsoft.com/en-us/dotnet/api/system.nullable-1 'System\.Nullable\`1')  
The operator, or `null` if not applicable\.

<a name='CloudyWing.SpreadsheetExporter.DataValidation.PromptMessage'></a>

## DataValidation\.PromptMessage Property

Gets or sets the input prompt message shown when the cell is selected\.

```csharp
public string PromptMessage { get; set; }
```

#### Property Value
[System\.String](https://learn.microsoft.com/en-us/dotnet/api/system.string 'System\.String')  
The prompt message\.

<a name='CloudyWing.SpreadsheetExporter.DataValidation.PromptTitle'></a>

## DataValidation\.PromptTitle Property

Gets or sets the title of the input prompt message\.

```csharp
public string PromptTitle { get; set; }
```

#### Property Value
[System\.String](https://learn.microsoft.com/en-us/dotnet/api/system.string 'System\.String')  
The prompt title\.

<a name='CloudyWing.SpreadsheetExporter.DataValidation.ValidationType'></a>

## DataValidation\.ValidationType Property

Gets or sets the type of validation to apply\.

```csharp
public CloudyWing.SpreadsheetExporter.DataValidationType ValidationType { get; set; }
```

#### Property Value
[DataValidationType](CloudyWing.SpreadsheetExporter.DataValidationType.md 'CloudyWing\.SpreadsheetExporter\.DataValidationType')  
The validation type\.

<a name='CloudyWing.SpreadsheetExporter.DataValidation.Value1'></a>

## DataValidation\.Value1 Property

Gets or sets the first value for comparison\.
For Between/NotBetween operators, this is the minimum value\.
For other operators, this is the value to compare against\.

```csharp
public object Value1 { get; set; }
```

#### Property Value
[System\.Object](https://learn.microsoft.com/en-us/dotnet/api/system.object 'System\.Object')  
The first comparison value\.

<a name='CloudyWing.SpreadsheetExporter.DataValidation.Value2'></a>

## DataValidation\.Value2 Property

Gets or sets the second value for comparison\.
Only used with Between and NotBetween operators \(represents the maximum value\)\.

```csharp
public object Value2 { get; set; }
```

#### Property Value
[System\.Object](https://learn.microsoft.com/en-us/dotnet/api/system.object 'System\.Object')  
The second comparison value, or `null` if not applicable\.