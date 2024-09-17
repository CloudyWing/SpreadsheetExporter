#### [CloudyWing.SpreadsheetExporter](index.md 'index')
### [CloudyWing.SpreadsheetExporter](CloudyWing.SpreadsheetExporter.md 'CloudyWing.SpreadsheetExporter')

## SheetCreatedEventArgs Class

Event arguments after sheet create.

```csharp
public class SheetCreatedEventArgs
```

Inheritance [System.Object](https://docs.microsoft.com/en-us/dotnet/api/System.Object 'System.Object') &#129106; SheetCreatedEventArgs

#### Exceptions

[System.ArgumentNullException](https://docs.microsoft.com/en-us/dotnet/api/System.ArgumentNullException 'System.ArgumentNullException')  
sheetObject  
            or  
            sheeterContext
### Constructors

<a name='CloudyWing.SpreadsheetExporter.SheetCreatedEventArgs.SheetCreatedEventArgs(object,CloudyWing.SpreadsheetExporter.SheeterContext)'></a>

## SheetCreatedEventArgs(object, SheeterContext) Constructor

Event arguments after sheet create.

```csharp
public SheetCreatedEventArgs(object sheetObject, CloudyWing.SpreadsheetExporter.SheeterContext sheeterContext);
```
#### Parameters

<a name='CloudyWing.SpreadsheetExporter.SheetCreatedEventArgs.SheetCreatedEventArgs(object,CloudyWing.SpreadsheetExporter.SheeterContext).sheetObject'></a>

`sheetObject` [System.Object](https://docs.microsoft.com/en-us/dotnet/api/System.Object 'System.Object')

The sheet object that represents the Excel sheet from the underlying Excel library.

<a name='CloudyWing.SpreadsheetExporter.SheetCreatedEventArgs.SheetCreatedEventArgs(object,CloudyWing.SpreadsheetExporter.SheeterContext).sheeterContext'></a>

`sheeterContext` [SheeterContext](CloudyWing.SpreadsheetExporter.SheeterContext.md 'CloudyWing.SpreadsheetExporter.SheeterContext')

The sheeter contexts.

#### Exceptions

[System.ArgumentNullException](https://docs.microsoft.com/en-us/dotnet/api/System.ArgumentNullException 'System.ArgumentNullException')  
sheetObject  
            or  
            sheeterContext
### Properties

<a name='CloudyWing.SpreadsheetExporter.SheetCreatedEventArgs.SheeterContext'></a>

## SheetCreatedEventArgs.SheeterContext Property

Gets the sheeter contexts.

```csharp
public CloudyWing.SpreadsheetExporter.SheeterContext SheeterContext { get; }
```

#### Property Value
[SheeterContext](CloudyWing.SpreadsheetExporter.SheeterContext.md 'CloudyWing.SpreadsheetExporter.SheeterContext')  
The sheeter contexts.

<a name='CloudyWing.SpreadsheetExporter.SheetCreatedEventArgs.SheetObject'></a>

## SheetCreatedEventArgs.SheetObject Property

Gets or sets the sheet object that represents the Excel sheet from the underlying Excel library.

```csharp
public object SheetObject { get; }
```

#### Property Value
[System.Object](https://docs.microsoft.com/en-us/dotnet/api/System.Object 'System.Object')  
The sheet object.