#### [CloudyWing.SpreadsheetExporter](index.md 'index')
### [CloudyWing.SpreadsheetExporter.Config](CloudyWing.SpreadsheetExporter.Config.md 'CloudyWing.SpreadsheetExporter.Config')

## CellStyleConfiguration Class

The spreadsheet default style.

```csharp
public class CellStyleConfiguration
```

Inheritance [System.Object](https://docs.microsoft.com/en-us/dotnet/api/System.Object 'System.Object') &#129106; CellStyleConfiguration
### Constructors

<a name='CloudyWing.SpreadsheetExporter.Config.CellStyleConfiguration.CellStyleConfiguration()'></a>

## CellStyleConfiguration() Constructor

Initializes a new instance of the [CellStyleConfiguration](CloudyWing.SpreadsheetExporter.Config.CellStyleConfiguration.md 'CloudyWing.SpreadsheetExporter.Config.CellStyleConfiguration') class.  
  
```csharp  
CellStyle: {  
    HorizontalAlignment: None,  
    VerticalAlignment: Middle,  
    HasBorder: false,  
    WrapText: false,  
    Font: {  
        Name: "新細明體",  
        Size: 10,  
        IsBold: false,  
        IsItalic: false,  
        HasUnderline: false,  
        IsStrikeout: false  
    }  
}  
  
GridCellStyle: {  
    StyleBase: CellStyle  
}  
  
HeaderStyle: {  
    StyleBase: CellStyle  
    HorizontalAlignment: Center,  
    HasBorder: true,  
    Font: {  
        IsBold: true,  
    }  
}  
  
FieldStyle: {  
    StyleBase: CellStyle  
    HasBorder: true  
}  
```

```csharp
public CellStyleConfiguration();
```

<a name='CloudyWing.SpreadsheetExporter.Config.CellStyleConfiguration.CellStyleConfiguration(System.Action_CloudyWing.SpreadsheetExporter.Config.CellStyleSetuper_)'></a>

## CellStyleConfiguration(Action<CellStyleSetuper>) Constructor

Initializes a new instance of the [CellStyleConfiguration](CloudyWing.SpreadsheetExporter.Config.CellStyleConfiguration.md 'CloudyWing.SpreadsheetExporter.Config.CellStyleConfiguration') class.

```csharp
public CellStyleConfiguration(System.Action<CloudyWing.SpreadsheetExporter.Config.CellStyleSetuper> loader);
```
#### Parameters

<a name='CloudyWing.SpreadsheetExporter.Config.CellStyleConfiguration.CellStyleConfiguration(System.Action_CloudyWing.SpreadsheetExporter.Config.CellStyleSetuper_).loader'></a>

`loader` [System.Action&lt;](https://docs.microsoft.com/en-us/dotnet/api/System.Action-1 'System.Action`1')[CellStyleSetuper](CloudyWing.SpreadsheetExporter.Config.CellStyleSetuper.md 'CloudyWing.SpreadsheetExporter.Config.CellStyleSetuper')[&gt;](https://docs.microsoft.com/en-us/dotnet/api/System.Action-1 'System.Action`1')

The loader.
### Properties

<a name='CloudyWing.SpreadsheetExporter.Config.CellStyleConfiguration.CellStyle'></a>

## CellStyleConfiguration.CellStyle Property

Gets the cell style.

```csharp
public virtual CloudyWing.SpreadsheetExporter.CellStyle CellStyle { get; set; }
```

#### Property Value
[CellStyle](CloudyWing.SpreadsheetExporter.CellStyle.md 'CloudyWing.SpreadsheetExporter.CellStyle')  
The cell style.

<a name='CloudyWing.SpreadsheetExporter.Config.CellStyleConfiguration.FieldStyle'></a>

## CellStyleConfiguration.FieldStyle Property

Gets the field style.

```csharp
public virtual CloudyWing.SpreadsheetExporter.CellStyle FieldStyle { get; set; }
```

#### Property Value
[CellStyle](CloudyWing.SpreadsheetExporter.CellStyle.md 'CloudyWing.SpreadsheetExporter.CellStyle')  
The field style.

<a name='CloudyWing.SpreadsheetExporter.Config.CellStyleConfiguration.GridCellStyle'></a>

## CellStyleConfiguration.GridCellStyle Property

Gets the grid cell style.

```csharp
public virtual CloudyWing.SpreadsheetExporter.CellStyle GridCellStyle { get; set; }
```

#### Property Value
[CellStyle](CloudyWing.SpreadsheetExporter.CellStyle.md 'CloudyWing.SpreadsheetExporter.CellStyle')  
The grid cell style.

<a name='CloudyWing.SpreadsheetExporter.Config.CellStyleConfiguration.HeaderStyle'></a>

## CellStyleConfiguration.HeaderStyle Property

Gets the header style.

```csharp
public virtual CloudyWing.SpreadsheetExporter.CellStyle HeaderStyle { get; set; }
```

#### Property Value
[CellStyle](CloudyWing.SpreadsheetExporter.CellStyle.md 'CloudyWing.SpreadsheetExporter.CellStyle')  
The header style.