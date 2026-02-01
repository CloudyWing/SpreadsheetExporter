# 自訂樣式

專案使用的儲存格樣式皆定義在 `SpreadsheetManager.DefaultCellStyles`，提供預設樣式給不同元件使用。

## 預設樣式

`CellStyleConfiguration` 提供以下預設樣式：

| 樣式 | 用途 |
| --- | --- |
| `CellStyle` | 預設儲存格樣式 |
| `GridCellStyle` | `GridTemplate` 使用的預設樣式 |
| `HeaderStyle` | `RecordSetTemplate` 標題列使用的預設樣式 |
| `FieldStyle` | `RecordSetTemplate` 資料列使用的預設樣式 |

## 更改預設樣式

### 基於現有樣式修改

以下範例將所有預設樣式的字體大小放大：

```csharp
SpreadsheetManager.DefaultCellStyles = new CellStyleConfiguration((setuper) => {
    CellStyleConfiguration oldCellStyles = SpreadsheetManager.DefaultCellStyles;
    
    setuper.CellStyle = oldCellStyles.CellStyle with {
        Font = oldCellStyles.CellStyle.Font with { Size = 16 }
    };
    
    setuper.GridCellStyle = oldCellStyles.GridCellStyle with {
        Font = oldCellStyles.GridCellStyle.Font with { Size = 14 }
    };
    
    setuper.HeaderStyle = oldCellStyles.HeaderStyle with {
        Font = oldCellStyles.HeaderStyle.Font with { Size = 14 }
    };
    
    setuper.FieldStyle = oldCellStyles.FieldStyle with {
        Font = oldCellStyles.FieldStyle.Font with { Size = 14 }
    };
});
```

### 完整自訂

您也可以使用 `new CellStyle()` 完整自訂所有屬性：

```csharp
SpreadsheetManager.DefaultCellStyles = new CellStyleConfiguration((setuper) => {
    CellFont customFont = new CellFont(
        fontName: "微軟正黑體",
        size: 12,
        color: Color.Black,
        style: FontStyles.None
    );
    
    setuper.CellStyle = new CellStyle(
        horizontalAlignment: HorizontalAlignment.Left,
        verticalAlignment: VerticalAlignment.Middle,
        hasBorder: false,
        wrapText: false,
        backgroundColor: Color.Empty,
        font: customFont
    );
    
    // ... 設定其他樣式 ...
});
```

## 從設定檔讀取樣式

### 設定檔結構

在專案中建立 `Spreadsheet.json` 並設定「複製到輸出目錄」：

```json
{
    "Spreadsheet": {
        "Cell": {
            "HorizontalAlignment": "Center",
            "VerticalAlignment": "Middle",
            "HasBorder": false,
            "WrapText": false,
            "Font": {
                "FontName": "新細明體",
                "FontSize": 10,
                "IsBold": false,
                "IsItalic": false,
                "HasUnderline": false,
                "IsStrikeout": false
            }
        }
    }
}
```

### 載入設定

在 `Program.cs` 中讀取並套用設定：

```csharp
public class ConfigSettings {
    public CellSettings Cell { get; set; }
}

public class CellSettings {
    public HorizontalAlignment HorizontalAlignment { get; set; }
    public VerticalAlignment VerticalAlignment { get; set; }
    public bool HasBorder { get; set; }
    public bool WrapText { get; set; }
    public FontSettings Font { get; set; }
}

public class FontSettings {
    public string FontName { get; set; }
    public short FontSize { get; set; }
    public bool IsBold { get; set; }
    public bool IsItalic { get; set; }
    public bool HasUnderline { get; set; }
    public bool IsStrikeout { get; set; }
}

// 建立設定載入器
IConfigurationRoot config = new ConfigurationBuilder()
    .SetBasePath(AppDomain.CurrentDomain.BaseDirectory)
    .AddJsonFile("Spreadsheet.json", optional: false, reloadOnChange: true)
    .Build();

// 設定自動重新載入
ChangeToken.OnChange(() => config.GetReloadToken(), () => {
    Initialize();
});

Initialize();

void Initialize() {
    ConfigSettings configSettings = config.GetSection("Spreadsheet").Get<ConfigSettings>();
    CellSettings cellSettings = configSettings.Cell;
    
    // 組合字型樣式
    FontStyles fontStyle = FontStyles.None;
    if (cellSettings.Font.IsBold) {
        fontStyle |= FontStyles.IsBold;
    }
    if (cellSettings.Font.IsItalic) {
        fontStyle |= FontStyles.IsItalic;
    }
    if (cellSettings.Font.HasUnderline) {
        fontStyle |= FontStyles.HasUnderline;
    }
    if (cellSettings.Font.IsStrikeout) {
        fontStyle |= FontStyles.IsStrikeout;
    }
    
    // 建立儲存格樣式
    CellStyle cellStyle = new CellStyle(
        cellSettings.HorizontalAlignment,
        cellSettings.VerticalAlignment,
        cellSettings.HasBorder,
        cellSettings.WrapText,
        Color.Empty,
        new CellFont(
            cellSettings.Font.FontName,
            cellSettings.Font.FontSize,
            Color.Black,
            fontStyle
        )
    );
    
    // 套用至全域設定
    SpreadsheetManager.DefaultCellStyles = new CellStyleConfiguration((setuper) => {
        setuper.CellStyle = cellStyle;
        setuper.GridCellStyle = cellStyle;
        setuper.HeaderStyle = cellStyle with {
            Font = cellStyle.Font with {
                Style = cellStyle.Font.Style | FontStyles.IsBold
            },
            HorizontalAlignment = HorizontalAlignment.Center,
            HasBorder = true
        };
        setuper.FieldStyle = cellStyle with {
            HorizontalAlignment = HorizontalAlignment.General,
            HasBorder = true
        };
    });
}
```

> [!TIP]
> 使用 `ChangeToken.OnChange` 可在設定檔變更時自動重新載入樣式，無需重新啟動應用程式。

## 相關主題

- [入門指南](getting-started.md) - 了解 SpreadsheetManager 的設定時機
- [Exporters](exporters.md) - 學習如何在 Exporter 層級套用全域樣式
- [Sheeters](sheeters.md) - 了解 Sheeter 如何使用預設樣式
- [Templates](templates.md) - 學習在 Template 中覆寫或使用自訂樣式
