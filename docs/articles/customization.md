# 自訂樣式

`SpreadsheetManager.DefaultCellStyles` 提供全域的預設儲存格樣式，各 template 類型會自動套用對應的預設值。

## 預設樣式

`CellStyleConfiguration` 包含以下四個樣式屬性：

| 屬性 | 用途 |
| --- | --- |
| `CellStyle` | 一般儲存格的基礎樣式 |
| `GridCellStyle` | `GridTemplate` 儲存格使用的預設樣式 |
| `HeaderStyle` | `RecordSetTemplate` 標題列使用的預設樣式 |
| `FieldStyle` | `RecordSetTemplate` 資料列使用的預設樣式 |

## 更改預設樣式

`CellStyleConfiguration` 使用 `init` 屬性，建立時透過物件初始化器覆寫需要變更的樣式，其餘保留建構子預設值：

### 基於現有樣式修改

```csharp
CellStyleConfiguration existing = SpreadsheetManager.DefaultCellStyles;

SpreadsheetManager.DefaultCellStyles = new CellStyleConfiguration {
    CellStyle = existing.CellStyle with {
        Font = existing.CellStyle.Font with { Size = 12 }
    },
    GridCellStyle = existing.GridCellStyle with {
        Font = existing.GridCellStyle.Font with { Size = 12 }
    },
    HeaderStyle = existing.HeaderStyle with {
        Font = existing.HeaderStyle.Font with { Size = 12 }
    },
    FieldStyle = existing.FieldStyle with {
        Font = existing.FieldStyle.Font with { Size = 12 }
    }
};
```

### 完整自訂

```csharp
CellFont customFont = new CellFont(
    Name: "微軟正黑體",
    Size: 12,
    Color: Color.Black,
    Style: FontStyles.None
);

CellStyle baseStyle = new CellStyle(
    HorizontalAlignment: HorizontalAlignment.Left,
    VerticalAlignment: VerticalAlignment.Middle,
    Font: customFont
);

SpreadsheetManager.DefaultCellStyles = new CellStyleConfiguration {
    CellStyle = baseStyle,
    GridCellStyle = baseStyle,
    HeaderStyle = baseStyle with {
        HorizontalAlignment = HorizontalAlignment.Center,
        HasBorder = true,
        Font = customFont with { Style = FontStyles.Bold }
    },
    FieldStyle = baseStyle with {
        HasBorder = true
    }
};
```

## 從設定檔讀取樣式

### 設定檔結構

在專案中建立 `spreadsheet.json`，並設定「複製到輸出目錄」：

```json
{
  "Spreadsheet": {
    "FontName": "新細明體",
    "FontSize": 10,
    "HorizontalAlignment": "Center",
    "VerticalAlignment": "Middle"
  }
}
```

### 載入設定

```csharp
IConfigurationRoot config = new ConfigurationBuilder()
    .SetBasePath(AppDomain.CurrentDomain.BaseDirectory)
    .AddJsonFile("spreadsheet.json", optional: false, reloadOnChange: true)
    .Build();

void ApplyStyles() {
    IConfigurationSection section = config.GetSection("Spreadsheet");

    CellFont font = new CellFont(
        Name: section["FontName"],
        Size: short.Parse(section["FontSize"]!)
    );

    CellStyle baseStyle = new CellStyle(
        HorizontalAlignment: Enum.Parse<HorizontalAlignment>(section["HorizontalAlignment"]!),
        VerticalAlignment: Enum.Parse<VerticalAlignment>(section["VerticalAlignment"]!),
        Font: font
    );

    SpreadsheetManager.DefaultCellStyles = new CellStyleConfiguration {
        CellStyle = baseStyle,
        GridCellStyle = baseStyle,
        HeaderStyle = baseStyle with {
            HasBorder = true,
            Font = font with { Style = FontStyles.Bold }
        },
        FieldStyle = baseStyle with { HasBorder = true }
    };
}

ChangeToken.OnChange(() => config.GetReloadToken(), ApplyStyles);
ApplyStyles();
```

> [!TIP]
> 使用 `ChangeToken.OnChange` 可在設定檔變更時自動重新載入樣式，無需重新啟動應用程式。

## 在 Template 層級覆寫樣式

各 template 也可以在加入欄位時直接傳入樣式，優先於全域預設值：

```csharp
CellStyle customHeader = SpreadsheetManager.DefaultCellStyles.HeaderStyle with {
    BackgroundColor = Color.FromArgb(31, 78, 121),
    Font = SpreadsheetManager.DefaultCellStyles.HeaderStyle.Font with {
        Color = Color.White
    }
};

template.Columns.Add<string>("客戶", order => order.Customer, headerStyle: customHeader);
```

## 相關主題

- [入門指南](getting-started.md) - 了解 SpreadsheetManager 的設定時機
- [SpreadsheetDocument](spreadsheet-document.md) - 學習文件層級的字型設定
- [Templates](templates.md) - 學習在 template 中覆寫或使用自訂樣式
