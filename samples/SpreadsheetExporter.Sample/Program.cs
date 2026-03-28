using System.Data;
using System.Drawing;
using CloudyWing.SpreadsheetExporter;
using CloudyWing.SpreadsheetExporter.Renderer.ClosedXML;
using CloudyWing.SpreadsheetExporter.Templates;
using CloudyWing.SpreadsheetExporter.Templates.DataTable;
using CloudyWing.SpreadsheetExporter.Templates.Grid;
using CloudyWing.SpreadsheetExporter.Templates.RecordSet;

SpreadsheetManager.SetRenderer(static () => new ExcelRenderer());

string outputDirectory = Path.Combine(AppContext.BaseDirectory, "artifacts");
Directory.CreateDirectory(outputDirectory);

SpreadsheetDocument fluentDocument = CreateFluentDocument();
string fluentPath = Path.Combine(outputDirectory, $"sales-report{fluentDocument.FileNameExtension}");
fluentDocument.ExportFile(fluentPath);

SpreadsheetDocument jsonDocument = SpreadsheetDocument.FromJson(CreateJsonTemplate());
string jsonPath = Path.Combine(outputDirectory, $"sales-report-json{jsonDocument.FileNameExtension}");
jsonDocument.ExportFile(jsonPath);

Console.WriteLine("SpreadsheetExporter Sample (ClosedXML)");
Console.WriteLine($"Fluent API workbook : {fluentPath}");
Console.WriteLine($"JSON template output: {jsonPath}");

// ────────────────────────────────────────────────
// Fluent API document
// ────────────────────────────────────────────────

static SpreadsheetDocument CreateFluentDocument() {
    SpreadsheetDocument document = SpreadsheetManager
        .CreateDocument()
        .SetDefaultFont(new CellFont("Calibri", 11));

    document
        .CreateSheet("Overview", defaultRowHeight: 20)
        .SetColumnWidth(0, 20)
        .SetColumnWidth(1, 30)
        .SetColumnWidth(2, 20)
        .SetColumnWidth(3, 30)
        .AddTemplate(CreateOverviewTemplate());

    document
        .CreateSheet("Orders", defaultRowHeight: 20)
        .ConfigurePageSettings(x => x.PageOrientation = PageOrientation.Landscape)
        .SetColumnWidth(0, 12)
        .SetColumnWidth(1, 22)
        .SetColumnWidth(2, 14)
        .SetColumnWidth(3, 16)
        .SetColumnWidth(4, 16)
        .SetColumnWidth(5, 16)
        .SetFreezePanes(0, 2)
        .SetAutoFilterEnabled()
        .AddTemplate(CreateOrdersTemplate(GetSampleOrders()));

    document
        .CreateSheet("Products", defaultRowHeight: 20)
        .SetColumnWidth(0, 10)
        .SetColumnWidth(1, 28)
        .SetColumnWidth(2, 16)
        .SetColumnWidth(3, 14)
        .SetColumnWidth(4, 12)
        .AddTemplate(CreateProductsTemplate(GetSampleProducts()));

    return document;
}

static GridTemplate CreateOverviewTemplate() {
    CellStyle titleStyle = new(
        HorizontalAlignment: HorizontalAlignment.Center,
        VerticalAlignment: VerticalAlignment.Middle,
        HasBorder: true,
        BackgroundColor: Color.FromArgb(31, 78, 121),
        Font: new CellFont("Calibri", 14, Color.White, FontStyles.Bold)
    );
    CellStyle labelStyle = new(
        HasBorder: true,
        BackgroundColor: Color.FromArgb(221, 235, 247),
        Font: new CellFont("Calibri", 11, Style: FontStyles.Bold)
    );
    CellStyle valueStyle = new(HasBorder: true);
    CellStyle wrapStyle = valueStyle with { WrapText = true };
    CellStyle statusStyle = valueStyle with { BackgroundColor = Color.FromArgb(226, 239, 218) };

    return new GridTemplate()
        .CreateRow(28)
        .CreateCell("SpreadsheetExporter ClosedXML — Feature Verification", columnSpan: 4, cellStyle: titleStyle)
        .CreateRow()
        .CreateCell("Generated", cellStyle: labelStyle)
        .CreateCell(DateTimeOffset.Now.ToString("yyyy-MM-dd HH:mm:ss zzz"), cellStyle: valueStyle)
        .CreateCell("Renderer", cellStyle: labelStyle)
        .CreateCell("ClosedXML", cellStyle: valueStyle)
        .CreateRow()
        .CreateCell("Template Source", cellStyle: labelStyle)
        .CreateCell("Fluent API", cellStyle: valueStyle)
        .CreateCell("Status", cellStyle: labelStyle)
        .CreateCell(
            cell => {
                cell.ValueGenerator = static (_, _) => "Ready";
                cell.DataValidationGenerator = static (_, _) => new DataValidation {
                    ValidationType = DataValidationType.List,
                    ListItems = ["Draft", "Ready", "Archived"],
                    PromptTitle = "Workbook status",
                    PromptMessage = "Choose one of the supported workflow states.",
                    IsInputPromptShown = true
                };
            },
            cellStyle: statusStyle
        )
        .CreateRow(36)
        .CreateCell("Notes", cellStyle: labelStyle)
        .CreateCell(
            "This workbook verifies: GridTemplate (merge, style, validation), " +
            "RecordSetTemplate (hierarchy, freeze, autofilter, formula, validation), " +
            "DataTableTemplate (auto-columns, custom columns, style generators), " +
            "MergedTemplate (header + data combined).",
            columnSpan: 3,
            cellStyle: wrapStyle
        );
}

static RecordSetTemplate<SampleOrder> CreateOrdersTemplate(IEnumerable<SampleOrder> orders) {
    CellStyle headerStyle = SpreadsheetManager.DefaultCellStyles.HeaderStyle;
    CellStyle currencyStyle = SpreadsheetManager.DefaultCellStyles.FieldStyle with {
        HorizontalAlignment = HorizontalAlignment.Right,
        DataFormat = "#,##0.00"
    };
    CellStyle intStyle = SpreadsheetManager.DefaultCellStyles.FieldStyle with {
        HorizontalAlignment = HorizontalAlignment.Right
    };

    RecordSetTemplate<SampleOrder> template = new(orders) {
        HeaderHeight = 22,
        RecordHeight = 20
    };

    // 階層欄位：Order Info / Financial
    template.Columns
        .Add("Order Info")
        .GetLastColumn().ChildColumns
            .Add("Order ID", static o => o.OrderId, fieldStyleGenerator: _ => intStyle)
            .Add("Customer", static o => o.Customer)
            .Add(
                "Region",
                static o => o.Region,
                configureGenerators: static cfg => cfg.UseDataValidation(
                    _ => new DataValidation {
                        ValidationType = DataValidationType.List,
                        ListItems = ["North", "South", "East", "West"],
                        ErrorTitle = "Unsupported region",
                        ErrorMessage = "Region must be one of the configured sales territories."
                    }
                )
            )
            .Parent
        .Add("Financial")
        .GetLastColumn().ChildColumns
            .Add("Amount", static o => o.Amount, fieldStyleGenerator: _ => currencyStyle)
            .Add(
                "Margin",
                configureGenerators: static cfg => cfg.UseFormula(
                    ctx => $"E{ctx.RowIndex + 3}*0.35"
                ),
                fieldStyleGenerator: _ => currencyStyle
            )
            .Add(
                "Qty",
                static o => o.Qty,
                configureGenerators: static cfg => cfg.UseDataValidation(
                    _ => new DataValidation {
                        ValidationType = DataValidationType.Integer,
                        Operator = DataValidationOperator.GreaterThan,
                        Value1 = 0,
                        ErrorTitle = "Invalid quantity",
                        ErrorMessage = "Quantity must be a positive integer."
                    }
                ),
                fieldStyleGenerator: _ => intStyle
            );

    return template;
}

static ISheetTemplate CreateProductsTemplate(DataTable products) {
    CellStyle titleStyle = new(
        HorizontalAlignment: HorizontalAlignment.Center,
        HasBorder: true,
        BackgroundColor: Color.FromArgb(84, 130, 53),
        Font: new CellFont("Calibri", 12, Color.White, FontStyles.Bold)
    );
    CellStyle priceStyle = SpreadsheetManager.DefaultCellStyles.FieldStyle with {
        HorizontalAlignment = HorizontalAlignment.Right,
        DataFormat = "#,##0.00"
    };
    CellStyle intStyle = SpreadsheetManager.DefaultCellStyles.FieldStyle with {
        HorizontalAlignment = HorizontalAlignment.Right
    };
    CellStyle lowStockStyle = SpreadsheetManager.DefaultCellStyles.FieldStyle with {
        Font = new CellFont(Style: FontStyles.Bold),
        BackgroundColor = Color.FromArgb(255, 235, 156)
    };

    GridTemplate header = new GridTemplate()
        .CreateRow(28)
        .CreateCell("Product Inventory — DataTableTemplate Demo", columnSpan: 5, cellStyle: titleStyle);

    DataTableTemplate dataTemplate = new(products) {
        HeaderHeight = 22,
        RecordHeight = 20
    };
    dataTemplate.Columns.Clear();
    dataTemplate.Columns.Add(new DataTableColumn {
        ColumnName = "ProductId",
        HeaderText = "ID",
        FieldStyleGenerator = _ => intStyle
    });
    dataTemplate.Columns.Add(new DataTableColumn {
        ColumnName = "Name",
        HeaderText = "Product Name"
    });
    dataTemplate.Columns.Add(new DataTableColumn {
        ColumnName = "Category",
        HeaderText = "Category"
    });
    dataTemplate.Columns.Add(new DataTableColumn {
        ColumnName = "UnitPrice",
        HeaderText = "Unit Price",
        FieldStyleGenerator = _ => priceStyle
    });
    dataTemplate.Columns.Add(new DataTableColumn {
        ColumnName = "Stock",
        HeaderText = "Stock",
        FieldStyleGenerator = value => (int?)value < 10 ? lowStockStyle : intStyle
    });

    return new MergedTemplate(header, dataTemplate);
}

static IEnumerable<SampleOrder> GetSampleOrders() => [
    new SampleOrder(1001, "Northwind", "North", 1250.40m, 8),
    new SampleOrder(1002, "Adventure Works", "West", 980.00m, 5),
    new SampleOrder(1003, "Tailspin Toys", "South", 1640.75m, 12),
    new SampleOrder(1004, "Litware", "East", 725.20m, 3),
    new SampleOrder(1005, "Contoso", "North", 2100.00m, 7),
    new SampleOrder(1006, "Fabrikam", "West", 430.50m, 15)
];

static DataTable GetSampleProducts() {
    DataTable table = new();
    table.Columns.Add("ProductId", typeof(int));
    table.Columns.Add("Name", typeof(string));
    table.Columns.Add("Category", typeof(string));
    table.Columns.Add("UnitPrice", typeof(decimal));
    table.Columns.Add("Stock", typeof(int));

    table.Rows.Add(1, "Laptop Pro 15", "Computers", 1299.00m, 24);
    table.Rows.Add(2, "Wireless Mouse", "Accessories", 29.99m, 5);
    table.Rows.Add(3, "Mechanical Keyboard", "Accessories", 89.00m, 8);
    table.Rows.Add(4, "4K Monitor 27\"", "Displays", 549.00m, 3);
    table.Rows.Add(5, "USB-C Hub", "Accessories", 49.99m, 42);
    table.Rows.Add(6, "SSD 1TB", "Storage", 119.00m, 17);
    table.Rows.Add(7, "Webcam HD", "Peripherals", 79.00m, 9);
    table.Rows.Add(8, "Desk Lamp LED", "Furniture", 39.99m, 2);

    return table;
}

// ────────────────────────────────────────────────
// JSON document
// ────────────────────────────────────────────────

static string CreateJsonTemplate() => """
    [
      {
        "SheetName": "Summary",
        "DefaultRowHeight": 20,
        "ColumnWidths": [
          { "Index": 0, "Width": 12 },
          { "Index": 1, "Width": 28 },
          { "Index": 2, "Width": 16 },
          { "Index": 3, "Width": 16 },
          { "Index": 4, "Width": 14 }
        ],
        "PageSettings": {
          "PageOrientation": "Landscape",
          "PaperSize": "A4"
        },
        "FreezePanes": { "Row": 3, "Column": 0 },
        "IsAutoFilterEnabled": true,
        "Templates": [
          {
            "Type": "Grid",
            "Rows": [
              {
                "Height": 28,
                "Cells": [
                  {
                    "Value": "Sales Summary — Generated from SpreadsheetDocument.FromJson",
                    "ColumnSpan": 5,
                    "Style": {
                      "HorizontalAlignment": "Center",
                      "HasBorder": true,
                      "Font": { "Size": 13, "Style": "Bold" }
                    }
                  }
                ]
              },
              {
                "Height": 22,
                "Cells": [
                  {
                    "Value": "Region",
                    "Style": { "HasBorder": true, "Font": { "Style": "Bold" }, "HorizontalAlignment": "Center" }
                  },
                  {
                    "Value": "Description",
                    "ColumnSpan": 2,
                    "Style": { "HasBorder": true, "Font": { "Style": "Bold" } }
                  },
                  {
                    "Value": "Target",
                    "Style": { "HasBorder": true, "Font": { "Style": "Bold" }, "HorizontalAlignment": "Right" }
                  },
                  {
                    "Value": "Achieved",
                    "Style": { "HasBorder": true, "Font": { "Style": "Bold" }, "HorizontalAlignment": "Right" }
                  }
                ]
              },
              {
                "Cells": [
                  { "Value": "North", "Style": { "HasBorder": true, "HorizontalAlignment": "Center" } },
                  { "Value": "Northwind + Contoso combined", "ColumnSpan": 2, "Style": { "HasBorder": true } },
                  { "Value": 3000.00, "Style": { "HasBorder": true, "HorizontalAlignment": "Right", "DataFormat": "#,##0.00" } },
                  { "Value": 3350.40, "Style": { "HasBorder": true, "HorizontalAlignment": "Right", "DataFormat": "#,##0.00" } }
                ]
              },
              {
                "Cells": [
                  { "Value": "West",  "Style": { "HasBorder": true, "HorizontalAlignment": "Center" } },
                  { "Value": "Adventure Works + Fabrikam", "ColumnSpan": 2, "Style": { "HasBorder": true } },
                  { "Value": 1500.00, "Style": { "HasBorder": true, "HorizontalAlignment": "Right", "DataFormat": "#,##0.00" } },
                  { "Value": 1410.50, "Style": { "HasBorder": true, "HorizontalAlignment": "Right", "DataFormat": "#,##0.00" } }
                ]
              }
            ]
          },
          {
            "Type": "RecordSet",
            "HeaderHeight": 22,
            "RecordHeight": 20,
            "Columns": [
              {
                "HeaderText": "Order ID",
                "FieldKey": "OrderId",
                "HeaderStyle": { "HasBorder": true, "HorizontalAlignment": "Center", "Font": { "Style": "Bold" } },
                "FieldStyle":  { "HasBorder": true, "HorizontalAlignment": "Right" }
              },
              {
                "HeaderText": "Customer",
                "FieldKey": "Customer",
                "HeaderStyle": { "HasBorder": true, "HorizontalAlignment": "Center", "Font": { "Style": "Bold" } },
                "FieldStyle":  { "HasBorder": true }
              },
              {
                "HeaderText": "Region",
                "FieldKey": "Region",
                "HeaderStyle": { "HasBorder": true, "HorizontalAlignment": "Center", "Font": { "Style": "Bold" } },
                "FieldStyle":  { "HasBorder": true, "HorizontalAlignment": "Center" }
              },
              {
                "HeaderText": "Amount",
                "FieldKey": "Amount",
                "HeaderStyle": { "HasBorder": true, "HorizontalAlignment": "Center", "Font": { "Style": "Bold" } },
                "FieldStyle":  { "HasBorder": true, "HorizontalAlignment": "Right", "DataFormat": "#,##0.00" }
              },
              {
                "HeaderText": "Qty",
                "FieldKey": "Qty",
                "HeaderStyle": { "HasBorder": true, "HorizontalAlignment": "Center", "Font": { "Style": "Bold" } },
                "FieldStyle":  { "HasBorder": true, "HorizontalAlignment": "Right" }
              }
            ],
            "Records": [
              { "OrderId": 1001, "Customer": "Northwind",       "Region": "North", "Amount": 1250.40, "Qty": 8 },
              { "OrderId": 1002, "Customer": "Adventure Works", "Region": "West",  "Amount":  980.00, "Qty": 5 },
              { "OrderId": 1003, "Customer": "Tailspin Toys",   "Region": "South", "Amount": 1640.75, "Qty": 12 },
              { "OrderId": 1004, "Customer": "Litware",         "Region": "East",  "Amount":  725.20, "Qty": 3 },
              { "OrderId": 1005, "Customer": "Contoso",         "Region": "North", "Amount": 2100.00, "Qty": 7 },
              { "OrderId": 1006, "Customer": "Fabrikam",        "Region": "West",  "Amount":  430.50, "Qty": 15 }
            ]
          }
        ]
      }
    ]
    """;

file sealed record SampleOrder(int OrderId, string Customer, string Region, decimal Amount, int Qty);
