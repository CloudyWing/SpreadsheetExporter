using System.Drawing;
using CloudyWing.SpreadsheetExporter.Exceptions;
using CloudyWing.SpreadsheetExporter.Templates;

namespace CloudyWing.SpreadsheetExporter.Tests;

[TestFixture]
internal class SpreadsheetDocumentTests {
    private static readonly byte[] ExportedBytes = [0x48, 0x65, 0x6C, 0x6C, 0x6F];

    [Test]
    public void Constructor_RendererIsNull_ShouldThrowArgumentNullException() {
        Assert.That(
            () => new SpreadsheetDocument(null!),
            Throws.TypeOf<ArgumentNullException>().And.Property(nameof(ArgumentNullException.ParamName)).EqualTo("renderer")
        );
    }

    [Test]
    public void ContentMetadata_ShouldExposeRendererValues() {
        SpreadsheetDocument sut = new(new FakeRenderer());

        Assert.Multiple(() => {
            Assert.That(sut.ContentType, Is.EqualTo(FakeRenderer.ExpectedContentType));
            Assert.That(sut.FileNameExtension, Is.EqualTo(FakeRenderer.ExpectedFileNameExtension));
        });
    }

    [Test]
    public void LastSheet_WhenNoSheetsExist_ShouldCreateDefaultSheet() {
        SpreadsheetDocument sut = new(new FakeRenderer());

        SheetDefinition sheet = sut.LastSheet;

        Assert.Multiple(() => {
            Assert.That(sheet, Is.SameAs(sut.GetSheet(0)));
            Assert.That(sheet.SheetName, Is.EqualTo("工作表1"));
        });
    }

    [Test]
    public void CreateSheet_WhenDuplicateNameProvided_ShouldGenerateUniqueName() {
        SpreadsheetDocument sut = new(new FakeRenderer());

        SheetDefinition firstSheet = sut.CreateSheet("Sales");
        SheetDefinition secondSheet = sut.CreateSheet("Sales");

        Assert.Multiple(() => {
            Assert.That(firstSheet.SheetName, Is.EqualTo("Sales"));
            Assert.That(secondSheet.SheetName, Is.EqualTo("Sales(1)"));
        });
    }

    [Test]
    public void GetSheet_WhenIndexIsOutOfRange_ShouldThrowArgumentOutOfRangeException() {
        SpreadsheetDocument sut = new(new FakeRenderer());
        sut.CreateSheet("Sheet1");

        ArgumentOutOfRangeException? exception = Assert.Throws<ArgumentOutOfRangeException>(() => sut.GetSheet(1));

        Assert.Multiple(() => {
            Assert.That(exception!.ParamName, Is.EqualTo("index"));
            Assert.That(exception!.ActualValue, Is.EqualTo(1));
            Assert.That(exception!.Message, Does.Contain("Index must be between 0 and 0."));
        });
    }

    [Test]
    public void TryGetSheet_WhenIndexIsValid_ShouldReturnTrueAndSheet() {
        SpreadsheetDocument sut = new(new FakeRenderer());
        SheetDefinition expected = sut.CreateSheet("Sheet1");

        bool result = sut.TryGetSheet(0, out SheetDefinition? actual);

        Assert.Multiple(() => {
            Assert.That(result, Is.True);
            Assert.That(actual, Is.SameAs(expected));
        });
    }

    [Test]
    public void TryGetSheet_WhenIndexIsInvalid_ShouldReturnFalseAndNull() {
        SpreadsheetDocument sut = new(new FakeRenderer());

        bool result = sut.TryGetSheet(0, out SheetDefinition? actual);

        Assert.Multiple(() => {
            Assert.That(result, Is.False);
            Assert.That(actual, Is.Null);
        });
    }

    [Test]
    public void Export_WhenNoSheetsExist_ShouldThrowSheetDefinitionNotFoundException() {
        SpreadsheetDocument sut = new(new FakeRenderer());

        Assert.That(() => sut.Export(), Throws.TypeOf<SheetDefinitionNotFoundException>());
    }

    [Test]
    public void Export_ShouldPassResolvedLayoutsToRenderer() {
        FakeRenderer renderer = new();
        SpreadsheetDocument sut = new(renderer) {
            DefaultFont = new CellFont("Calibri", 11)
        };

        SheetDefinition sheet = sut.CreateSheet("Orders", 18);
        sheet.Password = "sheet-password";
        sheet.Metadata["source"] = "unit-test";
        sheet.SetColumnWidth(0, 24d);
        sheet.FreezePanes = new Point(0, 1);
        sheet.IsAutoFilterEnabled = true;
        sheet.AddTemplate(new StubSheetTemplate(new TemplateLayout(
            [new Cell {
                Point = new Point(1, 2),
                Size = new Size(1, 1),
                ValueGenerator = (_, _) => "value"
            }],
            3,
            new Dictionary<int, double?> { [0] = 12d, [1] = null, [2] = 16d }
        )));

        byte[] result = sut.Export();
        SheetLayout layout = renderer.RenderedLayouts!.Single();

        Assert.Multiple(() => {
            Assert.That(result, Is.EqualTo(ExportedBytes));
            Assert.That(layout.SheetName, Is.EqualTo("Orders"));
            Assert.That(layout.DefaultRowHeight, Is.EqualTo(18d));
            Assert.That(layout.Password, Is.EqualTo("sheet-password"));
            Assert.That(layout.IsProtected, Is.True);
            Assert.That(layout.DefaultFont, Is.EqualTo(sut.DefaultFont));
            Assert.That(layout.ColumnWidths[0], Is.EqualTo(24d));
            Assert.That(layout.RowHeights[0], Is.EqualTo(12d));
            Assert.That(layout.Cells, Has.Count.EqualTo(1));
            Assert.That(layout.FreezePanes, Is.EqualTo(new Point(0, 1)));
            Assert.That(layout.IsAutoFilterEnabled, Is.True);
            Assert.That(layout.Metadata, Is.SameAs(sheet.Metadata));
            Assert.That(layout.Metadata["source"], Is.EqualTo("unit-test"));
        });
    }

    [Test]
    public void GetLayoutDiagnostics_WithOverlappingCells_ShouldReturnDiagnostic() {
        SpreadsheetDocument sut = new(new FakeRenderer());
        SheetDefinition sheet = sut.CreateSheet("Diagnostics");
        sheet.AddTemplate(new StubSheetTemplate(new TemplateLayout(
            [
                new Cell {
                    Point = new Point(0, 0),
                    Size = new Size(2, 1),
                    ValueGenerator = (_, _) => "A"
                },
                new Cell {
                    Point = new Point(1, 0),
                    Size = new Size(1, 1),
                    ValueGenerator = (_, _) => "B"
                }
            ],
            1,
            new Dictionary<int, double?>()
        )));

        IReadOnlyList<LayoutDiagnostic> diagnostics = sut.GetLayoutDiagnostics();

        using (Assert.EnterMultipleScope()) {
            Assert.That(diagnostics, Has.Count.EqualTo(1));
            Assert.That(diagnostics[0].Code, Is.EqualTo(LayoutDiagnosticCodes.OverlappingCellRange));
            Assert.That(diagnostics[0].Range, Is.EqualTo(new CellRange(1, 0, 1, 1)));
            Assert.That(diagnostics[0].Source?.SheetName, Is.EqualTo("Diagnostics"));
        }
    }

    [Test]
    public void GetLayoutDiagnostics_WithInvalidRangeAndDimension_ShouldReturnDiagnostics() {
        SpreadsheetDocument sut = new(new FakeRenderer());
        SheetDefinition sheet = sut.CreateSheet("Diagnostics");
        sheet.SetColumnWidth(-1, 12d);
        sheet.SetColumnWidth(1, -5d);
        sheet.AddTemplate(new StubSheetTemplate(new TemplateLayout(
            [
                new Cell {
                    Point = new Point(-1, 0),
                    Size = new Size(1, 1),
                    ValueGenerator = (_, _) => "A"
                }
            ],
            1,
            new Dictionary<int, double?> {
                [0] = -5d
            }
        )));

        IReadOnlyList<LayoutDiagnostic> diagnostics = sut.GetLayoutDiagnostics();

        using (Assert.EnterMultipleScope()) {
            Assert.That(diagnostics.Select(x => x.Code), Does.Contain(LayoutDiagnosticCodes.InvalidCellRange));
            Assert.That(diagnostics.Select(x => x.Code).Count(x => x == LayoutDiagnosticCodes.InvalidDimension), Is.EqualTo(3));
        }
    }

    [Test]
    public void GetLayoutDiagnostics_WithValueAndFormula_ShouldReturnDiagnostic() {
        SpreadsheetDocument sut = new(new FakeRenderer());
        SheetDefinition sheet = sut.CreateSheet("Diagnostics");
        sheet.AddTemplate(new StubSheetTemplate(new TemplateLayout(
            [
                new Cell {
                    Point = new Point(0, 0),
                    Size = new Size(1, 1),
                    ValueGenerator = (_, _) => "A",
                    FormulaGenerator = (_, _) => "SUM(A1:A2)"
                }
            ],
            1,
            new Dictionary<int, double?>()
        )));

        IReadOnlyList<LayoutDiagnostic> diagnostics = sut.GetLayoutDiagnostics();

        using (Assert.EnterMultipleScope()) {
            Assert.That(diagnostics, Has.Count.EqualTo(1));
            Assert.That(diagnostics[0].Code, Is.EqualTo(LayoutDiagnosticCodes.ValueAndFormulaConflict));
            Assert.That(diagnostics[0].Range, Is.EqualTo(new CellRange(0, 0, 1, 1)));
        }
    }

    [Test]
    public void GetLayoutDiagnostics_WithValueAndFormula_ShouldNotInvokeGenerators() {
        bool valueGeneratorInvoked = false;
        bool formulaGeneratorInvoked = false;
        SpreadsheetDocument sut = new(new FakeRenderer());
        SheetDefinition sheet = sut.CreateSheet("Diagnostics");
        sheet.AddTemplate(new StubSheetTemplate(new TemplateLayout(
            [
                new Cell {
                    Point = new Point(0, 0),
                    Size = new Size(1, 1),
                    ValueGenerator = (_, _) => {
                        valueGeneratorInvoked = true;
                        return "A";
                    },
                    FormulaGenerator = (_, _) => {
                        formulaGeneratorInvoked = true;
                        return "SUM(A1:A2)";
                    }
                }
            ],
            1,
            new Dictionary<int, double?>()
        )));

        IReadOnlyList<LayoutDiagnostic> diagnostics = sut.GetLayoutDiagnostics();

        using (Assert.EnterMultipleScope()) {
            Assert.That(diagnostics, Has.Count.EqualTo(1));
            Assert.That(valueGeneratorInvoked, Is.False);
            Assert.That(formulaGeneratorInvoked, Is.False);
        }
    }

    [Test]
    public void ValidateLayout_WithDiagnostics_ShouldThrowSpreadsheetLayoutException() {
        SpreadsheetDocument sut = new(new FakeRenderer());
        SheetDefinition sheet = sut.CreateSheet("Diagnostics");
        sheet.AddTemplate(new StubSheetTemplate(new TemplateLayout(
            [
                new Cell {
                    Point = new Point(0, 0),
                    Size = new Size(0, 1),
                    ValueGenerator = (_, _) => "A"
                }
            ],
            1,
            new Dictionary<int, double?>()
        )));

        SpreadsheetLayoutException? exception = Assert.Throws<SpreadsheetLayoutException>(() => sut.ValidateLayout());

        using (Assert.EnterMultipleScope()) {
            Assert.That(exception, Is.Not.Null);
            Assert.That(exception!.Diagnostics, Has.Count.EqualTo(1));
            Assert.That(exception.Diagnostics[0].Code, Is.EqualTo(LayoutDiagnosticCodes.InvalidCellRange));
        }
    }

    [Test]
    public void Export_WithLayoutDiagnostics_ShouldNotRunValidation() {
        FakeRenderer renderer = new();
        SpreadsheetDocument sut = new(renderer);
        SheetDefinition sheet = sut.CreateSheet("Diagnostics");
        sheet.AddTemplate(new StubSheetTemplate(new TemplateLayout(
            [
                new Cell {
                    Point = new Point(0, 0),
                    Size = new Size(0, 1),
                    ValueGenerator = (_, _) => "A"
                }
            ],
            1,
            new Dictionary<int, double?>()
        )));

        byte[] result = sut.Export();

        using (Assert.EnterMultipleScope()) {
            Assert.That(result, Is.EqualTo(ExportedBytes));
            Assert.That(renderer.RenderedLayouts, Has.Count.EqualTo(1));
        }
    }

    [Test]
    public void GetRendererCapabilityDiagnostics_WhenRendererDoesNotExposeCapabilities_ShouldReturnEmpty() {
        SpreadsheetDocument sut = CreateDocumentUsingAllCapabilities(new FakeRenderer());

        IReadOnlyList<LayoutDiagnostic> diagnostics = sut.GetRendererCapabilityDiagnostics();

        Assert.That(diagnostics, Is.Empty);
    }

    [Test]
    public void GetRendererCapabilityDiagnostics_WithUnsupportedUsedCapabilities_ShouldReturnDiagnostics() {
        SpreadsheetDocument sut = CreateDocumentUsingAllCapabilities(
            new CapabilityRenderer(new RendererCapabilities())
        );

        IReadOnlyList<LayoutDiagnostic> diagnostics = sut.GetRendererCapabilityDiagnostics();
        string[] messages = diagnostics.Select(x => x.Message).ToArray();

        using (Assert.EnterMultipleScope()) {
            Assert.That(diagnostics, Has.Count.EqualTo(9));
            Assert.That(
                diagnostics.Select(x => x.Code),
                Is.All.EqualTo(LayoutDiagnosticCodes.UnsupportedRendererCapability)
            );
            Assert.That(messages, Has.Some.Contains(nameof(RendererCapabilities.SupportsMultipleSheets)));
            Assert.That(messages, Has.Some.Contains(nameof(RendererCapabilities.SupportsStyles)));
            Assert.That(messages, Has.Some.Contains(nameof(RendererCapabilities.SupportsMergedCells)));
            Assert.That(messages, Has.Some.Contains(nameof(RendererCapabilities.SupportsFormulas)));
            Assert.That(messages, Has.Some.Contains(nameof(RendererCapabilities.SupportsDataValidation)));
            Assert.That(messages, Has.Some.Contains(nameof(RendererCapabilities.SupportsFreezePanes)));
            Assert.That(messages, Has.Some.Contains(nameof(RendererCapabilities.SupportsAutoFilter)));
            Assert.That(messages, Has.Some.Contains(nameof(RendererCapabilities.SupportsWorksheetProtection)));
            Assert.That(messages, Has.Some.Contains(nameof(RendererCapabilities.SupportsPageSettings)));
        }
    }

    [Test]
    public void GetRendererCapabilityDiagnostics_WithSupportedUsedCapabilities_ShouldReturnEmpty() {
        RendererCapabilities capabilities = new() {
            SupportsStyles = true,
            SupportsMergedCells = true,
            SupportsFormulas = true,
            SupportsDataValidation = true,
            SupportsFreezePanes = true,
            SupportsAutoFilter = true,
            SupportsMultipleSheets = true,
            SupportsWorksheetProtection = true,
            SupportsPageSettings = true
        };
        SpreadsheetDocument sut = CreateDocumentUsingAllCapabilities(new CapabilityRenderer(capabilities));

        IReadOnlyList<LayoutDiagnostic> diagnostics = sut.GetRendererCapabilityDiagnostics();

        Assert.That(diagnostics, Is.Empty);
    }

    [Test]
    public void FromJson_WithFreezePanesAndAutoFilter_ShouldPopulateSheetDefinition() {
        SpreadsheetManager.SetRenderer(() => new FakeRenderer());

        string json = """
            [
              {
                "SheetName": "Data",
                "FreezePanes": { "Row": 2, "Column": 0 },
                "IsAutoFilterEnabled": true,
                "Templates": []
              }
            ]
            """;

        SpreadsheetDocument sut = SpreadsheetDocument.FromJson(json);
        SheetDefinition sheet = sut.GetSheet(0);

        Assert.Multiple(() => {
            Assert.That(sheet.FreezePanes, Is.EqualTo(new Point(0, 2)));
            Assert.That(sheet.IsAutoFilterEnabled, Is.True);
        });
    }

    [Test]
    public void FromJson_WithoutFreezePanesAndAutoFilter_ShouldUseDefaults() {
        SpreadsheetManager.SetRenderer(() => new FakeRenderer());

        string json = """
            [
              {
                "SheetName": "Data",
                "Templates": []
              }
            ]
            """;

        SpreadsheetDocument sut = SpreadsheetDocument.FromJson(json);
        SheetDefinition sheet = sut.GetSheet(0);

        Assert.Multiple(() => {
            Assert.That(sheet.FreezePanes, Is.Null);
            Assert.That(sheet.IsAutoFilterEnabled, Is.False);
        });
    }

    [Test]
    public void FromJson_WithMissingTemplateType_ShouldThrowPathDiagnostic() {
        SpreadsheetManager.SetRenderer(() => new FakeRenderer());

        string json = """
            [
              {
                "SheetName": "Data",
                "Templates": [
                  {
                    "Rows": []
                  }
                ]
              }
            ]
            """;

        FormatException? exception = Assert.Throws<FormatException>(() => SpreadsheetDocument.FromJson(json));

        using (Assert.EnterMultipleScope()) {
            Assert.That(exception!.Message, Does.Contain(JsonDiagnosticCodes.MissingRequiredProperty));
            Assert.That(exception.Message, Does.Contain("$[0].Templates[0].Type"));
        }
    }

    [Test]
    public void FromJson_WithInvalidNestedGridCellType_ShouldThrowPathDiagnostic() {
        SpreadsheetManager.SetRenderer(() => new FakeRenderer());

        string json = """
            [
              {
                "SheetName": "Data",
                "Templates": [
                  {
                    "Type": "Grid",
                    "Rows": [
                      {
                        "Cells": [
                          { "Value": "A", "ColumnSpan": "abc" }
                        ]
                      }
                    ]
                  }
                ]
              }
            ]
            """;

        FormatException? exception = Assert.Throws<FormatException>(() => SpreadsheetDocument.FromJson(json));

        using (Assert.EnterMultipleScope()) {
            Assert.That(exception!.Message, Does.Contain(JsonDiagnosticCodes.InvalidType));
            Assert.That(exception.Message, Does.Contain("$[0].Templates[0].Rows[0].Cells[0].ColumnSpan"));
        }
    }

    [Test]
    public void FromJson_WithGridCellValueAndFormula_ShouldThrowPathDiagnostic() {
        SpreadsheetManager.SetRenderer(() => new FakeRenderer());

        string json = """
            [
              {
                "SheetName": "Data",
                "Templates": [
                  {
                    "Type": "Grid",
                    "Rows": [
                      {
                        "Cells": [
                          { "Value": "A", "Formula": "SUM(A1:A2)" }
                        ]
                      }
                    ]
                  }
                ]
              }
            ]
            """;

        InvalidOperationException? exception = Assert.Throws<InvalidOperationException>(
            () => SpreadsheetDocument.FromJson(json)
        );

        using (Assert.EnterMultipleScope()) {
            Assert.That(exception!.Message, Does.Contain(JsonDiagnosticCodes.InvalidValue));
            Assert.That(exception.Message, Does.Contain("$[0].Templates[0].Rows[0].Cells[0]"));
        }
    }

    [Test]
    public void FromJson_WithNamedStyles_ShouldApplyStyleNameAndInlineOverrides() {
        SpreadsheetManager.SetRenderer(() => new FakeRenderer());
        SpreadsheetStyleRegistry documentStyles = new();
        documentStyles.Set(
            "DocumentHeader",
            new CellStyle(
                HasBorder: true,
                Font: new CellFont("Document Font", 12, Style: FontStyles.Bold)
            )
        );

        string json = """
            [
              {
                "SheetName": "Data",
                "Styles": {
                  "SheetAmount": {
                    "Font": {
                      "Name": "Sheet Font",
                      "Size": 14
                    },
                    "DataFormat": "#,##0.00"
                  }
                },
                "Templates": [
                  {
                    "Type": "Grid",
                    "Rows": [
                      {
                        "Cells": [
                          {
                            "Value": "Header",
                            "StyleName": "DocumentHeader",
                            "Style": {
                              "Font": { "Style": "Italic" },
                              "WrapText": true
                            }
                          },
                          {
                            "Value": 1250.40,
                            "StyleName": "SheetAmount"
                          }
                        ]
                      }
                    ]
                  }
                ]
              }
            ]
            """;

        SpreadsheetDocument sut = SpreadsheetDocument.FromJson(json, documentStyles);
        SheetLayout layout = new(sut.GetSheet(0));
        Cell headerCell = layout.Cells.Single(x => x.Point == new Point(0, 0));
        Cell amountCell = layout.Cells.Single(x => x.Point == new Point(1, 0));
        CellStyle headerStyle = headerCell.GetCellStyle();
        CellStyle amountStyle = amountCell.GetCellStyle();

        using (Assert.EnterMultipleScope()) {
            Assert.That(headerStyle.HasBorder, Is.True);
            Assert.That(headerStyle.WrapText, Is.True);
            Assert.That(headerStyle.Font.Name, Is.EqualTo("Document Font"));
            Assert.That(headerStyle.Font.Size, Is.EqualTo(12));
            Assert.That(headerStyle.Font.Style, Is.EqualTo(FontStyles.Italic));
            Assert.That(amountStyle.Font.Name, Is.EqualTo("Sheet Font"));
            Assert.That(amountStyle.Font.Size, Is.EqualTo(14));
            Assert.That(amountStyle.DataFormat, Is.EqualTo("#,##0.00"));
        }
    }

    [Test]
    public void FromJson_WithNamedStyleAndNullFont_ShouldClearInheritedFont() {
        SpreadsheetManager.SetRenderer(() => new FakeRenderer());
        SpreadsheetStyleRegistry documentStyles = new();
        documentStyles.Set("Header", new CellStyle(Font: new CellFont("Document Font", 12, Style: FontStyles.Bold)));

        string json = """
            [
              {
                "SheetName": "Data",
                "Templates": [
                  {
                    "Type": "Grid",
                    "Rows": [
                      {
                        "Cells": [
                          {
                            "Value": "Header",
                            "StyleName": "Header",
                            "Style": {
                              "Font": null
                            }
                          }
                        ]
                      }
                    ]
                  }
                ]
              }
            ]
            """;

        SpreadsheetDocument sut = SpreadsheetDocument.FromJson(json, documentStyles);
        SheetLayout layout = new(sut.GetSheet(0));
        CellStyle actual = layout.Cells.Single(x => x.Point == new Point(0, 0)).GetCellStyle();

        Assert.That(actual.Font, Is.EqualTo(CellFont.Empty));
    }

    [Test]
    public void FromJson_WithUnknownStyleName_ShouldThrowPathDiagnostic() {
        SpreadsheetManager.SetRenderer(() => new FakeRenderer());

        string json = """
            [
              {
                "SheetName": "Data",
                "Templates": [
                  {
                    "Type": "Grid",
                    "Rows": [
                      {
                        "Cells": [
                          {
                            "Value": "Header",
                            "StyleName": "Missing"
                          }
                        ]
                      }
                    ]
                  }
                ]
              }
            ]
            """;

        FormatException? exception = Assert.Throws<FormatException>(() => SpreadsheetDocument.FromJson(json));

        using (Assert.EnterMultipleScope()) {
            Assert.That(exception!.Message, Does.Contain(JsonDiagnosticCodes.InvalidValue));
            Assert.That(exception.Message, Does.Contain("$[0].Templates[0].Rows[0].Cells[0].StyleName"));
        }
    }

    [Test]
    public void FromJson_WithGridDataValidation_ShouldPopulateCellValidation() {
        SpreadsheetManager.SetRenderer(() => new FakeRenderer());

        string json = """
            [
              {
                "SheetName": "Data",
                "Templates": [
                  {
                    "Type": "Grid",
                    "Rows": [
                      {
                        "Cells": [
                          {
                            "Value": "Status",
                            "DataValidation": {
                              "ValidationType": "List",
                              "ListItems": ["Pending", "Paid"],
                              "IsBlankAllowed": false,
                              "ErrorTitle": "Invalid status",
                              "ErrorMessage": "Choose a listed status.",
                              "IsInputPromptShown": true,
                              "PromptTitle": "Status",
                              "PromptMessage": "Choose a listed status."
                            }
                          }
                        ]
                      }
                    ]
                  }
                ]
              }
            ]
            """;

        SpreadsheetDocument sut = SpreadsheetDocument.FromJson(json);
        SheetLayout layout = new(sut.GetSheet(0));
        DataValidation? validation = layout.Cells.Single().GetDataValidation();

        using (Assert.EnterMultipleScope()) {
            Assert.That(validation, Is.Not.Null);
            Assert.That(validation!.ValidationType, Is.EqualTo(DataValidationType.List));
            Assert.That(validation.ListItems, Is.EqualTo(new[] { "Pending", "Paid" }));
            Assert.That(validation.IsBlankAllowed, Is.False);
            Assert.That(validation.ErrorTitle, Is.EqualTo("Invalid status"));
            Assert.That(validation.ErrorMessage, Is.EqualTo("Choose a listed status."));
            Assert.That(validation.IsInputPromptShown, Is.True);
            Assert.That(validation.PromptTitle, Is.EqualTo("Status"));
            Assert.That(validation.PromptMessage, Is.EqualTo("Choose a listed status."));
        }
    }

    [Test]
    public void FromJson_WithRecordSetDataValidation_ShouldPopulateFieldValidation() {
        SpreadsheetManager.SetRenderer(() => new FakeRenderer());

        string json = """
            [
              {
                "SheetName": "Data",
                "Templates": [
                  {
                    "Type": "RecordSet",
                    "Columns": [
                      {
                        "HeaderText": "Quantity",
                        "FieldKey": "Quantity",
                        "DataValidation": {
                          "ValidationType": "Integer",
                          "Operator": "Between",
                          "Value1": 1,
                          "Value2": 10
                        }
                      }
                    ],
                    "Records": [
                      { "Quantity": 2 }
                    ]
                  }
                ]
              }
            ]
            """;

        SpreadsheetDocument sut = SpreadsheetDocument.FromJson(json);
        SheetLayout layout = new(sut.GetSheet(0));
        Cell fieldCell = layout.Cells.Single(x => x.Point == new Point(0, 1));
        DataValidation? validation = fieldCell.GetDataValidation();

        using (Assert.EnterMultipleScope()) {
            Assert.That(validation, Is.Not.Null);
            Assert.That(validation!.ValidationType, Is.EqualTo(DataValidationType.Integer));
            Assert.That(validation.Operator, Is.EqualTo(DataValidationOperator.Between));
            Assert.That(validation.Value1, Is.EqualTo(1));
            Assert.That(validation.Value2, Is.EqualTo(10));
        }
    }

    [Test]
    public void FromJson_WithDataTableTemplate_ShouldPopulateTemplateLayout() {
        SpreadsheetManager.SetRenderer(() => new FakeRenderer());

        string json = """
            [
              {
                "SheetName": "Data",
                "Templates": [
                  {
                    "Type": "DataTable",
                    "HeaderHeight": 22,
                    "RecordHeight": 18,
                    "Columns": [
                      {
                        "ColumnName": "OrderId",
                        "HeaderText": "Order ID",
                        "HeaderStyle": {
                          "Font": { "Style": "Bold" }
                        }
                      },
                      {
                        "ColumnName": "Amount",
                        "FieldStyle": {
                          "DataFormat": "#,##0.00"
                        }
                      }
                    ],
                    "Records": [
                      { "OrderId": 1001, "Amount": 1250.40 },
                      { "OrderId": 1002 }
                    ]
                  }
                ]
              }
            ]
            """;

        SpreadsheetDocument sut = SpreadsheetDocument.FromJson(json);
        SheetLayout layout = new(sut.GetSheet(0));
        Cell orderIdHeaderCell = layout.Cells.Single(x => x.Point == new Point(0, 0));
        Cell amountHeaderCell = layout.Cells.Single(x => x.Point == new Point(1, 0));
        Cell amountCell = layout.Cells.Single(x => x.Point == new Point(1, 1));
        Cell missingAmountCell = layout.Cells.Single(x => x.Point == new Point(1, 2));

        using (Assert.EnterMultipleScope()) {
            Assert.That(layout.Cells, Has.Count.EqualTo(6));
            Assert.That(layout.RowHeights[0], Is.EqualTo(22));
            Assert.That(layout.RowHeights[1], Is.EqualTo(18));
            Assert.That(orderIdHeaderCell.GetValue(), Is.EqualTo("Order ID"));
            Assert.That(orderIdHeaderCell.GetCellStyle().Font.Style, Is.EqualTo(FontStyles.Bold));
            Assert.That(amountHeaderCell.GetValue(), Is.EqualTo("Amount"));
            Assert.That(amountCell.GetValue(), Is.EqualTo(1250.40m));
            Assert.That(amountCell.GetCellStyle().DataFormat, Is.EqualTo("#,##0.00"));
            Assert.That(missingAmountCell.GetValue(), Is.Null);
        }
    }

    [Test]
    public void FromJson_WithDataTableTemplateAndNoColumns_ShouldInferColumnsFromRecords() {
        SpreadsheetManager.SetRenderer(() => new FakeRenderer());

        string json = """
            [
              {
                "SheetName": "Data",
                "Templates": [
                  {
                    "Type": "DataTable",
                    "Records": [
                      { "Id": 1, "Name": "Alice" },
                      { "Name": "Bob", "Amount": 20 }
                    ]
                  }
                ]
              }
            ]
            """;

        SpreadsheetDocument sut = SpreadsheetDocument.FromJson(json);
        SheetLayout layout = new(sut.GetSheet(0));
        List<Cell> headerCells = layout.Cells
            .Where(x => x.Point.Y == 0)
            .OrderBy(x => x.Point.X)
            .ToList();
        Cell missingIdCell = layout.Cells.Single(x => x.Point == new Point(0, 2));
        Cell amountCell = layout.Cells.Single(x => x.Point == new Point(2, 2));

        using (Assert.EnterMultipleScope()) {
            Assert.That(headerCells.Select(x => x.GetValue()), Is.EqualTo(new[] { "Id", "Name", "Amount" }));
            Assert.That(missingIdCell.GetValue(), Is.Null);
            Assert.That(amountCell.GetValue(), Is.EqualTo(20));
        }
    }

    [Test]
    public void FromJson_WithMergedTemplate_ShouldStackChildTemplates() {
        SpreadsheetManager.SetRenderer(() => new FakeRenderer());

        string json = """
            [
              {
                "SheetName": "Data",
                "Templates": [
                  {
                    "Type": "Merged",
                    "Templates": [
                      {
                        "Type": "Grid",
                        "Rows": [
                          {
                            "Cells": [
                              { "Value": "Summary" }
                            ]
                          }
                        ]
                      },
                      {
                        "Type": "DataTable",
                        "Columns": [
                          { "ColumnName": "Name" }
                        ],
                        "Records": [
                          { "Name": "Alice" }
                        ]
                      }
                    ]
                  }
                ]
              }
            ]
            """;

        SpreadsheetDocument sut = SpreadsheetDocument.FromJson(json);
        SheetLayout layout = new(sut.GetSheet(0));
        Cell summaryCell = layout.Cells.Single(x => x.Point == new Point(0, 0));
        Cell nameHeaderCell = layout.Cells.Single(x => x.Point == new Point(0, 1));
        Cell nameCell = layout.Cells.Single(x => x.Point == new Point(0, 2));

        using (Assert.EnterMultipleScope()) {
            Assert.That(layout.Cells, Has.Count.EqualTo(3));
            Assert.That(summaryCell.GetValue(), Is.EqualTo("Summary"));
            Assert.That(nameHeaderCell.GetValue(), Is.EqualTo("Name"));
            Assert.That(nameCell.GetValue(), Is.EqualTo("Alice"));
        }
    }

    [Test]
    public void ExportFile_ShouldWriteExportedBytes() {
        FakeRenderer renderer = new();
        SpreadsheetDocument sut = new(renderer);
        sut.CreateSheet("Sheet1").AddTemplate(new StubSheetTemplate(new TemplateLayout([], 0, new Dictionary<int, double?>())));

        string path = Path.Combine(TestContext.CurrentContext.WorkDirectory, $"{Guid.NewGuid():N}.bin");

        try {
            sut.ExportFile(path);

            Assert.That(File.ReadAllBytes(path), Is.EqualTo(ExportedBytes));
        } finally {
            if (File.Exists(path)) {
                File.Delete(path);
            }
        }
    }

    [Test]
    public void ExportFile_WhenCreateNewAndFileExists_ShouldThrowIOException() {
        SpreadsheetDocument sut = new(new FakeRenderer());
        sut.CreateSheet("Sheet1").AddTemplate(new StubSheetTemplate(new TemplateLayout([], 0, new Dictionary<int, double?>())));

        string path = Path.Combine(TestContext.CurrentContext.WorkDirectory, $"{Guid.NewGuid():N}.bin");

        try {
            File.WriteAllBytes(path, [0x01]);

            Assert.That(
                () => sut.ExportFile(path, SpreadsheetFileMode.CreateNew),
                Throws.TypeOf<IOException>()
            );
        } finally {
            if (File.Exists(path)) {
                File.Delete(path);
            }
        }
    }

    private sealed class FakeRenderer : ISpreadsheetRenderer {
        public const string ExpectedContentType = "application/test";
        public const string ExpectedFileNameExtension = ".test";

        public string ContentType => ExpectedContentType;

        public string FileNameExtension => ExpectedFileNameExtension;

        public IReadOnlyList<SheetLayout> RenderedLayouts { get; private set; } = [];

        public byte[] Render(IEnumerable<SheetLayout> layouts) {
            RenderedLayouts = layouts.ToList();
            return ExportedBytes;
        }
    }

    private sealed class CapabilityRenderer : ISpreadsheetRendererWithCapabilities {
        public CapabilityRenderer(RendererCapabilities capabilities) {
            Capabilities = capabilities;
        }

        public string ContentType => FakeRenderer.ExpectedContentType;

        public string FileNameExtension => FakeRenderer.ExpectedFileNameExtension;

        public RendererCapabilities Capabilities { get; }

        public byte[] Render(IEnumerable<SheetLayout> layouts) {
            return ExportedBytes;
        }
    }

    private sealed class StubSheetTemplate(TemplateLayout layout) : ISheetTemplate {
        public int ColumnSpan => layout.Cells.Count == 0 ? 0 : layout.Cells.Max(x => x.Point.X + x.Size.Width);

        public int RowSpan => layout.RowSpan;

        public TemplateLayout GetLayout() => layout;
    }

    private static SpreadsheetDocument CreateDocumentUsingAllCapabilities(ISpreadsheetRenderer renderer) {
        SpreadsheetDocument document = new(renderer) {
            DefaultFont = new CellFont("Calibri", 11)
        };

        SheetDefinition sheet = document.CreateSheet("Capabilities");
        sheet.SetFreezePanes(1, 1);
        sheet.SetAutoFilterEnabled();
        sheet.SetPassword("password");
        sheet.ConfigurePageSettings(settings => settings.PageOrientation = PageOrientation.Landscape);
        sheet.AddTemplate(new StubSheetTemplate(new TemplateLayout(
            [
                new Cell {
                    Point = new Point(0, 0),
                    Size = new Size(2, 1),
                    FormulaGenerator = (_, _) => "SUM(A1:A2)",
                    DataValidationGenerator = (_, _) => new DataValidation(),
                    CellStyleGenerator = (_, _) => CellStyle.Empty
                }
            ],
            1,
            new Dictionary<int, double?>()
        )));
        document.CreateSheet("Second");

        return document;
    }
}
