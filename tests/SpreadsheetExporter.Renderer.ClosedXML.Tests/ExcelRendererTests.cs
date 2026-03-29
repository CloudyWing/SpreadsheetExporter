using System.Drawing;
using ClosedXML.Excel;
using CloudyWing.SpreadsheetExporter.Templates;

namespace CloudyWing.SpreadsheetExporter.Renderer.ClosedXML.Tests;

[TestFixture]
internal sealed class ExcelRendererTests {
    [Test]
    public void ExcelRenderer_ShouldExposeSpreadsheetRendererMetadata() {
        ISpreadsheetRenderer renderer = new ExcelRenderer();

        Assert.Multiple(() => {
            Assert.That(renderer, Is.InstanceOf<ExcelRenderer>());
            Assert.That(renderer.ContentType, Is.EqualTo("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"));
            Assert.That(renderer.FileNameExtension, Is.EqualTo(".xlsx"));
        });
    }

    [Test]
    public void Export_WithStyledSheet_ShouldGenerateExpectedWorkbook() {
        SpreadsheetDocument document = new(new ExcelRenderer()) {
            DefaultFont = new CellFont("Consolas", 11)
        };

        SheetDefinition sheet = document.CreateSheet("Report", 24);
        sheet.Password = "sheet-pass";
        sheet.SetColumnWidth(0, 18);
        sheet.SetColumnWidth(1, Constants.HiddenColumn);
        sheet.PageSettings.PageOrientation = PageOrientation.Landscape;
        sheet.PageSettings.PaperSize = PaperSize.A4;
        sheet.AddTemplate(CreateFeatureTemplate());

        byte[] bytes = document.Export();

        Assert.That(bytes.Length, Is.GreaterThan(0));

        using MemoryStream stream = new(bytes);
        using XLWorkbook workbook = new(stream);
        IXLWorksheet worksheet = workbook.Worksheet("Report");

        Assert.Multiple(() => {
            Assert.That(worksheet.Cell("A1").GetString(), Is.EqualTo("Title"));
            Assert.That(worksheet.MergedRanges, Has.Count.EqualTo(1));
            Assert.That(worksheet.Range("A1:B1").IsMerged(), Is.True);
            Assert.That(worksheet.Cell("A1").Style.Alignment.Horizontal, Is.EqualTo(XLAlignmentHorizontalValues.Center));
            Assert.That(worksheet.Cell("A1").Style.Alignment.Vertical, Is.EqualTo(XLAlignmentVerticalValues.Center));
            Assert.That(
                worksheet.Cell("A1").Style.Fill.BackgroundColor.Color.ToArgb(),
                Is.EqualTo(Color.LightGoldenrodYellow.ToArgb())
            );
            Assert.That(worksheet.Cell("A1").Style.Border.TopBorder, Is.EqualTo(XLBorderStyleValues.Thin));
            Assert.That(worksheet.Cell("A1").Style.Font.Bold, Is.True);
            Assert.That(worksheet.Cell("A1").Style.Protection.Locked, Is.True);
            Assert.That(worksheet.Cell("C2").GetString(), Is.EqualTo("Green"));
            Assert.That(worksheet.Cell("B2").FormulaA1, Is.EqualTo("SUM(A2,1)"));
            Assert.That(worksheet.Column(1).Width, Is.EqualTo(18).Within(0.1));
            Assert.That(worksheet.Column(2).IsHidden, Is.True);
            Assert.That(worksheet.Row(1).Height, Is.EqualTo(32).Within(0.1));
            Assert.That(worksheet.Row(2).IsHidden, Is.True);
            Assert.That(worksheet.PageSetup.PageOrientation, Is.EqualTo(XLPageOrientation.Landscape));
            Assert.That(worksheet.PageSetup.PaperSize, Is.EqualTo(XLPaperSize.A4Paper));
            Assert.That(worksheet.SheetView.SplitColumn, Is.EqualTo(1));
            Assert.That(worksheet.SheetView.SplitRow, Is.EqualTo(1));
            Assert.That(worksheet.AutoFilter.IsEnabled, Is.True);
            Assert.That(worksheet.AutoFilter.Range.RangeAddress.ToStringFixed(), Is.EqualTo("$A$1:$C$2"));
        });

        Assert.That(GetWorksheetBooleanProperty(worksheet, "IsProtected"), Is.True);
    }

    [Test]
    public void Render_WithDataValidation_ShouldPersistValidationSettings() {
        SpreadsheetDocument document = new(new ExcelRenderer());
        SheetDefinition sheet = document.CreateSheet("Validation");
        sheet.AddTemplate(CreateValidationTemplate());

        byte[] bytes = document.Export();

        using MemoryStream stream = new(bytes);
        using XLWorkbook workbook = new(stream);
        IXLCell cell = workbook.Worksheet("Validation").Cell("A1");
        IXLDataValidation validation = cell.GetDataValidation();

        Assert.Multiple(() => {
            Assert.That(cell.HasDataValidation, Is.True);
            Assert.That(validation.AllowedValues, Is.EqualTo(XLAllowedValues.List));
            Assert.That(validation.InCellDropdown, Is.False);
            Assert.That(validation.IgnoreBlanks, Is.False);
            Assert.That(validation.ShowErrorMessage, Is.True);
            Assert.That(validation.ErrorTitle, Is.EqualTo("Invalid option"));
            Assert.That(validation.ErrorMessage, Is.EqualTo("Choose a configured option."));
        });
    }

    private static ISheetTemplate CreateFeatureTemplate() {
        List<Cell> cells = [
            CreateValueCell(
                0,
                0,
                "Title",
                new CellStyle(
                    HorizontalAlignment.Center,
                    VerticalAlignment.Middle,
                    true,
                    true,
                    Color.LightGoldenrodYellow,
                    new CellFont("Arial", 14, Color.DarkBlue, FontStyles.Bold),
                    null,
                    true
                ),
                columnSpan: 2
            ),
            CreateValueCell(
                0,
                1,
                10,
                new CellStyle(
                    HorizontalAlignment.Right,
                    VerticalAlignment.Bottom,
                    false,
                    false,
                    default,
                    default,
                    "0",
                    false
                )
            ),
            CreateFormulaCell(1, 1, "=SUM(A2,1)"),
            CreateValueCell(2, 1, "Green")
        ];

        TemplateLayout context = new(
            cells,
            2,
            new Dictionary<int, double?> {
                [0] = 32,
                [1] = Constants.HiddenRow
            }
        );

        return new TestTemplate(context);
    }

    private static ISheetTemplate CreateValidationTemplate() {
        DataValidation validation = new() {
            ValidationType = DataValidationType.List,
            ListItems = ["Red", "Green", "Blue"],
            IsDropdownShown = false,
            IsBlankAllowed = false,
            ErrorTitle = "Invalid option",
            ErrorMessage = "Choose a configured option.",
            IsErrorAlertShown = true
        };

        TemplateLayout context = new(
            [
                CreateValueCell(0, 0, "Green", dataValidation: validation)
            ],
            1,
            new Dictionary<int, double?>()
        );

        return new TestTemplate(context);
    }

    private static Cell CreateValueCell(
        int column,
        int row,
        object value,
        CellStyle? style = null,
        int columnSpan = 1,
        int rowSpan = 1,
        DataValidation? dataValidation = null
    ) {
        return new Cell {
            Point = new Point(column, row),
            Size = new Size(columnSpan, rowSpan),
            ValueGenerator = (_, _) => value,
            CellStyleGenerator = (_, _) => style ?? CellStyle.Empty,
            DataValidationGenerator = dataValidation is null ? null : (_, _) => dataValidation
        };
    }

    private static Cell CreateFormulaCell(int column, int row, string formula, CellStyle? style = null) {
        return new Cell {
            Point = new Point(column, row),
            FormulaGenerator = (_, _) => formula,
            CellStyleGenerator = (_, _) => style ?? CellStyle.Empty
        };
    }

    private static bool GetWorksheetBooleanProperty(IXLWorksheet worksheet, string propertyName) {
        object? propertyValue = worksheet.GetType().GetProperty(propertyName)?.GetValue(worksheet);
        return propertyValue is bool value && value;
    }

    private sealed class TestTemplate : ISheetTemplate {
        private readonly TemplateLayout context;

        public TestTemplate(TemplateLayout context) {
            this.context = context;
        }

        public int ColumnSpan => context.Cells.Count == 0 ? 0 : context.Cells.Max(x => x.Point.X + x.Size.Width);

        public int RowSpan => context.RowSpan;

        public TemplateLayout GetLayout() {
            return context;
        }
    }
}
