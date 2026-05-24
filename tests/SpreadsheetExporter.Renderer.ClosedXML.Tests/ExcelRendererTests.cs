using System.Drawing;
using System.Text.Json;
using ClosedXML.Excel;
using CloudyWing.SpreadsheetExporter.Templates;

namespace CloudyWing.SpreadsheetExporter.Renderer.ClosedXML.Tests;

[TestFixture]
internal sealed class ExcelRendererTests {
    private static readonly JsonSerializerOptions SnapshotJsonOptions = new() {
        WriteIndented = true
    };

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
    public void Capabilities_ShouldExposeSupportedFeatures() {
        ISpreadsheetRendererWithCapabilities renderer = new ExcelRenderer();

        RendererCapabilities capabilities = renderer.Capabilities;

        using (Assert.EnterMultipleScope()) {
            Assert.That(capabilities.SupportsStyles, Is.True);
            Assert.That(capabilities.SupportsMergedCells, Is.True);
            Assert.That(capabilities.SupportsFormulas, Is.True);
            Assert.That(capabilities.SupportsDataValidation, Is.True);
            Assert.That(capabilities.SupportsFreezePanes, Is.True);
            Assert.That(capabilities.SupportsAutoFilter, Is.True);
            Assert.That(capabilities.SupportsMultipleSheets, Is.True);
            Assert.That(capabilities.SupportsWorksheetProtection, Is.True);
            Assert.That(capabilities.SupportsPageSettings, Is.True);
        }
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
        sheet.SetFreezePanes(1, 1);
        sheet.SetAutoFilterEnabled();
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
    public void Export_WithStyledSheet_ShouldMatchGoldenSnapshot() {
        SpreadsheetDocument document = new(new ExcelRenderer()) {
            DefaultFont = new CellFont("Consolas", 11)
        };

        SheetDefinition sheet = document.CreateSheet("Report", 24);
        sheet.Password = "sheet-pass";
        sheet.SetColumnWidth(0, 18);
        sheet.SetColumnWidth(1, Constants.HiddenColumn);
        sheet.SetFreezePanes(1, 1);
        sheet.SetAutoFilterEnabled();
        sheet.PageSettings.PageOrientation = PageOrientation.Landscape;
        sheet.PageSettings.PaperSize = PaperSize.A4;
        sheet.AddTemplate(CreateFeatureTemplate());

        byte[] bytes = document.Export();

        using MemoryStream stream = new(bytes);
        using XLWorkbook workbook = new(stream);
        string actual = NormalizeLineEndings(CreateWorkbookSnapshotJson(workbook));
        string expected = NormalizeLineEndings(ReadGoldenFile("styled-sheet.json", actual));

        Assert.That(actual, Is.EqualTo(expected));
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

    private static string CreateWorkbookSnapshotJson(XLWorkbook workbook) {
        WorkbookSnapshot snapshot = new(
            CreateFontSnapshot(workbook.Style.Font),
            workbook.Worksheets.Select(CreateWorksheetSnapshot).ToArray()
        );

        return JsonSerializer.Serialize(snapshot, SnapshotJsonOptions);
    }

    private static WorksheetSnapshot CreateWorksheetSnapshot(IXLWorksheet worksheet) {
        IXLRange? usedRange = worksheet.RangeUsed();
        int lastRowNumber = usedRange?.RangeAddress.LastAddress.RowNumber ?? 0;
        int lastColumnNumber = usedRange?.RangeAddress.LastAddress.ColumnNumber ?? 0;

        return new WorksheetSnapshot(
            worksheet.Name,
            Round(worksheet.RowHeight),
            GetWorksheetBooleanProperty(worksheet, "IsProtected"),
            worksheet.PageSetup.PageOrientation.ToString(),
            worksheet.PageSetup.PaperSize.ToString(),
            new FreezePanesSnapshot(
                worksheet.SheetView.SplitRow,
                worksheet.SheetView.SplitColumn
            ),
            worksheet.AutoFilter.IsEnabled
                ? worksheet.AutoFilter.Range.RangeAddress.ToStringFixed()
                : null,
            Enumerable.Range(1, lastColumnNumber).Select(columnNumber => {
                IXLColumn column = worksheet.Column(columnNumber);
                return new ColumnSnapshot(columnNumber, Round(column.Width), column.IsHidden);
            }).ToArray(),
            Enumerable.Range(1, lastRowNumber).Select(rowNumber => {
                IXLRow row = worksheet.Row(rowNumber);
                return new RowSnapshot(rowNumber, Round(row.Height), row.IsHidden);
            }).ToArray(),
            worksheet.MergedRanges
                .Select(range => range.RangeAddress.ToStringFixed())
                .OrderBy(address => address, StringComparer.Ordinal)
                .ToArray(),
            CreateCellSnapshots(worksheet, lastRowNumber, lastColumnNumber)
        );
    }

    private static IReadOnlyList<CellSnapshot> CreateCellSnapshots(
        IXLWorksheet worksheet, int lastRowNumber, int lastColumnNumber
    ) {
        List<CellSnapshot> cells = [];

        for (int rowNumber = 1; rowNumber <= lastRowNumber; rowNumber++) {
            for (int columnNumber = 1; columnNumber <= lastColumnNumber; columnNumber++) {
                IXLCell cell = worksheet.Cell(rowNumber, columnNumber);
                if (cell.IsEmpty() && !cell.HasFormula && !cell.HasDataValidation && !cell.IsMerged()) {
                    continue;
                }

                cells.Add(CreateCellSnapshot(cell));
            }
        }

        return cells.AsReadOnly();
    }

    private static CellSnapshot CreateCellSnapshot(IXLCell cell) {
        return new CellSnapshot(
            cell.Address.ToStringFixed(),
            cell.HasFormula ? null : cell.GetString(),
            cell.HasFormula ? cell.FormulaA1 : null,
            cell.IsMerged()
                ? cell.MergedRange().RangeAddress.ToStringFixed()
                : null,
            CreateStyleSnapshot(cell.Style),
            cell.HasDataValidation
                ? CreateDataValidationSnapshot(cell.GetDataValidation())
                : null
        );
    }

    private static CellStyleSnapshot CreateStyleSnapshot(IXLStyle style) {
        return new CellStyleSnapshot(
            style.Alignment.Horizontal.ToString(),
            style.Alignment.Vertical.ToString(),
            style.Alignment.WrapText,
            style.Protection.Locked,
            ToArgb(style.Fill.BackgroundColor.Color),
            style.NumberFormat.Format,
            style.Border.TopBorder.ToString(),
            CreateFontSnapshot(style.Font)
        );
    }

    private static FontSnapshot CreateFontSnapshot(IXLFont font) {
        return new FontSnapshot(
            font.FontName,
            Round(font.FontSize),
            ToArgb(font.FontColor.Color),
            font.Bold,
            font.Italic,
            font.Underline.ToString(),
            font.Strikethrough
        );
    }

    private static DataValidationSnapshot CreateDataValidationSnapshot(IXLDataValidation validation) {
        return new DataValidationSnapshot(
            validation.AllowedValues.ToString(),
            validation.Operator.ToString(),
            validation.InCellDropdown,
            validation.IgnoreBlanks,
            validation.ShowErrorMessage,
            validation.ErrorTitle,
            validation.ErrorMessage,
            validation.ShowInputMessage,
            validation.InputTitle,
            validation.InputMessage
        );
    }

    private static string ReadGoldenFile(string fileName, string actual) {
        string path = Path.Combine(TestContext.CurrentContext.TestDirectory, "GoldenFiles", fileName);
        if (!File.Exists(path)) {
            Assert.Fail($"Golden snapshot file does not exist: {path}{Environment.NewLine}{actual}");
        }

        return File.ReadAllText(path);
    }

    private static string NormalizeLineEndings(string value) {
        return value.Replace("\r\n", "\n", StringComparison.Ordinal).TrimEnd('\n');
    }

    private static double Round(double value) {
        return Math.Round(value, 2, MidpointRounding.AwayFromZero);
    }

    private static string ToArgb(Color color) {
        return color.IsEmpty
            ? ""
            : $"#{color.ToArgb():X8}";
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
            CreateValueCell(2, 1, "Green", dataValidation: CreateColorListValidation())
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
        TemplateLayout context = new(
            [
                CreateValueCell(0, 0, "Green", dataValidation: CreateColorListValidation())
            ],
            1,
            new Dictionary<int, double?>()
        );

        return new TestTemplate(context);
    }

    private static DataValidation CreateColorListValidation() {
        return new DataValidation {
            ValidationType = DataValidationType.List,
            ListItems = ["Red", "Green", "Blue"],
            IsDropdownShown = false,
            IsBlankAllowed = false,
            ErrorTitle = "Invalid option",
            ErrorMessage = "Choose a configured option.",
            IsErrorAlertShown = true
        };
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

    private sealed record WorkbookSnapshot(
        FontSnapshot DefaultFont,
        IReadOnlyList<WorksheetSnapshot> Worksheets
    );

    private sealed record WorksheetSnapshot(
        string Name,
        double DefaultRowHeight,
        bool IsProtected,
        string PageOrientation,
        string PaperSize,
        FreezePanesSnapshot FreezePanes,
        string? AutoFilterRange,
        IReadOnlyList<ColumnSnapshot> Columns,
        IReadOnlyList<RowSnapshot> Rows,
        IReadOnlyList<string> MergedRanges,
        IReadOnlyList<CellSnapshot> Cells
    );

    private sealed record FreezePanesSnapshot(
        int Row,
        int Column
    );

    private sealed record ColumnSnapshot(
        int Index,
        double Width,
        bool IsHidden
    );

    private sealed record RowSnapshot(
        int Index,
        double Height,
        bool IsHidden
    );

    private sealed record CellSnapshot(
        string Address,
        string? Value,
        string? Formula,
        string? MergedRange,
        CellStyleSnapshot Style,
        DataValidationSnapshot? DataValidation
    );

    private sealed record CellStyleSnapshot(
        string HorizontalAlignment,
        string VerticalAlignment,
        bool WrapText,
        bool IsLocked,
        string FillArgb,
        string NumberFormat,
        string TopBorder,
        FontSnapshot Font
    );

    private sealed record FontSnapshot(
        string Name,
        double Size,
        string ColorArgb,
        bool Bold,
        bool Italic,
        string Underline,
        bool Strikeout
    );

    private sealed record DataValidationSnapshot(
        string AllowedValues,
        string Operator,
        bool InCellDropdown,
        bool IgnoreBlanks,
        bool ShowErrorMessage,
        string ErrorTitle,
        string ErrorMessage,
        bool ShowInputMessage,
        string InputTitle,
        string InputMessage
    );
}
