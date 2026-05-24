using System.Drawing;
using System.Text;
using System.Text.Json;
using ClosedXML.Excel;
using CloudyWing.SpreadsheetExporter.Sample;

namespace CloudyWing.SpreadsheetExporter.Renderer.ClosedXML.Tests;

[TestFixture]
[NonParallelizable]
internal sealed class SampleWorkbookSnapshotTests {
    private static readonly DateTimeOffset FixedGeneratedAt = new(
        2026,
        5,
        14,
        10,
        30,
        0,
        TimeSpan.FromHours(8)
    );

    private static readonly JsonSerializerOptions SnapshotJsonOptions = new() {
        WriteIndented = true
    };

    [SetUp]
    public void SetUp() {
        SpreadsheetManager.DefaultCellStyles = null;
        SpreadsheetManager.SetRenderer(static () => new ExcelRenderer());
    }

    [TestCase("fluent", "sample-fluent.json")]
    [TestCase("json", "sample-json.json")]
    public void ExportSampleWorkbook_ShouldMatchGoldenSnapshot(string sampleKind, string goldenFileName) {
        SpreadsheetDocument document = sampleKind switch {
            "fluent" => SampleWorkbookFactory.CreateFluentDocument(FixedGeneratedAt),
            "json" => SampleWorkbookFactory.CreateJsonDocument(),
            _ => throw new ArgumentOutOfRangeException(nameof(sampleKind), sampleKind, "Unknown sample kind.")
        };

        using (Assert.EnterMultipleScope()) {
            Assert.That(document.GetLayoutDiagnostics(), Is.Empty);
            Assert.That(document.GetRendererCapabilityDiagnostics(), Is.Empty);
        }

        byte[] bytes = document.Export();

        Assert.That(bytes.Length, Is.GreaterThan(0));

        using MemoryStream stream = new(bytes);
        using XLWorkbook workbook = new(stream);
        string actual = NormalizeLineEndings(CreateWorkbookSnapshotJson(workbook));
        string expected = NormalizeLineEndings(ReadGoldenFile(goldenFileName, actual));

        Assert.That(actual, Is.EqualTo(expected));
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
            new BorderSnapshot(
                style.Border.TopBorder.ToString(),
                style.Border.RightBorder.ToString(),
                style.Border.BottomBorder.ToString(),
                style.Border.LeftBorder.ToString()
            ),
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
            validation.Ranges
                .Select(range => range.RangeAddress.ToStringFixed())
                .OrderBy(address => address, StringComparer.Ordinal)
                .ToArray(),
            validation.AllowedValues.ToString(),
            validation.Operator.ToString(),
            validation.InCellDropdown,
            validation.IgnoreBlanks,
            validation.ShowErrorMessage,
            validation.ErrorTitle,
            validation.ErrorMessage,
            validation.ShowInputMessage,
            validation.InputTitle,
            validation.InputMessage,
            GetDataValidationProperty(validation, "MinValue"),
            GetDataValidationProperty(validation, "MaxValue")
        );
    }

    private static string GetDataValidationProperty(IXLDataValidation validation, string propertyName) {
        object? propertyValue = validation.GetType().GetProperty(propertyName)?.GetValue(validation);
        return propertyValue?.ToString() ?? "";
    }

    private static string ReadGoldenFile(string fileName, string actual) {
        string path = GetGoldenFilePath(fileName);
        if (ShouldUpdateGoldenFiles()) {
            Directory.CreateDirectory(Path.GetDirectoryName(path)!);
            string fileContent = NormalizeLineEndings(actual)
                .Replace("\n", Environment.NewLine, StringComparison.Ordinal)
                + Environment.NewLine;
            File.WriteAllText(path, fileContent, new UTF8Encoding(false));
        }

        if (!File.Exists(path)) {
            Assert.Fail($"Golden snapshot file does not exist: {path}{Environment.NewLine}{actual}");
        }

        return File.ReadAllText(path);
    }

    private static string GetGoldenFilePath(string fileName) {
        string? directory = TestContext.CurrentContext.TestDirectory;
        while (directory is not null) {
            string solutionPath = Path.Combine(directory, "SpreadsheetExporter.slnx");
            if (File.Exists(solutionPath)) {
                return Path.Combine(
                    directory,
                    "tests",
                    "SpreadsheetExporter.Renderer.ClosedXML.Tests",
                    "GoldenFiles",
                    fileName
                );
            }

            directory = Directory.GetParent(directory)?.FullName;
        }

        return Path.Combine(TestContext.CurrentContext.TestDirectory, "GoldenFiles", fileName);
    }

    private static bool ShouldUpdateGoldenFiles() {
        string? value = Environment.GetEnvironmentVariable("SPREADSHEET_EXPORTER_UPDATE_GOLDEN");
        return string.Equals(value, "1", StringComparison.Ordinal)
            || string.Equals(value, "true", StringComparison.OrdinalIgnoreCase);
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

    private static bool GetWorksheetBooleanProperty(IXLWorksheet worksheet, string propertyName) {
        object? propertyValue = worksheet.GetType().GetProperty(propertyName)?.GetValue(worksheet);
        return propertyValue is bool value && value;
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
        BorderSnapshot Border,
        FontSnapshot Font
    );

    private sealed record BorderSnapshot(
        string Top,
        string Right,
        string Bottom,
        string Left
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
        IReadOnlyList<string> Ranges,
        string AllowedValues,
        string Operator,
        bool InCellDropdown,
        bool IgnoreBlanks,
        bool ShowErrorMessage,
        string ErrorTitle,
        string ErrorMessage,
        bool ShowInputMessage,
        string InputTitle,
        string InputMessage,
        string MinValue,
        string MaxValue
    );
}
