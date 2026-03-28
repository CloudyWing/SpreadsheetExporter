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

    private sealed class StubSheetTemplate(TemplateLayout layout) : ISheetTemplate {
        public int ColumnSpan => layout.Cells.Count == 0 ? 0 : layout.Cells.Max(x => x.Point.X + x.Size.Width);

        public int RowSpan => layout.RowSpan;

        public TemplateLayout GetLayout() => layout;
    }
}
