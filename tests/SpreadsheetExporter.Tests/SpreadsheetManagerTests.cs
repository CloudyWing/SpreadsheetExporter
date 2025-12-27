using CloudyWing.SpreadsheetExporter.Config;

namespace CloudyWing.SpreadsheetExporter.Tests;

[TestFixture]
internal class SpreadsheetManagerTests {
    [SetUp]
    public void SetUp() {
        // 重置靜態狀態
        SpreadsheetManager.DefaultCellStyles = null;
    }

    [Test]
    public void DefaultCellStyles_WhenNotSet_ShouldReturnDefaultConfiguration() {
        CellStyleConfiguration config = SpreadsheetManager.DefaultCellStyles;

        Assert.That(config, Is.Not.Null);
        Assert.That(config.CellStyle, Is.Not.Null);
        Assert.That(config.CellStyle.Font.Name, Is.EqualTo("新細明體"));
        Assert.That(config.CellStyle.Font.Size, Is.EqualTo(10));
    }

    [Test]
    public void DefaultCellStyles_WhenSet_ShouldReturnUserConfiguration() {
        CellStyle customCellStyle = new(
            HorizontalAlignment.Left,
            VerticalAlignment.Top,
            false, false,
            default,
            new CellFont("Custom Font", 14, default, FontStyles.None),
            null,
            false
        );

        CellStyleConfiguration userConfig = new(setuper => {
            setuper.CellStyle = customCellStyle;
        });

        SpreadsheetManager.DefaultCellStyles = userConfig;

        CellStyleConfiguration result = SpreadsheetManager.DefaultCellStyles;
        Assert.That(result, Is.SameAs(userConfig));
        Assert.That(result.CellStyle.Font.Name, Is.EqualTo("Custom Font"));
        Assert.That(result.CellStyle.Font.Size, Is.EqualTo(14));
    }

    [Test]
    public void DefaultCellStyles_WhenSetToNull_ShouldReturnDefaultConfiguration() {
        CellStyleConfiguration userConfig = new();
        SpreadsheetManager.DefaultCellStyles = userConfig;

        SpreadsheetManager.DefaultCellStyles = null;

        CellStyleConfiguration result = SpreadsheetManager.DefaultCellStyles;
        Assert.That(result, Is.Not.Null);
        Assert.That(result.CellStyle.Font.Name, Is.EqualTo("新細明體"));
    }

    [Test]
    public void SetExporter_WithValidFactory_ShouldSetFactory() {
        FakeExporter exporter = new();
        ISpreadsheetExporter factory() => exporter;

        Assert.DoesNotThrow(() => SpreadsheetManager.SetExporter(factory));

        ISpreadsheetExporter result = SpreadsheetManager.CreateExporter();
        Assert.That(result, Is.SameAs(exporter));
    }

    [Test]
    public void SetExporter_WithNullFactory_ShouldThrowArgumentNullException() {
        ArgumentNullException ex = Assert.Throws<ArgumentNullException>(
            () => SpreadsheetManager.SetExporter(null)
        );

        Assert.That(ex.ParamName, Is.EqualTo("exporterFactory"));
    }

    [Test]
    public void SetExporter_WithFactoryReturningNull_ShouldThrowArgumentException() {
        ISpreadsheetExporter factory() => null;

        ArgumentException ex = Assert.Throws<ArgumentException>(
            () => SpreadsheetManager.SetExporter(factory)
        );

        Assert.That(ex.ParamName, Is.EqualTo("exporterFactory"));
        Assert.That(ex.Message, Does.Contain("Factory return value cannot be null"));
    }

    [Test]
    public void CreateExporter_WhenFactoryNotSet_ShouldThrowInvalidOperationException() {
        InvalidOperationException ex = Assert.Throws<InvalidOperationException>(
            () => SpreadsheetManager.CreateExporter()
        );

        Assert.That(ex.Message, Does.Contain("Exporter factory is not set"));
    }

    [Test]
    public void CreateExporter_WhenFactorySet_ShouldReturnNewExporterInstance() {
        SpreadsheetManager.SetExporter(() => new FakeExporter());

        ISpreadsheetExporter exporter1 = SpreadsheetManager.CreateExporter();
        ISpreadsheetExporter exporter2 = SpreadsheetManager.CreateExporter();

        Assert.That(exporter1, Is.Not.Null);
        Assert.That(exporter2, Is.Not.Null);
        Assert.That(exporter1, Is.Not.SameAs(exporter2));
    }

    [Test]
    public void SetExporter_CalledMultipleTimes_ShouldUpdateFactory() {
        FakeExporter exporter1 = new();
        FakeExporter exporter2 = new();

        SpreadsheetManager.SetExporter(() => exporter1);
        ISpreadsheetExporter result1 = SpreadsheetManager.CreateExporter();

        SpreadsheetManager.SetExporter(() => exporter2);
        ISpreadsheetExporter result2 = SpreadsheetManager.CreateExporter();

        Assert.That(result1, Is.SameAs(exporter1));
        Assert.That(result2, Is.SameAs(exporter2));
        Assert.That(result1, Is.Not.SameAs(result2));
    }

    private class FakeExporter : ISpreadsheetExporter {
        public event EventHandler<SpreadsheetExportingEventArgs> SpreadsheetExportingEvent;

        public event EventHandler<SpreadsheetExportedEventArgs> SpreadsheetExportedEvent;

        public event EventHandler<SheetCreatedEventArgs> SheetCreatedEvent;

        public string ContentType => "application/test";

        public string FileNameExtension => ".test";

        public string Password { get; set; } = "";

        public bool HasPassword => !string.IsNullOrEmpty(Password);

        public CellFont? DefaultFont { get; set; }

        public string DefaultBasicSheetName { get; set; }

        public Sheeter LastSheeter => throw new NotImplementedException();

        public Sheeter CreateSheeter(string sheetName = null, double? defaultRowHeight = null) {
            throw new NotImplementedException();
        }

        public Sheeter GetSheeter(int index) {
            throw new NotImplementedException();
        }

        public bool TryGetSheeter(int index, out Sheeter sheeter) {
            throw new NotImplementedException();
        }

        public byte[] Export() {
            return [];
        }

        public void ExportFile(string path, SpreadsheetFileMode fileMode = SpreadsheetFileMode.Create) {
            throw new NotImplementedException();
        }
    }
}
