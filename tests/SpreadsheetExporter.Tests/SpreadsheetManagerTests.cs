using System.Reflection;
using CloudyWing.SpreadsheetExporter.Config;

namespace CloudyWing.SpreadsheetExporter.Tests;

[TestFixture]
internal class SpreadsheetManagerTests {
    [SetUp]
    public void SetUp() {
        SpreadsheetManager.DefaultCellStyles = null;
        ResetRendererFactory();
    }

    [Test]
    public void DefaultCellStyles_WhenNotOverridden_ShouldReturnBuiltInConfiguration() {
        CellStyleConfiguration config = SpreadsheetManager.DefaultCellStyles;

        Assert.Multiple(() => {
            Assert.That(config, Is.Not.Null);
            Assert.That(config.CellStyle.Font.Name, Is.EqualTo("新細明體"));
            Assert.That(config.CellStyle.Font.Size, Is.EqualTo(10));
            Assert.That(config.HeaderStyle.Font.Style, Is.EqualTo(FontStyles.Bold));
        });
    }

    [Test]
    public void DefaultCellStyles_WhenSet_ShouldReturnAssignedConfiguration() {
        CellStyleConfiguration custom = new() {
            CellStyle = new CellStyle(Font: new CellFont("Custom", 14)),
            GridCellStyle = new CellStyle(Font: new CellFont("Grid", 12)),
            HeaderStyle = new CellStyle(Font: new CellFont("Header", 13, Style: FontStyles.Bold)),
            FieldStyle = new CellStyle(Font: new CellFont("Field", 11))
        };

        SpreadsheetManager.DefaultCellStyles = custom;

        Assert.That(SpreadsheetManager.DefaultCellStyles, Is.SameAs(custom));
    }

    [Test]
    public void DefaultCellStyles_WhenResetToNull_ShouldReturnBuiltInConfigurationAgain() {
        SpreadsheetManager.DefaultCellStyles = new CellStyleConfiguration {
            CellStyle = new CellStyle(Font: new CellFont("Custom", 14)),
            GridCellStyle = new CellStyle(Font: new CellFont("Custom", 14)),
            HeaderStyle = new CellStyle(Font: new CellFont("Custom", 14)),
            FieldStyle = new CellStyle(Font: new CellFont("Custom", 14))
        };

        SpreadsheetManager.DefaultCellStyles = null;

        Assert.That(SpreadsheetManager.DefaultCellStyles.CellStyle.Font.Name, Is.EqualTo("新細明體"));
    }

    [Test]
    public void SetRenderer_WithNullFactory_ShouldThrowArgumentNullException() {
        ArgumentNullException exception = Assert.Throws<ArgumentNullException>(() => SpreadsheetManager.SetRenderer(null!))!;

        Assert.That(exception.ParamName, Is.EqualTo("rendererFactory"));
    }

    [Test]
    public void SetRenderer_WithFactoryReturningNull_ShouldThrowArgumentException() {
        ArgumentException exception = Assert.Throws<ArgumentException>(() => SpreadsheetManager.SetRenderer(() => null!))!;

        Assert.Multiple(() => {
            Assert.That(exception.ParamName, Is.EqualTo("rendererFactory"));
            Assert.That(exception.Message, Does.Contain("Factory return value cannot be null."));
        });
    }

    [Test]
    public void CreateDocument_WhenRendererFactoryNotSet_ShouldThrowInvalidOperationException() {
        InvalidOperationException exception = Assert.Throws<InvalidOperationException>(() => SpreadsheetManager.CreateDocument())!;

        Assert.That(exception.Message, Does.Contain("Renderer factory is not set."));
    }

    [Test]
    public void CreateDocument_WhenRendererFactorySet_ShouldReturnNewDocumentWithRendererMetadata() {
        SpreadsheetManager.SetRenderer(() => new FakeRenderer("application/test-1", ".one"));

        SpreadsheetDocument document1 = SpreadsheetManager.CreateDocument();
        SpreadsheetDocument document2 = SpreadsheetManager.CreateDocument();

        Assert.Multiple(() => {
            Assert.That(document1, Is.Not.SameAs(document2));
            Assert.That(document1.ContentType, Is.EqualTo("application/test-1"));
            Assert.That(document1.FileNameExtension, Is.EqualTo(".one"));
            Assert.That(document2.ContentType, Is.EqualTo("application/test-1"));
        });
    }

    [Test]
    public void SetRenderer_CalledMultipleTimes_ShouldUseLatestFactory() {
        SpreadsheetManager.SetRenderer(() => new FakeRenderer("application/old", ".old"));
        SpreadsheetDocument oldDocument = SpreadsheetManager.CreateDocument();

        SpreadsheetManager.SetRenderer(() => new FakeRenderer("application/new", ".new"));
        SpreadsheetDocument newDocument = SpreadsheetManager.CreateDocument();

        Assert.Multiple(() => {
            Assert.That(oldDocument.ContentType, Is.EqualTo("application/old"));
            Assert.That(newDocument.ContentType, Is.EqualTo("application/new"));
            Assert.That(newDocument.FileNameExtension, Is.EqualTo(".new"));
        });
    }

    private static void ResetRendererFactory() {
        FieldInfo field = typeof(SpreadsheetManager).GetField("rendererFactory", BindingFlags.NonPublic | BindingFlags.Static)!;
        field.SetValue(null, null);
    }

    private sealed class FakeRenderer(string contentType, string fileNameExtension) : ISpreadsheetRenderer {
        public string ContentType { get; } = contentType;

        public string FileNameExtension { get; } = fileNameExtension;

        public byte[] Render(IEnumerable<SheetLayout> contexts) => [];
    }
}
