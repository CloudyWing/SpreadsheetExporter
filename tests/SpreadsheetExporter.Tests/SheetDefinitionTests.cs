using System.Drawing;
using CloudyWing.SpreadsheetExporter.Templates;

namespace CloudyWing.SpreadsheetExporter.Tests;

[TestFixture]
internal class SheetDefinitionTests {
    [Test]
    public void CreateSheet_ShouldInitializeDefaultState() {
        SpreadsheetDocument document = new(new NullRenderer());

        SheetDefinition sut = document.CreateSheet("Sheet1");

        Assert.Multiple(() => {
            Assert.That(sut.SheetName, Is.EqualTo("Sheet1"));
            Assert.That(sut.Password, Is.Null);
            Assert.That(sut.FreezePanes, Is.Null);
            Assert.That(sut.IsAutoFilterEnabled, Is.False);
            Assert.That(sut.PageSettings, Is.Not.Null);
            Assert.That(sut.ColumnWidths, Is.Empty);
            Assert.That(sut.Templates, Is.Empty);
            Assert.That(sut.Metadata, Is.Empty);
        });
    }

    [Test]
    public void SetFreezePanes_ShouldStoreColumnAndRow() {
        SpreadsheetDocument document = new(new NullRenderer());
        SheetDefinition sut = document.CreateSheet("Sheet1");

        sut.SetFreezePanes(0, 2);

        Assert.That(sut.FreezePanes, Is.EqualTo(new Point(0, 2)));
    }

    [Test]
    public void SetAutoFilterEnabled_ShouldStoreTrue() {
        SpreadsheetDocument document = new(new NullRenderer());
        SheetDefinition sut = document.CreateSheet("Sheet1");

        sut.SetAutoFilterEnabled();

        Assert.That(sut.IsAutoFilterEnabled, Is.True);
    }

    [Test]
    public void SetAutoFilterEnabled_WhenFalse_ShouldStoreFalse() {
        SpreadsheetDocument document = new(new NullRenderer());
        SheetDefinition sut = document.CreateSheet("Sheet1");
        sut.SetAutoFilterEnabled();

        sut.SetAutoFilterEnabled(false);

        Assert.That(sut.IsAutoFilterEnabled, Is.False);
    }

    [Test]
    public void SetColumnWidth_ShouldStoreConfiguredWidth() {
        SpreadsheetDocument document = new(new NullRenderer());
        SheetDefinition sut = document.CreateSheet("Sheet1");

        sut.SetColumnWidth(1, 24d);

        Assert.That(sut.ColumnWidths[1], Is.EqualTo(24d));
    }

    [Test]
    public void AddTemplateAndAddTemplates_ShouldAppendTemplatesInOrder() {
        SpreadsheetDocument document = new(new NullRenderer());
        SheetDefinition sut = document.CreateSheet("Sheet1");
        ISheetTemplate first = Substitute.For<ISheetTemplate>();
        ISheetTemplate second = Substitute.For<ISheetTemplate>();
        ISheetTemplate third = Substitute.For<ISheetTemplate>();

        sut.AddTemplate(first);
        sut.AddTemplates(second, third);

        Assert.That(sut.Templates, Is.EqualTo(new[] { first, second, third }));
    }

    [Test]
    public void Metadata_ShouldBeMutableDictionary() {
        SpreadsheetDocument document = new(new NullRenderer());
        SheetDefinition sut = document.CreateSheet("Sheet1");

        sut.Metadata["key"] = "value";

        Assert.That(sut.Metadata["key"], Is.EqualTo("value"));
    }

    private sealed class NullRenderer : ISpreadsheetRenderer {
        public string ContentType => "application/test";

        public string FileNameExtension => ".test";

        public byte[] Render(IEnumerable<SheetLayout> contexts) => [];
    }
}
