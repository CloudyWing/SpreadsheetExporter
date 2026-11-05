using System.Drawing;
using CloudyWing.SpreadsheetExporter.Templates;
using CloudyWing.SpreadsheetExporter.Templates.DataTable;
using CloudyWing.SpreadsheetExporter.Templates.Grid;
using CloudyWing.SpreadsheetExporter.Templates.RecordSet;

namespace CloudyWing.SpreadsheetExporter.Tests;

[TestFixture]
internal class SheetLayoutTests {
    [Test]
    public void Constructor_ShouldProjectSheetDefinitionState() {
        SpreadsheetDocument document = new(new NullRenderer());
        SheetDefinition sheet = document.CreateSheet("Sheet1", 22d);
        sheet.Password = "sheet-password";
        sheet.PageSettings.PaperSize = PaperSize.A4;
        sheet.Metadata["source"] = "test";
        sheet.SetColumnWidth(1, 50d);
        sheet.SetColumnWidth(2, 75d);
        sheet.FreezePanes = new Point(1, 2);
        sheet.IsAutoFilterEnabled = true;

        TemplateLayout templateLayout = new(
            [new Cell {
                Point = new Point(0, 0),
                Size = new Size(1, 1),
                ValueGenerator = (_, _) => "header"
            }],
            2,
            new Dictionary<int, double?> { [0] = 10d, [1] = 12d }
        );

        sheet.AddTemplate(new StubSheetTemplate(templateLayout));
        CellFont defaultFont = new("Calibri", 11);

        SheetLayout sut = new(sheet, defaultFont);

        Assert.Multiple(() => {
            Assert.That(sut.SheetName, Is.EqualTo(sheet.SheetName));
            Assert.That(sut.DefaultRowHeight, Is.EqualTo(22d));
            Assert.That(sut.Cells, Has.Count.EqualTo(1));
            Assert.That(sut.ColumnWidths[1], Is.EqualTo(50d));
            Assert.That(sut.ColumnWidths[2], Is.EqualTo(75d));
            Assert.That(sut.RowHeights[0], Is.EqualTo(10d));
            Assert.That(sut.Password, Is.EqualTo("sheet-password"));
            Assert.That(sut.IsProtected, Is.True);
            Assert.That(sut.PageSettings, Is.SameAs(sheet.PageSettings));
            Assert.That(sut.FreezePanes, Is.EqualTo(new Point(1, 2)));
            Assert.That(sut.IsAutoFilterEnabled, Is.True);
            Assert.That(sut.Metadata, Is.SameAs(sheet.Metadata));
            Assert.That(sut.Metadata["source"], Is.EqualTo("test"));
            Assert.That(sut.DefaultFont, Is.EqualTo(defaultFont));
        });
    }

    [Test]
    public void Constructor_WhenOptionalValuesOmitted_ShouldExposeDefaults() {
        SpreadsheetDocument document = new(new NullRenderer());
        SheetDefinition sheet = document.CreateSheet("Sheet1");
        sheet.AddTemplate(new StubSheetTemplate(new TemplateLayout([], 0, new Dictionary<int, double?>())));

        SheetLayout sut = new(sheet);

        Assert.Multiple(() => {
            Assert.That(sut.DefaultFont, Is.Null);
            Assert.That(sut.FreezePanes, Is.Null);
            Assert.That(sut.IsAutoFilterEnabled, Is.False);
        });
    }

    [Test]
    public void Constructor_AutoFilterRangeIsSet_ShouldUseConfiguredRange() {
        SheetDefinition sheet = new SpreadsheetDocument(new NullRenderer()).CreateSheet("Sheet1");
        sheet.AddTemplate(CreateTitleGrid()).AddTemplate(CreateRecordSet());
        sheet.SetAutoFilterRange(new CellRange(1, 2, 1, 3));

        SheetLayout sut = new(sheet);

        Assert.Multiple(() => {
            Assert.That(sut.IsAutoFilterEnabled, Is.True);
            Assert.That(sut.AutoFilterRange, Is.EqualTo(new CellRange(1, 2, 1, 3)));
        });
    }

    [Test]
    public void Constructor_TitleGridAboveRecordSet_ShouldDefaultAutoFilterRangeToRecordSet() {
        SheetDefinition sheet = new SpreadsheetDocument(new NullRenderer()).CreateSheet("Sheet1");
        sheet.AddTemplate(CreateTitleGrid()).AddTemplate(CreateRecordSet()).AddTemplate(CreateTitleGrid());

        SheetLayout sut = new(sheet);

        Assert.That(sut.AutoFilterRange, Is.EqualTo(new CellRange(0, 1, 2, 3)));
    }

    [Test]
    public void Constructor_TitleGridAboveDataTable_ShouldDefaultAutoFilterRangeToDataTable() {
        System.Data.DataTable dataTable = new();
        dataTable.Columns.Add("Id", typeof(int));
        dataTable.Rows.Add(1);
        SheetDefinition sheet = new SpreadsheetDocument(new NullRenderer()).CreateSheet("Sheet1");
        sheet.AddTemplate(CreateTitleGrid()).AddTemplate(new DataTableTemplate(dataTable));

        SheetLayout sut = new(sheet);

        Assert.That(sut.AutoFilterRange, Is.EqualTo(new CellRange(0, 1, 1, 2)));
    }

    [Test]
    public void Constructor_SingleRecordSet_ShouldLeaveAutoFilterRangeToRenderer() {
        SheetDefinition sheet = new SpreadsheetDocument(new NullRenderer()).CreateSheet("Sheet1");
        sheet.AddTemplate(CreateRecordSet());

        SheetLayout sut = new(sheet);

        Assert.That(sut.AutoFilterRange, Is.Null);
    }

    [Test]
    public void Constructor_MultipleTemplatesWithoutTable_ShouldLeaveAutoFilterRangeToRenderer() {
        SheetDefinition sheet = new SpreadsheetDocument(new NullRenderer()).CreateSheet("Sheet1");
        sheet.AddTemplate(CreateTitleGrid()).AddTemplate(CreateTitleGrid());

        SheetLayout sut = new(sheet);

        Assert.That(sut.AutoFilterRange, Is.Null);
    }

    private static GridTemplate CreateTitleGrid() {
        return new GridTemplate().CreateCell("Title", columnSpan: 2);
    }

    private static RecordSetTemplate<(int Id, string Name)> CreateRecordSet() {
        RecordSetTemplate<(int Id, string Name)> template = new([(1, "A"), (2, "B")]);
        template.Columns
            .Add("Id", x => x.Id)
            .Add("Name", x => x.Name);
        return template;
    }

    private sealed class StubSheetTemplate(TemplateLayout layout) : ISheetTemplate {
        public int ColumnSpan => 0;

        public int RowSpan => layout.RowSpan;

        public TemplateLayout GetLayout() => layout;
    }

    private sealed class NullRenderer : ISpreadsheetRenderer {
        public string ContentType => "application/test";

        public string FileNameExtension => ".test";

        public byte[] Render(IEnumerable<SheetLayout> layouts) => [];
    }
}
