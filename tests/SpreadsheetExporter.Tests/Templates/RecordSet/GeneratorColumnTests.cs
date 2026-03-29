using CloudyWing.SpreadsheetExporter.Templates.RecordSet;

namespace CloudyWing.SpreadsheetExporter.Tests.Templates.RecordSet;

[TestFixture]
internal class GeneratorColumnTests {
    private readonly RecordContext<Record> context = new(0, 0, new Record { Id = 5 });

    [Test]
    public void DefaultGenerators_ShouldReturnNullOrDefaultFieldStyle() {
        GeneratorColumn<Record> sut = new();

        Assert.Multiple(() => {
            Assert.That(sut.GetFieldValue(context), Is.Null);
            Assert.That(sut.GetFieldFormula(context), Is.Null);
            Assert.That(sut.GetFieldDataValidation(context), Is.Null);
            Assert.That(sut.GetFieldStyle(context), Is.EqualTo(SpreadsheetManager.DefaultCellStyles.FieldStyle));
        });
    }

    [Test]
    public void ConfiguredGenerators_ShouldReturnGeneratedResults() {
        DataValidation validation = new() {
            ValidationType = DataValidationType.List,
            ListItems = ["A", "B"]
        };

        GeneratorColumn<Record> sut = new() {
            FieldValueGenerator = x => x.Record.Id,
            FieldFormulaGenerator = x => $"=A{x.RowIndex + 1}",
            FieldStyleGenerator = _ => new CellStyle(Font: new CellFont(Style: FontStyles.Bold)),
            FieldDataValidationGenerator = _ => validation
        };

        Assert.Multiple(() => {
            Assert.That(sut.GetFieldValue(context), Is.EqualTo(5));
            Assert.That(sut.GetFieldFormula(context), Is.EqualTo("=A1"));
            Assert.That(sut.GetFieldStyle(context).Font.Style, Is.EqualTo(FontStyles.Bold));
            Assert.That(sut.GetFieldDataValidation(context), Is.SameAs(validation));
        });
    }

    private sealed class Record {
        public int Id { get; init; }
    }
}
