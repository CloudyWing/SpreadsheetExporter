using CloudyWing.SpreadsheetExporter.Templates.RecordSet;

namespace CloudyWing.SpreadsheetExporter.Tests.Templates.RecordSet;

[TestFixture]
internal class FieldColumnTests {
    private readonly RecordContext<Record> context = new(0, 0, new Record {
        Id = 7,
        Name = "Alice",
        Nested = new NestedRecord { Code = "A-01" }
    });

    [Test]
    public void GetFieldValue_WhenFieldKeyIsEmptyAndGeneratorMissing_ShouldReturnEmptyString() {
        FieldColumn<Record, int> sut = new("");

        Assert.That(sut.GetFieldValue(context), Is.EqualTo(""));
    }

    [Test]
    public void GetFieldValue_WhenFieldKeyExists_ShouldReturnRecordValue() {
        FieldColumn<Record, int> sut = new(x => x.Id);

        Assert.That(sut.GetFieldValue(context), Is.EqualTo(7));
    }

    [Test]
    public void GetFieldValue_WhenNestedFieldKeyExists_ShouldReturnNestedValue() {
        FieldColumn<Record, string> sut = new(x => x.Nested!.Code!);

        Assert.That(sut.GetFieldValue(context), Is.EqualTo("A-01"));
    }

    [Test]
    public void GetFieldValue_WhenFieldKeyDoesNotExist_ShouldThrowArgumentException() {
        FieldColumn<Record, int> sut = new("Missing");

        Assert.That(
            () => sut.GetFieldValue(context),
            Throws.TypeOf<ArgumentException>()
                .And.Message.Contains("Data source does not contain property 'Missing'.")
                .And.Message.Contains("Available properties")
        );
    }

    [Test]
    public void Generators_ShouldOverrideDefaultValueAndProvideFormulaStyleAndValidation() {
        FieldColumn<Record, int> sut = new(x => x.Id) {
            FieldValueGenerator = field => $"{field.Record.Name}:{field.Value}",
            FieldFormulaGenerator = field => $"=A{field.RowIndex + 1}",
            FieldStyleGenerator = _ => new CellStyle(Font: new CellFont(Style: FontStyles.Bold)),
            FieldDataValidationGenerator = _ => new DataValidation {
                ValidationType = DataValidationType.Integer,
                Operator = DataValidationOperator.GreaterThan,
                Value1 = 0
            }
        };

        Assert.Multiple(() => {
            Assert.That(sut.GetFieldValue(context), Is.EqualTo("Alice:7"));
            Assert.That(sut.GetFieldFormula(context), Is.EqualTo("=A1"));
            Assert.That(sut.GetFieldStyle(context)!.Font.Style, Is.EqualTo(FontStyles.Bold));
            Assert.That(sut.GetFieldDataValidation(context)!.ValidationType, Is.EqualTo(DataValidationType.Integer));
        });
    }

    [Test]
    public void GetFieldFormulaAndValidation_WhenGeneratorsAreMissing_ShouldReturnNull() {
        FieldColumn<Record, int> sut = new(x => x.Id);

        Assert.Multiple(() => {
            Assert.That(sut.GetFieldFormula(context), Is.Null);
            Assert.That(sut.GetFieldDataValidation(context), Is.Null);
            Assert.That(sut.GetFieldStyle(context), Is.EqualTo(SpreadsheetManager.DefaultCellStyles.FieldStyle));
        });
    }

    private sealed class Record {
        public int Id { get; init; }

        public string? Name { get; init; }

        public NestedRecord? Nested { get; init; }
    }

    private sealed class NestedRecord {
        public string? Code { get; init; }
    }
}
