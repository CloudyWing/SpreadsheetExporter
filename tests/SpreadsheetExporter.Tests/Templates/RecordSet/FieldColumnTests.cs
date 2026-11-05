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

    [Test]
    public void GetFieldValue_NestedObjectIsNullWithExpression_ShouldReturnNull() {
        RecordContext<Record> nullNestedContext = new(0, 0, new Record { Id = 1 });
        FieldColumn<Record, string> sut = new(x => x.Nested!.Code!);

        Assert.That(sut.GetFieldValue(nullNestedContext), Is.Null);
    }

    [Test]
    public void GetFieldValue_NestedObjectIsNullWithFieldKey_ShouldReturnNull() {
        RecordContext<Record> nullNestedContext = new(0, 0, new Record { Id = 1 });
        FieldColumn<Record, string> sut = new("Nested.Code");

        Assert.That(sut.GetFieldValue(nullNestedContext), Is.Null);
    }

    [TestCase("Nested.Code")]
    [TestCase("nested.code")]
    public void GetFieldValue_NestedFieldKey_ShouldReturnNestedValue(string fieldKey) {
        FieldColumn<Record, string> sut = new(fieldKey);

        Assert.That(sut.GetFieldValue(context), Is.EqualTo("A-01"));
    }

    [Test]
    public void GetFieldValue_RecordTypeHasIndexer_ShouldReturnPropertyValue() {
        RecordContext<IndexedRecord> indexedContext = new(0, 0, new IndexedRecord { Name = "Indexed" });
        FieldColumn<IndexedRecord, string> sut = new("Name");

        Assert.That(sut.GetFieldValue(indexedContext), Is.EqualTo("Indexed"));
    }

    [Test]
    public void GetFieldValue_PropertyHiddenByNewWithFieldKey_ShouldReturnMostDerivedValue() {
        DerivedRecord record = new() { Code = "Derived" };
        ((BaseRecord)record).Code = 9;
        FieldColumn<BaseRecord, object> sut = new("Code");

        Assert.That(sut.GetFieldValue(new RecordContext<BaseRecord>(0, 0, record)), Is.EqualTo("Derived"));
    }

    [Test]
    public void GetFieldValue_PropertyHiddenByNewWithExpression_ShouldUseStaticallyBoundProperty() {
        DerivedRecord record = new() { Code = "Derived" };
        ((BaseRecord)record).Code = 9;
        FieldColumn<BaseRecord, int> sut = new(x => x.Code);

        Assert.That(sut.GetFieldValue(new RecordContext<BaseRecord>(0, 0, record)), Is.EqualTo(9));
    }

    [Test]
    public void GetFieldValue_ExpressionConvertsFieldType_ShouldConvertValue() {
        FieldColumn<Record, long> sut = new(x => x.Id);

        Assert.That(sut.GetFieldValue(context), Is.EqualTo(7L));
    }

    [Test]
    public void GetFieldValue_InterfaceExpression_ShouldReadImplementationValue() {
        RecordContext<INamed> namedContext = new(0, 0, new ExplicitNamed());
        FieldColumn<INamed, string> sut = new(x => x.Name);

        Assert.That(sut.GetFieldValue(namedContext), Is.EqualTo("Explicit"));
    }

    [Test]
    public void GetFieldValue_ParallelCalls_ShouldReturnEachRecordValue() {
        FieldColumn<Record, int> sut = new(x => x.Id);
        FieldColumn<Record, string> nestedSut = new("Nested.Code");
        Record[] records = Enumerable.Range(0, 2_000)
            .Select(i => new Record { Id = i, Nested = new NestedRecord { Code = $"C{i}" } })
            .ToArray();
        object?[] values = new object?[records.Length];
        object?[] nestedValues = new object?[records.Length];

        Parallel.For(0, records.Length, i => {
            RecordContext<Record> recordContext = new(0, i, records[i]);
            values[i] = sut.GetFieldValue(recordContext);
            nestedValues[i] = nestedSut.GetFieldValue(recordContext);
        });

        Assert.Multiple(() => {
            Assert.That(values, Is.EqualTo(records.Select(x => (object?)x.Id).ToArray()));
            Assert.That(nestedValues, Is.EqualTo(records.Select(x => (object?)x.Nested!.Code).ToArray()));
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

    private interface INamed {
        string Name { get; }
    }

    private sealed class ExplicitNamed : INamed {
        string INamed.Name => "Explicit";
    }

    private sealed class IndexedRecord {
        public string? Name { get; init; }

        public string this[int index] => $"{Name}{index}";
    }

    private class BaseRecord {
        public int Code { get; set; }
    }

    private sealed class DerivedRecord : BaseRecord {
        public new string? Code { get; set; }
    }
}
