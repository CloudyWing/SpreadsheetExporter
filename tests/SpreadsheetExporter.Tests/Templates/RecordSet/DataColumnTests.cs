using CloudyWing.SpreadsheetExporter.Templates.RecordSet;

namespace CloudyWing.SpreadsheetExporter.Tests.Templates.RecordSet {
    [TestFixture]
    internal class DataColumnTests {
        private DataColumn<Record, int>? columnWithEmptyKey;
        private DataColumn<Record, int>? columnWithCorrectKey;
        private DataColumn<Record, int>? columnWithIncorrectKey;
        private readonly RecordContext<Record> context = new(0, 0, new Record { Id = 0 });
        private readonly CellStyle cellStyle = new();

        [SetUp]
        public void SetUp() {
            columnWithEmptyKey = new DataColumn<Record, int>("");
            columnWithCorrectKey = new DataColumn<Record, int>(x => x.Id);
            columnWithIncorrectKey = new DataColumn<Record, int>("Id2");
        }

        [Test]
        public void GetFieldValue_FieldKeyIsEmptyAndFieldValueGeneratorIsNull_ShouldReturnEmpty() {
            object result = columnWithEmptyKey!.GetFieldValue(context);

            result.Should().Be("");
        }

        [Test]
        public void GetFieldValue_FieldKeyIsCorrectAndFieldValueGeneratorIsNull_ReturnRecordValue() {
            object result = columnWithCorrectKey!.GetFieldValue(context);

            result.Should().Be(context.Record.Id);
        }

        [Test]
        public void GetFieldValue_FieldKeyIsIncorrectAndFieldValueGeneratorIsNull_ThrowArgumentException() {
            Func<object> act = () => columnWithIncorrectKey!.GetFieldValue(context);

            act.Should().Throw<ArgumentException>();
        }

        [Test]
        public void GetFieldValue_FieldKeyIsIncorrectAndFieldValueGeneratorIsNotNull_ThrowArgumentException() {
            string expected = "ExpectedResult";
            Func<object> act = () => columnWithIncorrectKey!.GetFieldValue(context).Should().Be(expected);

            act.Should().Throw<ArgumentException>();
        }

        [Test]
        public void GetFieldValue_FieldValueGeneratorIsNotNull_ShouldReturnGeneratedValue() {
            string expected = "ExpectedResult";
            string generator(FieldContext<Record, int> x) => expected;
            columnWithEmptyKey!.FieldValueGenerator = (Func<FieldContext<Record, int>, string>)generator;
            columnWithCorrectKey!.FieldValueGenerator = (Func<FieldContext<Record, int>, string>)generator;

            columnWithEmptyKey!.GetFieldValue(context).Should().Be(expected);
            columnWithCorrectKey!.GetFieldValue(context).Should().Be(expected);
        }

        [Test]
        public void GetFieldFormula_FormulaGeneratorIsNull_ShouldReturnNull() {
            columnWithEmptyKey!.GetFieldFormula(context).Should().BeNull();
            columnWithCorrectKey!.GetFieldFormula(context).Should().BeNull();
            columnWithIncorrectKey!.GetFieldFormula(context).Should().BeNull();
        }

        [Test]
        public void GetFieldFormula_FieldKeyIsIncorrectAndFormulaGeneratorIsNotNull_ThrowArgumentException() {
            string expected = "=SUM(A1:A5)";
            columnWithIncorrectKey!.FieldFormulaGenerator = x => expected;
            Func<object> act = () => columnWithIncorrectKey!.GetFieldFormula(context).Should().Be(expected);

            act.Should().Throw<ArgumentException>();
        }

        [Test]
        public void GetFieldFormula_FieldFormulaGeneratorIsNotNull_ShouldReturnGeneratedFormula() {
            string expected = "SUM(A1:A5)";
            columnWithEmptyKey!.FieldFormulaGenerator = x => expected;
            columnWithCorrectKey!.FieldFormulaGenerator = x => expected;

            columnWithEmptyKey.GetFieldFormula(context).Should().Be(expected);
            columnWithCorrectKey.GetFieldFormula(context).Should().Be(expected);
        }

        [Test]
        public void GetGetFieldStyle_FieldStyleGeneratorIsNull_ShouldReturnDefaultStyle() {
            columnWithEmptyKey!.GetFieldStyle(context).Should().Be(SpreadsheetManager.DefaultCellStyles.FieldStyle);
            columnWithCorrectKey!.GetFieldStyle(context).Should().Be(SpreadsheetManager.DefaultCellStyles.FieldStyle);
            columnWithIncorrectKey!.GetFieldStyle(context).Should().Be(SpreadsheetManager.DefaultCellStyles.FieldStyle);
        }

        [Test]
        public void GetGetFieldStyle_FieldKeyIsIncorrectAndFieldStyleGeneratorIsNotNull_ThrowArgumentException() {
            CellStyle expected = cellStyle;
            CellStyle generator(FieldContext<Record, int> x) => cellStyle;
            columnWithIncorrectKey!.FieldStyleGenerator = generator;
            Func<object> act = () => columnWithIncorrectKey!.GetFieldStyle(context).Should().Be(expected);

            act.Should().Throw<ArgumentException>();
        }

        [Test]
        public void GetGetFieldStyle_FieldStyleGeneratorIsNotNull_ShouldReturnGeneratedStyle() {
            CellStyle expected = cellStyle;
            CellStyle generator(FieldContext<Record, int> x) => cellStyle;
            columnWithEmptyKey!.FieldStyleGenerator = generator;
            columnWithCorrectKey!.FieldStyleGenerator = generator;

            columnWithEmptyKey!.GetFieldStyle(context).Should().Be(expected);
            columnWithCorrectKey!.GetFieldStyle(context).Should().Be(expected);
        }

        private class Record {
            public int Id { get; set; }
        }
    }
}
