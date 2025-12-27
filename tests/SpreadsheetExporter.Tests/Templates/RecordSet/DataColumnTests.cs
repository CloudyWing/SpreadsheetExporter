using CloudyWing.SpreadsheetExporter.Templates.RecordSet;

namespace CloudyWing.SpreadsheetExporter.Tests.Templates.RecordSet {
    [TestFixture]
    internal class DataColumnTests {
        private DataColumn<Record, int> columnWithEmptyKey;
        private DataColumn<Record, int> columnWithCorrectKey;
        private DataColumn<Record, int> columnWithIncorrectKey;
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
            object result = columnWithEmptyKey.GetFieldValue(context);

            Assert.That(result, Is.EqualTo(""));
        }

        [Test]
        public void GetFieldValue_FieldKeyIsCorrectAndFieldValueGeneratorIsNull_ReturnRecordValue() {
            object result = columnWithCorrectKey.GetFieldValue(context);

            Assert.That(result, Is.EqualTo(context.Record.Id));
        }

        [Test]
        public void GetFieldValue_FieldKeyIsIncorrectAndFieldValueGeneratorIsNull_ThrowArgumentException() {
            Func<object> act = () => columnWithIncorrectKey.GetFieldValue(context);

            Assert.That(act, Throws.TypeOf<ArgumentException>());
        }

        [Test]
        public void GetFieldValue_FieldKeyIsIncorrectAndFieldValueGeneratorIsNotNull_ThrowArgumentException() {
            Func<object> act = () => columnWithIncorrectKey.GetFieldValue(context);

            Assert.That(act, Throws.TypeOf<ArgumentException>());
        }

        [Test]
        public void GetFieldValue_FieldValueGeneratorIsNotNull_ShouldReturnGeneratedValue() {
            string expected = "ExpectedResult";
            string generator(FieldContext<Record, int> x) => expected;
            columnWithEmptyKey.FieldValueGenerator = (Func<FieldContext<Record, int>, string>)generator;
            columnWithCorrectKey.FieldValueGenerator = (Func<FieldContext<Record, int>, string>)generator;

            Assert.That(columnWithEmptyKey.GetFieldValue(context), Is.EqualTo(expected));
            Assert.That(columnWithCorrectKey.GetFieldValue(context), Is.EqualTo(expected));
        }

        [Test]
        public void GetFieldFormula_FormulaGeneratorIsNull_ShouldReturnNull() {
            Assert.That(columnWithEmptyKey.GetFieldFormula(context), Is.Null);
            Assert.That(columnWithCorrectKey.GetFieldFormula(context), Is.Null);
            Assert.That(columnWithIncorrectKey.GetFieldFormula(context), Is.Null);
        }

        [Test]
        public void GetFieldFormula_FieldKeyIsIncorrectAndFormulaGeneratorIsNotNull_ThrowArgumentException() {
            string expected = "=SUM(A1:A5)";
            columnWithIncorrectKey.FieldFormulaGenerator = x => expected;
            Func<object> act = () => columnWithIncorrectKey.GetFieldFormula(context);

            Assert.That(act, Throws.TypeOf<ArgumentException>());
        }

        [Test]
        public void GetFieldFormula_FieldFormulaGeneratorIsNotNull_ShouldReturnGeneratedFormula() {
            string expected = "SUM(A1:A5)";
            columnWithEmptyKey.FieldFormulaGenerator = x => expected;
            columnWithCorrectKey.FieldFormulaGenerator = x => expected;

            Assert.That(columnWithEmptyKey.GetFieldFormula(context), Is.EqualTo(expected));
            Assert.That(columnWithCorrectKey.GetFieldFormula(context), Is.EqualTo(expected));
        }

        [Test]
        public void GetGetFieldStyle_FieldStyleGeneratorIsNull_ShouldReturnDefaultStyle() {
            Assert.That(columnWithEmptyKey.GetFieldStyle(context), Is.EqualTo(SpreadsheetManager.DefaultCellStyles.FieldStyle));
            Assert.That(columnWithCorrectKey.GetFieldStyle(context), Is.EqualTo(SpreadsheetManager.DefaultCellStyles.FieldStyle));
            Assert.That(columnWithIncorrectKey.GetFieldStyle(context), Is.EqualTo(SpreadsheetManager.DefaultCellStyles.FieldStyle));
        }

        [Test]
        public void GetGetFieldStyle_FieldKeyIsIncorrectAndFieldStyleGeneratorIsNotNull_ThrowArgumentException() {
            CellStyle expected = cellStyle;
            CellStyle generator(FieldContext<Record, int> x) => cellStyle;
            columnWithIncorrectKey.FieldStyleGenerator = generator;
            Func<object> act = () => columnWithIncorrectKey.GetFieldStyle(context);

            Assert.That(act, Throws.TypeOf<ArgumentException>());
        }

        [Test]
        public void GetGetFieldStyle_FieldStyleGeneratorIsNotNull_ShouldReturnGeneratedStyle() {
            CellStyle expected = cellStyle;
            CellStyle generator(FieldContext<Record, int> x) => cellStyle;
            columnWithEmptyKey.FieldStyleGenerator = generator;
            columnWithCorrectKey.FieldStyleGenerator = generator;

            Assert.That(columnWithEmptyKey.GetFieldStyle(context), Is.EqualTo(expected));
            Assert.That(columnWithCorrectKey.GetFieldStyle(context), Is.EqualTo(expected));
        }

        private class Record {
            public int Id { get; set; }
        }
    }
}
