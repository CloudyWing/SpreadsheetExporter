using CloudyWing.SpreadsheetExporter.Templates.RecordSet;

namespace CloudyWing.SpreadsheetExporter.Tests.Templates.RecordSet {
    [TestFixture]
    internal class RecordDataColumnTests {
        private readonly RecordContext<Record> context = new(0, 0, new Record { Id = 0 });
        private RecordDataColumn<Record> column;

        [SetUp]
        public void SetUp() {
            column = new RecordDataColumn<Record>();
        }

        [Test]
        public void GetFieldValue_FieldValueGeneratorIsNull_ShouldReturnNull() {
            Assert.That(column.GetFieldValue(context), Is.Null);
        }

        [Test]
        public void GetFieldValue_FieldValueGeneratorIsNotNull_ShouldReturnGeneratedValue() {
            string expected = "test";
            column.FieldValueGenerator = (x) => expected;

            Assert.That(column.GetFieldValue(context), Is.EqualTo(expected));
        }

        [Test]
        public void GetFieldStyle_FieldFormulaGeneratorIsNull_ShouldReturnDefaultStyle() {
            Assert.That(column.GetFieldStyle(context), Is.EqualTo(SpreadsheetManager.DefaultCellStyles.FieldStyle));
        }

        [Test]
        public void GetFieldStyle_FieldFormulaGeneratorIsNotNull_ShouldReturnGeneratedStyle() {
            CellStyle cellStyle = new();
            column.FieldStyleGenerator = (x) => cellStyle;

            Assert.That(column.GetFieldStyle(context), Is.EqualTo(cellStyle));
        }

        [Test]
        public void GetFieldFormula_FieldFormulaGeneratorIsNull_ShouldReturnNull() {
            Assert.That(column.GetFieldFormula(context), Is.Null);
        }

        [Test]
        public void GetFieldFormula_FieldFormulaGeneratorIsNotNull_ShouldReturnGeneratedFormula() {
            string expected = "SUM(A1:A5)";
            column.FieldFormulaGenerator = x => expected;

            Assert.That(column.GetFieldFormula(context), Is.EqualTo(expected));
        }

        private class Record {
            public int Id { get; set; }
        }
    }
}
