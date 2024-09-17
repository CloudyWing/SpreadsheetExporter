using CloudyWing.SpreadsheetExporter.Templates.RecordSet;

namespace CloudyWing.SpreadsheetExporter.Tests.Templates.RecordSet {
    [TestFixture]
    internal class RecordDataColumnTests {
        private readonly RecordContext<Record> context = new(0, 0, new Record { Id = 0 });
        private RecordDataColumn<Record>? column;

        [SetUp]
        public void SetUp() {
            column = new RecordDataColumn<Record>();
        }

        [Test]
        public void GetFieldValue_FieldValueGeneratorIsNull_ShouldReturnNull() {
            column!.GetFieldValue(context).Should().BeNull();
        }

        [Test]
        public void GetFieldValue_FieldValueGeneratorIsNotNull_ShouldReturnGeneratedValue() {
            string expected = "test";
            column!.FieldValueGenerator = (x) => expected;

            column!.GetFieldValue(context).Should().Be(expected);
        }

        [Test]
        public void GetFieldStyle_FieldFormulaGeneratorIsNull_ShouldReturnDefaultStyle() {
            column!.GetFieldStyle(context).Should().Be(SpreadsheetManager.DefaultCellStyles.FieldStyle);
        }

        [Test]
        public void GetFieldStyle_FieldFormulaGeneratorIsNotNull_ShouldReturnGeneratedStyle() {
            CellStyle cellStyle = new();
            column!.FieldStyleGenerator = (x) => cellStyle;

            column!.GetFieldStyle(context).Should().Be(cellStyle);
        }

        [Test]
        public void GetFieldFormula_FieldFormulaGeneratorIsNull_ShouldReturnNull() {
            column!.GetFieldFormula(context).Should().BeNull();
        }

        [Test]
        public void GetFieldFormula_FieldFormulaGeneratorIsNotNull_ShouldReturnGeneratedFormula() {
            string expected = "SUM(A1:A5)";
            column!.FieldFormulaGenerator = x => expected;

            column.GetFieldFormula(context).Should().Be(expected);
        }

        private class Record {
            public int Id { get; set; }
        }
    }
}
