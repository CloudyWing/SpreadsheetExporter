using CloudyWing.SpreadsheetExporter.Templates.RecordSet;

namespace CloudyWing.SpreadsheetExporter.Tests.Templates.RecordSet {
    [TestFixture]
    internal class RecordContextTest {
        [Test]
        public void Constructor_GivenValidArguments_ShouldSetProperties() {
            int cellIndex = 1;
            int rowIndex = 2;
            int record = 1;

            RecordContext<int> sut = new(cellIndex, rowIndex, record);

            Assert.That(sut.CellIndex, Is.EqualTo(cellIndex));
            Assert.That(sut.RowIndex, Is.EqualTo(rowIndex));
            Assert.That(sut.Record, Is.EqualTo(record));
        }
    }
}
