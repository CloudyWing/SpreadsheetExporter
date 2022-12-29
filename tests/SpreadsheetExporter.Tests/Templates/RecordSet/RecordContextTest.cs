using CloudyWing.SpreadsheetExporter.Templates.RecordSet;

namespace CloudyWing.SpreadsheetExporter.Tests.Templates.RecordSet {
    [TestFixture]
    internal class RecordContextTest {
        [Test]
        public void Constructor_GivenValidArguments_ShouldSetProperties() {
            int cellIndex = 1;
            int rowIndex = 2;
            int record = 1;

            RecordContext<int> sut = new RecordContext<int>(cellIndex, rowIndex, record);

            sut.CellIndex.Should().Be(cellIndex);
            sut.RowIndex.Should().Be(rowIndex);
            sut.Record.Should().Be(record);
        }
    }
}
