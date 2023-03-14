using CloudyWing.SpreadsheetExporter.Templates.RecordSet;

namespace CloudyWing.SpreadsheetExporter.Tests.Templates.RecordSet {
    [TestFixture]
    internal class FieldContextTests {
        [Test]
        public void Constructor_KeyIsNull_ShouldThrowArgumentNullException() {
            RecordContext<object> recordContext = Substitute.For<RecordContext<object>>(0, 0, null);

            Action action = () => new FieldContext<object, object?>(recordContext, null, null);

            action.Should().Throw<ArgumentNullException>().WithMessage("*key*");
        }

        [Test]
        public void FieldContext_Properties_ShouldGetCorrectValue() {
            object record = new object();
            RecordContext<object> recordContext = new RecordContext<object>(1, 2, record);
            string key = "key";
            object value = new object();
            FieldContext<object, object> fieldContext = new FieldContext<object, object>(recordContext, key, value);

            int cellIndex = fieldContext.CellIndex;
            int rowIndex = fieldContext.RowIndex;
            object fieldRecord = fieldContext.Record;
            string fieldKey = fieldContext.Key;
            object fieldValue = fieldContext.Value;

            cellIndex.Should().Be(1);
            rowIndex.Should().Be(2);
            fieldRecord.Should().Be(record);
            fieldKey.Should().Be(key);
            fieldValue.Should().Be(value);
        }
    }
}
