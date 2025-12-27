using CloudyWing.SpreadsheetExporter.Templates.RecordSet;

namespace CloudyWing.SpreadsheetExporter.Tests.Templates.RecordSet {
    [TestFixture]
    internal class FieldContextTests {
        [Test]
        public void Constructor_KeyIsNull_ShouldThrowArgumentNullException() {
            RecordContext<object> recordContext = Substitute.For<RecordContext<object>>(0, 0, null);

            Action action = () => new FieldContext<object, object>(recordContext, null, null);

            Assert.That(action, Throws.TypeOf<ArgumentNullException>().And.Message.Contains("key"));
        }

        [Test]
        public void FieldContext_Properties_ShouldGetCorrectValue() {
            object record = new();
            RecordContext<object> recordContext = new(1, 2, record);
            string key = "key";
            object value = new();
            FieldContext<object, object> fieldContext = new(recordContext, key, value);

            int cellIndex = fieldContext.CellIndex;
            int rowIndex = fieldContext.RowIndex;
            object fieldRecord = fieldContext.Record;
            string fieldKey = fieldContext.Key;
            object fieldValue = fieldContext.Value;

            Assert.That(cellIndex, Is.EqualTo(1));
            Assert.That(rowIndex, Is.EqualTo(2));
            Assert.That(fieldRecord, Is.EqualTo(record));
            Assert.That(fieldKey, Is.EqualTo(key));
            Assert.That(fieldValue, Is.EqualTo(value));
        }
    }
}
