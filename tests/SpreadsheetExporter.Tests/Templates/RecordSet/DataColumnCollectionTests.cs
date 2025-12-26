using System.Linq.Expressions;
using CloudyWing.SpreadsheetExporter.Templates.RecordSet;

namespace CloudyWing.SpreadsheetExporter.Tests.Templates.RecordSet {
    [TestFixture]
    internal class DataColumnCollectionTests {
        private const string HeaderText = "Header Text";
        private static CellStyle headerStyle = new();
        private static CellStyle fieldStyle = new();
        private static readonly Func<RecordContext<Record>, CellStyle> fieldStyleGenerator = (x) => fieldStyle;
        private DataColumnCollection<Record>? columns;

        [SetUp]
        public void SetUp() {
            columns = new DataColumnCollection<Record>(null);
        }

        [Test]
        public void ColumnSpan_EmptyColumns_ShouldReturnZero() {
            Assert.That(columns!.ColumnSpan, Is.EqualTo(0));
        }

        [Test]
        public void ColumnSpan_HasMultipleColumns_ShouldReturnSum() {
            columns!.Add("Column1");
            columns!.Add("Column2");

            Assert.That(columns!.ColumnSpan, Is.EqualTo(2));
        }

        [Test]
        public void RowSpan_EmptyColumns_ShouldReturnZero() {
            Assert.That(columns!.RowSpan, Is.EqualTo(0));
        }

        [Test]
        public void RootColumns_FromChildColumns_ShouldReturnRootColumns_() {
            columns!.Add("Column1");

            DataColumnCollection<Record> childColumns = columns.Last().ChildColumns;

            Assert.That(childColumns.RootColumns, Is.SameAs(columns));
        }

        [Test]
        public void RowSpan_ColumnsWithChildColumn_ShouldReturnTwo() {
            columns!.Add("Column1");
            columns.AddChildToLast("Column2");

            Assert.That(columns.RowSpan, Is.EqualTo(2));
        }

        [Test]
        public void Add_ByRecordDataColumn_ShouldAddToColumns() {
            RecordDataColumn<Record> column = new();
            columns!.Add(column);

            Assert.That(columns, Does.Contain(column));
        }

        [Test]
        public void Add_ByHeaderText_ShouldAddToColumns() {
            columns!.Add(HeaderText, headerStyle, fieldStyleGenerator);

            Assert.That(columns, Has.Count.EqualTo(1));

            RecordDataColumn<Record>? column = columns.Single() as RecordDataColumn<Record>;

            Assert.That(column, Is.Not.Null);
            Assert.That(column!.HeaderText, Is.EqualTo(HeaderText));
            Assert.That(column.HeaderStyle, Is.EqualTo(headerStyle));
            Assert.That(column.FieldStyleGenerator, Is.EqualTo(fieldStyleGenerator));
        }

        [Test]
        public void Add_ByValueGenerator_ShouldAddToColumns() {
            Func<RecordContext<Record>, object> fieldValueGenerator = x => x.Record.Id;

            columns!.Add(HeaderText, x => x.UseValue(fieldValueGenerator), headerStyle, fieldStyleGenerator);

            Assert.That(columns, Has.Count.EqualTo(1));
            RecordDataColumn<Record>? column = columns.Single() as RecordDataColumn<Record>;

            Assert.That(column, Is.Not.Null);
            Assert.That(column!.HeaderText, Is.EqualTo(HeaderText));
            Assert.That(column.FieldValueGenerator, Is.EqualTo(fieldValueGenerator));
            Assert.That(column.HeaderStyle, Is.EqualTo(headerStyle));
            Assert.That(column.FieldStyleGenerator, Is.EqualTo(fieldStyleGenerator));
        }

        [Test]
        public void Add_ByFormulaGenerator_ShouldAddToColumns() {
            Func<RecordContext<Record>, string> fieldFormulaGenerator = x => "SUM(A1 + B1)";

            columns!.Add(HeaderText, x => x.UseFormula(fieldFormulaGenerator), headerStyle, fieldStyleGenerator);

            Assert.That(columns, Has.Count.EqualTo(1));
            RecordDataColumn<Record>? column = columns.Single() as RecordDataColumn<Record>;

            Assert.That(column, Is.Not.Null);
            Assert.That(column!.HeaderText, Is.EqualTo(HeaderText));
            Assert.That(column.FieldFormulaGenerator, Is.EqualTo(fieldFormulaGenerator));
            Assert.That(column.HeaderStyle, Is.EqualTo(headerStyle));
            Assert.That(column.FieldStyleGenerator, Is.EqualTo(fieldStyleGenerator));
        }

        [Test]
        public void Add_ByFieldKey_ShouldAddToColumns() {
            columns!.Add<int>(HeaderText, "Id", headerStyle, fieldStyleGenerator);

            Assert.That(columns, Has.Count.EqualTo(1));
            DataColumn<Record, int>? column = columns.Single() as DataColumn<Record, int>;

            Assert.That(column, Is.Not.Null);
            Assert.That(column!.HeaderText, Is.EqualTo(HeaderText));
            Assert.That(column.FieldKey, Is.EqualTo("Id"));
            Assert.That(column.HeaderStyle, Is.EqualTo(headerStyle));
            Assert.That(column.FieldStyleGenerator, Is.EqualTo(fieldStyleGenerator));
        }

        [Test]
        public void Add_ByFieldKeyExpression_ShouldAddToColumns() {
            Expression<Func<Record, int>> fieldKeyExpression = x => x.Id;

            columns!.Add(HeaderText, x => x.Id, headerStyle, fieldStyleGenerator);

            Assert.That(columns, Has.Count.EqualTo(1));
            DataColumn<Record, int>? column = columns.Single() as DataColumn<Record, int>;

            Assert.That(column, Is.Not.Null);
            Assert.That(column!.HeaderText, Is.EqualTo(HeaderText));
            Assert.That(column.FieldKey, Is.EqualTo("Id"));
            Assert.That(column.HeaderStyle, Is.EqualTo(headerStyle));
            Assert.That(column.FieldStyleGenerator, Is.EqualTo(fieldStyleGenerator));
        }

        [Test]
        public void Add_ByFieldKeyExpressionWithGenerator_ShouldAddToColumns() {
            Func<FieldContext<Record, int>, object> fieldValueGenerator = x => x.Record.Id;

            columns!.Add(HeaderText, x => x.Id, x => x.UseValue(fieldValueGenerator), headerStyle, fieldStyleGenerator);

            Assert.That(columns, Has.Count.EqualTo(1));
            DataColumn<Record, int>? column = columns.Single() as DataColumn<Record, int>;

            Assert.That(column, Is.Not.Null);
            Assert.That(column!.HeaderText, Is.EqualTo(HeaderText));
            Assert.That(column.FieldKey, Is.EqualTo("Id"));
            Assert.That(column.FieldValueGenerator, Is.EqualTo(fieldValueGenerator));
            Assert.That(column.HeaderStyle, Is.EqualTo(headerStyle));
            Assert.That(column.FieldStyleGenerator, Is.EqualTo(fieldStyleGenerator));
        }

        [Test]
        public void AddChildToLast_ByRecordDataColumn_ShouldAddToColumns() {
            RecordDataColumn<Record> column = new();
            columns!.Add(HeaderText);
            columns.AddChildToLast(column);

            Assert.That(columns.Last().ChildColumns, Does.Contain(column));
        }

        [Test]
        public void AddChildToLast_ByHeaderText_ShouldAddToColumns() {
            columns!.Add(HeaderText);
            columns.AddChildToLast(HeaderText, headerStyle, fieldStyleGenerator);

            Assert.That(columns, Has.Count.EqualTo(1));

            DataColumnCollection<Record> lastChildColumns = columns.Last().ChildColumns;
            Assert.That(lastChildColumns, Has.Count.EqualTo(1));

            RecordDataColumn<Record>? column = lastChildColumns.Single() as RecordDataColumn<Record>;

            Assert.That(column, Is.Not.Null);
            Assert.That(column!.HeaderText, Is.EqualTo(HeaderText));
            Assert.That(column.HeaderStyle, Is.EqualTo(headerStyle));
            Assert.That(column.FieldStyleGenerator, Is.EqualTo(fieldStyleGenerator));
        }

        [Test]
        public void AddChildToLast_ByValueGenerator_ShouldAddToColumns() {
            Func<RecordContext<Record>, object> fieldValueGenerator = x => x.Record.Id;

            columns!.Add(HeaderText);
            columns.AddChildToLast(HeaderText, x => x.UseValue(fieldValueGenerator), headerStyle, fieldStyleGenerator);

            Assert.That(columns, Has.Count.EqualTo(1));

            DataColumnCollection<Record> lastChildColumns = columns.Last().ChildColumns;
            Assert.That(lastChildColumns, Has.Count.EqualTo(1));

            RecordDataColumn<Record>? column = lastChildColumns.Single() as RecordDataColumn<Record>;

            Assert.That(column, Is.Not.Null);
            Assert.That(column!.HeaderText, Is.EqualTo(HeaderText));
            Assert.That(column.FieldValueGenerator, Is.EqualTo(fieldValueGenerator));
            Assert.That(column.HeaderStyle, Is.EqualTo(headerStyle));
            Assert.That(column.FieldStyleGenerator, Is.EqualTo(fieldStyleGenerator));
        }

        [Test]
        public void AddChildToLast_ByFormulaGenerator_ShouldAddToColumns() {
            Func<RecordContext<Record>, string> fieldFormulaGenerator = x => "SUM(A1 + B1)";

            columns!.Add(HeaderText);
            columns.AddChildToLast(HeaderText, x => x.UseFormula(fieldFormulaGenerator), headerStyle, fieldStyleGenerator);

            Assert.That(columns, Has.Count.EqualTo(1));

            DataColumnCollection<Record> lastChildColumns = columns.Last().ChildColumns;
            Assert.That(lastChildColumns, Has.Count.EqualTo(1));

            RecordDataColumn<Record>? column = lastChildColumns.Single() as RecordDataColumn<Record>;

            Assert.That(column, Is.Not.Null);
            Assert.That(column!.HeaderText, Is.EqualTo(HeaderText));
            Assert.That(column.FieldFormulaGenerator, Is.EqualTo(fieldFormulaGenerator));
            Assert.That(column.HeaderStyle, Is.EqualTo(headerStyle));
            Assert.That(column.FieldStyleGenerator, Is.EqualTo(fieldStyleGenerator));
        }

        [Test]
        public void AddChildToLast_ByFieldKey_ShouldAddToColumns() {
            columns!.Add(HeaderText);
            columns.AddChildToLast<int>(HeaderText, "Id", headerStyle, fieldStyleGenerator);

            Assert.That(columns, Has.Count.EqualTo(1));

            DataColumnCollection<Record> lastChildColumns = columns.Last().ChildColumns;
            Assert.That(lastChildColumns, Has.Count.EqualTo(1));

            DataColumn<Record, int>? column = lastChildColumns.Single() as DataColumn<Record, int>;

            Assert.That(column, Is.Not.Null);
            Assert.That(column!.HeaderText, Is.EqualTo(HeaderText));
            Assert.That(column.FieldKey, Is.EqualTo("Id"));
            Assert.That(column.HeaderStyle, Is.EqualTo(headerStyle));
            Assert.That(column.FieldStyleGenerator, Is.EqualTo(fieldStyleGenerator));
        }

        [Test]
        public void AddChildToLast_ByFieldKeyExpression_ShouldAddToColumns() {
            Expression<Func<Record, int>> fieldKeyExpression = x => x.Id;

            columns!.Add(HeaderText);
            columns.AddChildToLast(HeaderText, x => x.Id, headerStyle, fieldStyleGenerator);

            Assert.That(columns, Has.Count.EqualTo(1));

            DataColumnCollection<Record> lastChildColumns = columns.Last().ChildColumns;
            Assert.That(lastChildColumns, Has.Count.EqualTo(1));

            DataColumn<Record, int>? column = lastChildColumns.Single() as DataColumn<Record, int>;

            Assert.That(column, Is.Not.Null);
            Assert.That(column!.HeaderText, Is.EqualTo(HeaderText));
            Assert.That(column.FieldKey, Is.EqualTo("Id"));
            Assert.That(column.HeaderStyle, Is.EqualTo(headerStyle));
            Assert.That(column.FieldStyleGenerator, Is.EqualTo(fieldStyleGenerator));
        }

        [Test]
        public void AddChildToLast_ByFieldKeyExpressionWithGenerator_ShouldAddToColumns() {
            Func<FieldContext<Record, int>, object> fieldValueGenerator = x => x.Record.Id;

            columns!.Add(HeaderText);
            columns.AddChildToLast(HeaderText, x => x.Id, x => x.UseValue(fieldValueGenerator), headerStyle, fieldStyleGenerator);

            Assert.That(columns, Has.Count.EqualTo(1));

            DataColumnCollection<Record> lastChildColumns = columns.Last().ChildColumns;
            Assert.That(lastChildColumns, Has.Count.EqualTo(1));

            DataColumn<Record, int>? column = lastChildColumns.Single() as DataColumn<Record, int>;

            Assert.That(column, Is.Not.Null);
            Assert.That(column!.HeaderText, Is.EqualTo(HeaderText));
            Assert.That(column.FieldKey, Is.EqualTo("Id"));
            Assert.That(column.FieldValueGenerator, Is.EqualTo(fieldValueGenerator));
            Assert.That(column.HeaderStyle, Is.EqualTo(headerStyle));
            Assert.That(column.FieldStyleGenerator, Is.EqualTo(fieldStyleGenerator));
        }

        private class Record {
            public int Id { get; set; }

            public string? Name { get; set; }
        }
    }
}
