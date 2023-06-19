using System.Linq.Expressions;
using CloudyWing.SpreadsheetExporter.Templates.RecordSet;

namespace CloudyWing.SpreadsheetExporter.Tests.Templates.RecordSet {
    [TestFixture]
    internal class DataColumnCollectionTests {
        private const string HeaderText = "Header Text";
        private static CellStyle headerStyle = new CellStyle();
        private static CellStyle fieldStyle = new CellStyle();
        private static readonly Func<RecordContext<Record>, CellStyle> fieldStyleGenerator = (x) => fieldStyle;
        private DataColumnCollection<Record>? columns;

        [SetUp]
        public void SetUp() {
            columns = new DataColumnCollection<Record>(null);
        }

        [Test]
        public void ColumnSpan_EmptyColumns_ShouldReturnZero() {
            columns!.ColumnSpan.Should().Be(0);
        }

        [Test]
        public void ColumnSpan_HasMultipleColumns_ShouldReturnSum() {
            columns!.Add("Column1");
            columns!.Add("Column2");

            columns!.ColumnSpan.Should().Be(2);
        }

        [Test]
        public void RowSpan_EmptyColumns_ShouldReturnZero() {
            columns!.RowSpan.Should().Be(0);
        }

        [Test]
        public void RootColumns_FromChildColumns_ShouldReturnRootColumns_() {
            columns!.Add("Column1");

            DataColumnCollection<Record> childColumns = columns.Last().ChildColumns;

            childColumns.RootColumns.Should().BeSameAs(columns);
        }

        [Test]
        public void RowSpan_ColumnsWithChildColumn_ShouldReturnTwo() {
            columns!.Add("Column1");
            columns.AddChildToLast("Column2");

            columns.RowSpan.Should().Be(2);
        }

        [Test]
        public void Add_ByRecordDataColumn_ShouldAddToColumns() {
            RecordDataColumn<Record> column = new RecordDataColumn<Record>();
            columns!.Add(column);

            columns.Should().Contain(column);
        }

        [Test]
        public void Add_ByHeaderText_ShouldAddToColumns() {
            columns!.Add(HeaderText, headerStyle, fieldStyleGenerator);

            columns.Should().HaveCount(1);

            RecordDataColumn<Record>? column = columns.Single() as RecordDataColumn<Record>;

            column.Should().NotBeNull();
            column!.HeaderText.Should().Be(HeaderText);
            column.HeaderStyle.Should().Be(headerStyle);
            column.FieldStyleGenerator.Should().Be(fieldStyleGenerator);
        }

        [Test]
        public void Add_ByValueGenerator_ShouldAddToColumns() {
            Func<RecordContext<Record>, object> fieldValueGenerator = x => x.Record.Id;

            columns!.Add(HeaderText, x => x.UseValue(fieldValueGenerator), headerStyle, fieldStyleGenerator);

            columns.Should().HaveCount(1);
            RecordDataColumn<Record>? column = columns.Single() as RecordDataColumn<Record>;

            column.Should().NotBeNull();
            column!.HeaderText.Should().Be(HeaderText);
            column.FieldValueGenerator.Should().Be(fieldValueGenerator);
            column.HeaderStyle.Should().Be(headerStyle);
            column.FieldStyleGenerator.Should().Be(fieldStyleGenerator);
        }

        [Test]
        public void Add_ByFormulaGenerator_ShouldAddToColumns() {
            Func<RecordContext<Record>, string> fieldFormulaGenerator = x => "SUM(A1 + B1)";

            columns!.Add(HeaderText, x => x.UseFormula(fieldFormulaGenerator), headerStyle, fieldStyleGenerator);

            columns.Should().HaveCount(1);
            RecordDataColumn<Record>? column = columns.Single() as RecordDataColumn<Record>;

            column.Should().NotBeNull();
            column!.HeaderText.Should().Be(HeaderText);
            column.FieldFormulaGenerator.Should().Be(fieldFormulaGenerator);
            column.HeaderStyle.Should().Be(headerStyle);
            column.FieldStyleGenerator.Should().Be(fieldStyleGenerator);
        }

        [Test]
        public void Add_ByFieldKey_ShouldAddToColumns() {
            columns!.Add<int>(HeaderText, "Id", headerStyle, fieldStyleGenerator);

            columns.Should().HaveCount(1);
            DataColumn<Record, int>? column = columns.Single() as DataColumn<Record, int>;

            column.Should().NotBeNull();
            column!.HeaderText.Should().Be(HeaderText);
            column.FieldKey.Should().Be("Id");
            column.HeaderStyle.Should().Be(headerStyle);
            column.FieldStyleGenerator.Should().Be(fieldStyleGenerator);
        }

        [Test]
        public void Add_ByFieldKeyExpression_ShouldAddToColumns() {
            Expression<Func<Record, int>> fieldKeyExpression = x => x.Id;

            columns!.Add(HeaderText, x => x.Id, headerStyle, fieldStyleGenerator);

            columns.Should().HaveCount(1);
            DataColumn<Record, int>? column = columns.Single() as DataColumn<Record, int>;

            column.Should().NotBeNull();
            column!.HeaderText.Should().Be(HeaderText);
            column.FieldKey.Should().Be("Id");
            column.HeaderStyle.Should().Be(headerStyle);
            column.FieldStyleGenerator.Should().Be(fieldStyleGenerator);
        }

        [Test]
        public void Add_ByFieldKeyExpressionWithGenerator_ShouldAddToColumns() {
            Func<FieldContext<Record, int>, object> fieldValueGenerator = x => x.Record.Id;

            columns!.Add(HeaderText, x => x.Id, x => x.UseValue(fieldValueGenerator), headerStyle, fieldStyleGenerator);

            columns.Should().HaveCount(1);
            DataColumn<Record, int>? column = columns.Single() as DataColumn<Record, int>;

            column.Should().NotBeNull();
            column!.HeaderText.Should().Be(HeaderText);
            column.FieldKey.Should().Be("Id");
            column.FieldValueGenerator.Should().Be(fieldValueGenerator);
            column.HeaderStyle.Should().Be(headerStyle);
            column.FieldStyleGenerator.Should().Be(fieldStyleGenerator);
        }

        [Test]
        public void AddChildToLast_ByRecordDataColumn_ShouldAddToColumns() {
            RecordDataColumn<Record> column = new RecordDataColumn<Record>();
            columns!.Add(HeaderText);
            columns.AddChildToLast(column);

            columns.Last().ChildColumns.Should().Contain(column);
        }

        [Test]
        public void AddChildToLast_ByHeaderText_ShouldAddToColumns() {
            columns!.Add(HeaderText);
            columns.AddChildToLast(HeaderText, headerStyle, fieldStyleGenerator);

            columns.Should().HaveCount(1);

            DataColumnCollection<Record> lastChildColumns = columns.Last().ChildColumns;
            lastChildColumns.Should().HaveCount(1);

            RecordDataColumn<Record>? column = lastChildColumns.Single() as RecordDataColumn<Record>;

            column.Should().NotBeNull();
            column!.HeaderText.Should().Be(HeaderText);
            column.HeaderStyle.Should().Be(headerStyle);
            column.FieldStyleGenerator.Should().Be(fieldStyleGenerator);
        }

        [Test]
        public void AddChildToLast_ByValueGenerator_ShouldAddToColumns() {
            Func<RecordContext<Record>, object> fieldValueGenerator = x => x.Record.Id;

            columns!.Add(HeaderText);
            columns.AddChildToLast(HeaderText, x => x.UseValue(fieldValueGenerator), headerStyle, fieldStyleGenerator);

            columns.Should().HaveCount(1);

            DataColumnCollection<Record> lastChildColumns = columns.Last().ChildColumns;
            lastChildColumns.Should().HaveCount(1);

            RecordDataColumn<Record>? column = lastChildColumns.Single() as RecordDataColumn<Record>;

            column.Should().NotBeNull();
            column!.HeaderText.Should().Be(HeaderText);
            column.FieldValueGenerator.Should().Be(fieldValueGenerator);
            column.HeaderStyle.Should().Be(headerStyle);
            column.FieldStyleGenerator.Should().Be(fieldStyleGenerator);
        }

        [Test]
        public void AddChildToLast_ByFormulaGenerator_ShouldAddToColumns() {
            Func<RecordContext<Record>, string> fieldFormulaGenerator = x => "SUM(A1 + B1)";

            columns!.Add(HeaderText);
            columns.AddChildToLast(HeaderText, x => x.UseFormula(fieldFormulaGenerator), headerStyle, fieldStyleGenerator);

            columns.Should().HaveCount(1);

            DataColumnCollection<Record> lastChildColumns = columns.Last().ChildColumns;
            lastChildColumns.Should().HaveCount(1);

            RecordDataColumn<Record>? column = lastChildColumns.Single() as RecordDataColumn<Record>;

            column.Should().NotBeNull();
            column!.HeaderText.Should().Be(HeaderText);
            column.FieldFormulaGenerator.Should().Be(fieldFormulaGenerator);
            column.HeaderStyle.Should().Be(headerStyle);
            column.FieldStyleGenerator.Should().Be(fieldStyleGenerator);
        }

        [Test]
        public void AddChildToLast_ByFieldKey_ShouldAddToColumns() {
            columns!.Add(HeaderText);
            columns.AddChildToLast<int>(HeaderText, "Id", headerStyle, fieldStyleGenerator);

            columns.Should().HaveCount(1);

            DataColumnCollection<Record> lastChildColumns = columns.Last().ChildColumns;
            lastChildColumns.Should().HaveCount(1);

            DataColumn<Record, int>? column = lastChildColumns.Single() as DataColumn<Record, int>;

            column.Should().NotBeNull();
            column!.HeaderText.Should().Be(HeaderText);
            column.FieldKey.Should().Be("Id");
            column.HeaderStyle.Should().Be(headerStyle);
            column.FieldStyleGenerator.Should().Be(fieldStyleGenerator);
        }

        [Test]
        public void AddChildToLast_ByFieldKeyExpression_ShouldAddToColumns() {
            Expression<Func<Record, int>> fieldKeyExpression = x => x.Id;

            columns!.Add(HeaderText);
            columns.AddChildToLast(HeaderText, x => x.Id, headerStyle, fieldStyleGenerator);

            columns.Should().HaveCount(1);

            DataColumnCollection<Record> lastChildColumns = columns.Last().ChildColumns;
            lastChildColumns.Should().HaveCount(1);

            DataColumn<Record, int>? column = lastChildColumns.Single() as DataColumn<Record, int>;

            column.Should().NotBeNull();
            column!.HeaderText.Should().Be(HeaderText);
            column.FieldKey.Should().Be("Id");
            column.HeaderStyle.Should().Be(headerStyle);
            column.FieldStyleGenerator.Should().Be(fieldStyleGenerator);
        }

        [Test]
        public void AddChildToLast_ByFieldKeyExpressionWithGenerator_ShouldAddToColumns() {
            Func<FieldContext<Record, int>, object> fieldValueGenerator = x => x.Record.Id;

            columns!.Add(HeaderText);
            columns.AddChildToLast(HeaderText, x => x.Id, x => x.UseValue(fieldValueGenerator), headerStyle, fieldStyleGenerator);

            columns.Should().HaveCount(1);

            DataColumnCollection<Record> lastChildColumns = columns.Last().ChildColumns;
            lastChildColumns.Should().HaveCount(1);

            DataColumn<Record, int>? column = lastChildColumns.Single() as DataColumn<Record, int>;

            column.Should().NotBeNull();
            column!.HeaderText.Should().Be(HeaderText);
            column.FieldKey.Should().Be("Id");
            column.FieldValueGenerator.Should().Be(fieldValueGenerator);
            column.HeaderStyle.Should().Be(headerStyle);
            column.FieldStyleGenerator.Should().Be(fieldStyleGenerator);
        }

        private class Record {
            public int Id { get; set; }

            public string? Name { get; set; }
        }
    }
}
