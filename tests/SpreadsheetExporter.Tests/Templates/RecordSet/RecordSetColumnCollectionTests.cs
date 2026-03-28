using System.Linq.Expressions;
using CloudyWing.SpreadsheetExporter.Templates.RecordSet;

namespace CloudyWing.SpreadsheetExporter.Tests.Templates.RecordSet;

[TestFixture]
internal class RecordSetColumnCollectionTests {
    private RecordSetColumnCollection<Record> columns = null!;

    [SetUp]
    public void SetUp() {
        columns = new RecordSetColumnCollection<Record>(null);
    }

    [Test]
    public void EmptyCollection_ShouldExposeZeroSpans() {
        Assert.Multiple(() => {
            Assert.That(columns.ColumnSpan, Is.EqualTo(0));
            Assert.That(columns.RowSpan, Is.EqualTo(0));
            Assert.That(columns.LeafColumns, Is.Empty);
        });
    }

    [Test]
    public void RootColumns_FromChildCollection_ShouldReturnTopLevelCollection() {
        columns.Add("Profile");

        RecordSetColumnCollection<Record> childColumns = columns.Last().ChildColumns;

        Assert.That(childColumns.RootColumns, Is.SameAs(columns));
    }

    [Test]
    public void Add_ByHeaderText_ShouldCreateGeneratorColumnWithDefaultStyles() {
        columns.Add("Name");

        GeneratorColumn<Record> column = (GeneratorColumn<Record>)columns.Single();
        RecordContext<Record> context = new(0, 0, new Record { Name = "Alice" });

        Assert.Multiple(() => {
            Assert.That(column.HeaderText, Is.EqualTo("Name"));
            Assert.That(column.HeaderStyle, Is.EqualTo(SpreadsheetManager.DefaultCellStyles.HeaderStyle));
            Assert.That(column.GetFieldStyle(context), Is.EqualTo(SpreadsheetManager.DefaultCellStyles.FieldStyle));
        });
    }

    [Test]
    public void Add_WithConfiguredGenerators_ShouldCreateGeneratorColumn() {
        DataValidation validation = new() {
            ValidationType = DataValidationType.List,
            ListItems = ["A", "B"]
        };

        columns.Add("Status", configureGenerators: x => {
            x.UseValue(context => context.Record.Name);
            x.UseDataValidation(_ => validation);
        });

        GeneratorColumn<Record> column = (GeneratorColumn<Record>)columns.Single();
        RecordContext<Record> context = new(1, 2, new Record { Name = "Approved" });

        Assert.Multiple(() => {
            Assert.That(column.GetFieldValue(context), Is.EqualTo("Approved"));
            Assert.That(column.GetFieldDataValidation(context), Is.SameAs(validation));
        });
    }

    [Test]
    public void Add_ByFieldExpression_ShouldCreateFieldColumn() {
        Expression<Func<Record, int>> expression = x => x.Id;

        columns.Add("Id", expression);

        FieldColumn<Record, int> column = (FieldColumn<Record, int>)columns.Single();

        Assert.Multiple(() => {
            Assert.That(column.HeaderText, Is.EqualTo("Id"));
            Assert.That(column.FieldKey, Is.EqualTo("Id"));
            Assert.That(column.HeaderStyle, Is.EqualTo(SpreadsheetManager.DefaultCellStyles.HeaderStyle));
        });
    }

    [Test]
    public void GetLastColumn_WhenCollectionHasItems_ShouldReturnLastColumn() {
        columns.Add("First").Add("Last");

        RecordSetColumnBase<Record> last = columns.GetLastColumn();

        Assert.That(last.HeaderText, Is.EqualTo("Last"));
    }

    [Test]
    public void GetLastColumn_WhenCollectionIsEmpty_ShouldThrowInvalidOperationException() {
        Assert.That(
            () => columns.GetLastColumn(),
            Throws.TypeOf<InvalidOperationException>().And.Message.EqualTo("No columns available.")
        );
    }

    [Test]
    public void Parent_FromChildCollection_ShouldReturnParentCollection() {
        columns.Add("Profile");
        RecordSetColumnCollection<Record> childColumns = columns.GetLastColumn().ChildColumns;

        RecordSetColumnCollection<Record> parent = childColumns.Parent;

        Assert.That(parent, Is.SameAs(columns));
    }

    [Test]
    public void Parent_WhenRootCollection_ShouldThrowInvalidOperationException() {
        Assert.That(
            () => columns.Parent,
            Throws.TypeOf<InvalidOperationException>()
                .And.Message.Contains("root collection")
        );
    }

    [Test]
    public void GetLastColumn_And_ChildColumns_ShouldBuildHierarchyAndLeafColumns() {
        columns
            .Add("Profile")
            .GetLastColumn().ChildColumns
                .Add("Name", (Record x) => x.Name)
                .Add("Age", (Record x) => x.Age)
                .Parent
            .Add("Status");

        RecordSetColumnBase<Record> parent = columns.First();

        Assert.Multiple(() => {
            Assert.That(columns.ColumnSpan, Is.EqualTo(3));
            Assert.That(columns.RowSpan, Is.EqualTo(2));
            Assert.That(parent.ChildColumns, Has.Count.EqualTo(2));
            Assert.That(columns.LeafColumns.Select(x => x.HeaderText), Is.EqualTo(new[] { "Name", "Age", "Status" }));
        });
    }

    [Test]
    public void Add_SameColumnToAnotherCollection_ShouldThrowArgumentException() {
        RecordSetColumnCollection<Record> anotherCollection = new(null);
        GeneratorColumn<Record> column = new() { HeaderText = "Shared" };
        columns.Add(column);

        Assert.That(
            () => anotherCollection.Add(column),
            Throws.TypeOf<ArgumentException>().And.Message.Contains("already contained by another RecordSetColumnCollection")
        );
    }

    private sealed class Record {
        public int Id { get; init; }

        public string? Name { get; init; }

        public int Age { get; init; }
    }
}
