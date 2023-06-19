using CloudyWing.SpreadsheetExporter.Templates;
using CloudyWing.SpreadsheetExporter.Templates.RecordSet;

namespace CloudyWing.SpreadsheetExporter.Tests.Templates.RecordSet {
    [TestFixture]
    internal class RecordSetTemplateTests {
        private static readonly IEnumerable<Record> records = new Record[] {
            new Record { Id = 0 },
            new Record { Id = 1 },
            new Record { Id = 2 },
        };
        private RecordSetTemplate<Record>? template;

        [SetUp]
        public void SetUp() {
            template = new RecordSetTemplate<Record>(records);
        }

        [Test]
        public void GetContext_NoColumns_ShouldReturnContextWithEmptyCells() {
            TemplateContext context = template!.GetContext();

            context.Cells.Should().HaveCount(0);
        }

        [Test]
        public void GetContext_AddTwoRowsWithOnlyColumnEach_ShouldReturnContextWithCells() {
            template!.Columns.Add("Column 1");
            template!.Columns.AddChildToLast("Child Column 1");

            var context = template.GetContext();

            context.Cells.Should().HaveCount(records.Count() + template.Columns.RowSpan);
            context.Cells.Should().OnlyContain(x => x.ValueGenerator != null);
        }

        [Test]
        public void GetContext_AddTwoColumns_ShouldReturnContextWithCells() {
            template!.Columns.Add("Column 1")
                .AddChildToLast("Column 2")
                .AddChildToLast("Column 3");

            var context = template.GetContext();

            context.Cells.Should().HaveCount((template.ColumnSpan * (records.Count() + 1)) + 1);
            context.Cells.Should().OnlyContain(x => x.ValueGenerator != null);
        }

        [Test]
        public void GetContext_SetDataSource_ShouldReturnTemplateContextWithCorrectRowSpan() {
            template!.GetContext().RowSpan.Should().Be(records.Count());
        }

        [Test]
        public void GetContext_SetHeight_ShouldReturnTemplateContextWithCorrectRowHeights() {
            template!.HeaderHeight = 10d;
            template.RecordHeight = 11d;
            template.Columns.Add("Header Text");

            TemplateContext context = template.GetContext();

            context.RowHeights[0].Should().Be(template.HeaderHeight);

            for (int i = 1; i < context.RowHeights.Count; i++) {
                context.RowHeights[i].Should().Be(template.RecordHeight);
            }
        }

        [Test]
        public void Add_SetHeaderTextAndHeaderStyle_ShouldCreateRecordDataColumnWithHeaderTextAndHeaderStyle() {
            string headerText = "Column 1";
            CellStyle headerStyle = new CellStyle();

            template!.Columns.Add(headerText, headerStyle);

            template.Columns[0].HeaderText.Should().Be(headerText);
            template.Columns[0].HeaderStyle.Should().Be(headerStyle);
        }

        [Test]
        public void ColumnSpan_AddTwoRowsButOneHasColumnLayersTwo_ShouldReturnTwo() {
            template!.Columns.Add("Column 1")
                .AddChildToLast("Child Column1")
                .AddChildToLast("Child Column2");

            template.ColumnSpan.Should().Be(2);
        }

        [Test]
        public void RowSpan_AddTwoRowsButOneHasColumnLayersTwo_ShouldReturnTwo() {
            template!.Columns.Add("Column 1")
                .AddChildToLast("Child Column1")
                .AddChildToLast("Child Column2");

            template.RowSpan.Should().Be(2 + records.Count());
        }

        private class Record {
            public int Id { get; set; }
        }
    }
}

