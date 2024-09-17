using System.Drawing;
using CloudyWing.SpreadsheetExporter.Templates;
using CloudyWing.SpreadsheetExporter.Templates.RecordSet;

namespace CloudyWing.SpreadsheetExporter.Tests.Templates.RecordSet {
    [TestFixture]
    internal class RecordSetTemplateTests {
        private static readonly IEnumerable<Record> records = new Record[] {
            new() { Id = 0 },
            new() { Id = 1 },
            new() { Id = 2 },
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

            TemplateContext context = template.GetContext();

            context.Cells.Should().HaveCount(records.Count() + template.Columns.RowSpan);
            context.Cells.Should().OnlyContain(x => x.ValueGenerator != null);
        }

        [Test]
        public void GetContext_AfterAddingMultipleColumns_ShouldReturnContextWithCorrectCells() {
            template!.Columns.Add("個人資料")
                .AddChildToLast("姓名")
                .AddChildToLast("年齡")
                .Add("聯絡資訊")
                .AddChildToLast("信箱")
                .AddChildToLast("電話相關");

            DataColumnCollection<Record> phoneColumn = template.Columns.Last().ChildColumns.Last().ChildColumns;
            phoneColumn
                .Add("市話")
                .AddChildToLast("電話")
                .AddChildToLast("分機")
                .Add("手機");

            template.Columns.Add("日期相關")
                .AddChildToLast("日期")
                .AddChildToLast("時間")
                .Add("是否啟用")
                .Add("金額")
                .Add("描述");
            var expectedCells = new List<Cell> {
                new() { Point = new Point(0, 0), Size = new Size(2, 3) }, // 個人資料
                new() { Point = new Point(0, 3), Size = new Size(1, 1) }, // 姓名
                new() { Point = new Point(1, 3), Size = new Size(1, 1) }, // 年齡
                new() { Point = new Point(2, 0), Size = new Size(4, 1) }, // 聯絡資訊
                new() { Point = new Point(2, 1), Size = new Size(1, 3) }, // 信箱
                new() { Point = new Point(3, 1), Size = new Size(3, 1) }, // 電話相關
                new() { Point = new Point(3, 2), Size = new Size(2, 1) }, // 市話
                new() { Point = new Point(3, 3), Size = new Size(1, 1) }, // 電話
                new() { Point = new Point(4, 3), Size = new Size(1, 1) }, // 分機
                new() { Point = new Point(5, 2), Size = new Size(1, 2) }, // 手機
                new() { Point = new Point(6, 0), Size = new Size(2, 3) }, // 日期相關
                new() { Point = new Point(6, 3), Size = new Size(1, 1) }, // 日期
                new() { Point = new Point(7, 3), Size = new Size(1, 1) }, // 時間
                new() { Point = new Point(8, 0), Size = new Size(1, 4) }, // 是否啟用
                new() { Point = new Point(9, 0), Size = new Size(1, 4) }, // 金額
                new() { Point = new Point(10, 0), Size = new Size(1, 4) } // 描述
            }
            .Select(x => new { x.Point, x.Size })
            .ToList();

            int row = 4;
            foreach (Record record in records) {
                for (int column = 0; column < 11; column++) {
                    expectedCells.Add(new {
                        Point = new Point(column, row),
                        Size = new Size(1, 1)
                    });
                }
                row++;
            }

            TemplateContext context = template.GetContext();

            context.Cells.Select(x => new { x.Point, x.Size }).ToList()
                .Should().BeEquivalentTo(expectedCells);
            context.Cells.Should().OnlyContain(x => x.CellStyleGenerator != null);
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
            CellStyle headerStyle = new();

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

