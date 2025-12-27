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

            Assert.That(context.Cells, Has.Count.EqualTo(0));
        }

        [Test]
        public void GetContext_AddTwoRowsWithOnlyColumnEach_ShouldReturnContextWithCells() {
            template!.Columns.Add("Column 1");
            template!.Columns.AddChildToLast("Child Column 1");

            TemplateContext context = template.GetContext();

            Assert.That(context.Cells, Has.Count.EqualTo(records.Count() + template.Columns.RowSpan));
            Assert.That(context.Cells.All(x => x.ValueGenerator != null), Is.True);
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

            Assert.That(context.Cells.Select(x => new { x.Point, x.Size }).ToList(), Is.EqualTo(expectedCells));
            Assert.That(context.Cells.All(x => x.CellStyleGenerator != null), Is.True);
            Assert.That(context.Cells.All(x => x.ValueGenerator != null), Is.True);
        }

        [Test]
        public void GetContext_SetDataSource_ShouldReturnTemplateContextWithCorrectRowSpan() {
            Assert.That(template!.GetContext().RowSpan, Is.EqualTo(records.Count()));
        }

        [Test]
        public void GetContext_SetHeight_ShouldReturnTemplateContextWithCorrectRowHeights() {
            template!.HeaderHeight = 10d;
            template.RecordHeight = 11d;
            template.Columns.Add("Header Text");

            TemplateContext context = template.GetContext();

            Assert.That(context.RowHeights[0], Is.EqualTo(template.HeaderHeight));

            for (int i = 1; i < context.RowHeights.Count; i++) {
                Assert.That(context.RowHeights[i], Is.EqualTo(template.RecordHeight));
            }
        }

        [Test]
        public void Add_SetHeaderTextAndHeaderStyle_ShouldCreateRecordDataColumnWithHeaderTextAndHeaderStyle() {
            string headerText = "Column 1";
            CellStyle headerStyle = new();

            template!.Columns.Add(headerText, headerStyle);

            Assert.That(template.Columns[0].HeaderText, Is.EqualTo(headerText));
            Assert.That(template.Columns[0].HeaderStyle, Is.EqualTo(headerStyle));
        }

        [Test]
        public void ColumnSpan_AddTwoRowsButOneHasColumnLayersTwo_ShouldReturnTwo() {
            template!.Columns.Add("Column 1")
                .AddChildToLast("Child Column1")
                .AddChildToLast("Child Column2");

            Assert.That(template.ColumnSpan, Is.EqualTo(2));
        }

        [Test]
        public void RowSpan_AddTwoRowsButOneHasColumnLayersTwo_ShouldReturnTwo() {
            template!.Columns.Add("Column 1")
                .AddChildToLast("Child Column1")
                .AddChildToLast("Child Column2");

            Assert.That(template.RowSpan, Is.EqualTo(2 + records.Count()));
        }

        [Test]
        public void GetContext_IsFreezeHeaderIsTrue_ShouldSetFreezePanesToHeaderRowCount() {
            template!.Columns.Add("Column 1")
                .AddChildToLast("Child Column1")
                .AddChildToLast("Child Column2");
            template.IsFreezeHeader = true;

            TemplateContext context = template.GetContext();

            Assert.That(context.FreezePanes, Is.Not.Null);
            Assert.That(context.FreezePanes!.Value.X, Is.EqualTo(0));
            Assert.That(context.FreezePanes.Value.Y, Is.EqualTo(template.Columns.RowSpan));
        }

        [Test]
        public void GetContext_IsFreezeHeaderIsFalse_ShouldNotSetFreezePanes() {
            template!.Columns.Add("Column 1");
            template.IsFreezeHeader = false;

            TemplateContext context = template.GetContext();

            Assert.That(context.FreezePanes, Is.Null);
        }

        [Test]
        public void GetContext_IsAutoFilterEnabledIsTrue_ShouldSetAutoFilterEnabled() {
            template!.Columns.Add("Column 1");
            template.IsAutoFilterEnabled = true;

            TemplateContext context = template.GetContext();

            Assert.That(context.IsAutoFilterEnabled, Is.True);
        }

        [Test]
        public void GetContext_IsAutoFilterEnabledIsFalse_ShouldNotSetAutoFilterEnabled() {
            template!.Columns.Add("Column 1");
            template.IsAutoFilterEnabled = false;

            TemplateContext context = template.GetContext();

            Assert.That(context.IsAutoFilterEnabled, Is.False);
        }

        [Test]
        public void DataSource_WhenEnumeratedMultipleTimes_ShouldNotReEnumerateOriginalSource() {
            int enumerationCount = 0;
            IEnumerable<Record> trackingEnumerable = GetTrackingEnumerable();

            template = new RecordSetTemplate<Record>(trackingEnumerable);
            template.Columns.Add("Column 1");

            template.GetContext();

            Assert.That(enumerationCount, Is.EqualTo(1), "DataSource should only be enumerated once due to caching");

            IEnumerable<Record> GetTrackingEnumerable() {
                enumerationCount++;
                foreach (Record record in records) {
                    yield return record;
                }
            }
        }

        [Test]
        public void DataSource_WhenSetToReadOnlyList_ShouldUseDirectlyWithoutConversion() {
            IReadOnlyList<Record> readOnlyList = records.ToList().AsReadOnly();

            template = new RecordSetTemplate<Record>(readOnlyList);
            template.Columns.Add("Column 1");

            TemplateContext context = template.GetContext();

            Assert.That(context.RowSpan, Is.EqualTo(readOnlyList.Count + template.Columns.RowSpan));
        }

        [Test]
        public void DataSource_WhenChangedAfterCaching_ShouldReEnumerateNewSource() {
            List<Record> firstSource = [new Record { Id = 1 }];
            List<Record> secondSource = [new Record { Id = 2 }, new Record { Id = 3 }];

            template = new RecordSetTemplate<Record>(firstSource);
            template.Columns.Add("Column 1");

            TemplateContext context1 = template.GetContext();
            Assert.That(context1.RowSpan, Is.EqualTo(firstSource.Count + template.Columns.RowSpan));

            template.DataSource = secondSource;

            TemplateContext context2 = template.GetContext();
            Assert.That(context2.RowSpan, Is.EqualTo(secondSource.Count + template.Columns.RowSpan));
        }

        private class Record {
            public int Id { get; set; }
        }
    }
}

