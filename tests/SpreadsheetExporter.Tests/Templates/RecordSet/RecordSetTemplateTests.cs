using System.Drawing;
using CloudyWing.SpreadsheetExporter.Templates;
using CloudyWing.SpreadsheetExporter.Templates.RecordSet;

namespace CloudyWing.SpreadsheetExporter.Tests.Templates.RecordSet {
    [TestFixture]
    internal class RecordSetTemplateTests {
        private static readonly IEnumerable<Record> Records = [
            new() { Id = 0 },
            new() { Id = 1 },
            new() { Id = 2 },
        ];
        private RecordSetTemplate<Record> template = null!;

        [SetUp]
        public void SetUp() {
            template = new RecordSetTemplate<Record>(Records);
        }

        [Test]
        public void GetLayout_NoColumns_ShouldReturnLayoutWithEmptyCells() {
            TemplateLayout layout = template.GetLayout();

            Assert.That(layout.Cells, Has.Count.EqualTo(0));
        }

        [Test]
        public void GetLayout_AddTwoRowsWithOnlyColumnEach_ShouldReturnLayoutWithCells() {
            template.Columns.Add("Column 1");
            template.Columns.GetLastColumn().ChildColumns.Add("Child Column 1");

            TemplateLayout layout = template.GetLayout();

            Assert.That(layout.Cells, Has.Count.EqualTo(Records.Count() + template.Columns.RowSpan));
            Assert.That(layout.Cells.All(x => x.ValueGenerator != null), Is.True);
        }

        [Test]
        public void GetLayout_AfterAddingMultipleColumns_ShouldReturnLayoutWithCorrectCells() {
            template.Columns
                .Add("個人資料")
                .GetLastColumn().ChildColumns
                    .Add("姓名")
                    .Add("年齡")
                    .Parent
                .Add("聯絡資訊")
                .GetLastColumn().ChildColumns
                    .Add("信箱")
                    .Add("電話相關")
                    .GetLastColumn().ChildColumns
                        .Add("市話")
                        .GetLastColumn().ChildColumns
                            .Add("電話")
                            .Add("分機")
                            .Parent
                        .Add("手機")
                        .Parent
                    .Parent
                .Add("日期相關")
                .GetLastColumn().ChildColumns
                    .Add("日期")
                    .Add("時間")
                    .Parent
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
            foreach (Record record in Records) {
                for (int column = 0; column < 11; column++) {
                    expectedCells.Add(new {
                        Point = new Point(column, row),
                        Size = new Size(1, 1)
                    });
                }
                row++;
            }

            TemplateLayout layout = template.GetLayout();

            Assert.That(layout.Cells.Select(x => new { x.Point, x.Size }).ToList(), Is.EqualTo(expectedCells));
            Assert.That(layout.Cells.All(x => x.CellStyleGenerator != null), Is.True);
            Assert.That(layout.Cells.All(x => x.ValueGenerator != null), Is.True);
        }

        [Test]
        public void GetLayout_SetDataSource_ShouldReturnLayoutWithCorrectRowSpan() {
            Assert.That(template.GetLayout().RowSpan, Is.EqualTo(Records.Count()));
        }

        [Test]
        public void GetLayout_SetHeight_ShouldReturnLayoutWithCorrectRowHeights() {
            template.HeaderHeight = 10d;
            template.RecordHeight = 11d;
            template.Columns.Add("Header Text");

            TemplateLayout layout = template.GetLayout();

            Assert.Multiple(() => {
                Assert.That(layout.RowHeights[0], Is.EqualTo(template.HeaderHeight));

                for (int i = 1; i < layout.RowHeights.Count; i++) {
                    Assert.That(layout.RowHeights[i], Is.EqualTo(template.RecordHeight));
                }
            });
        }

        [Test]
        public void Add_SetHeaderTextAndHeaderStyle_ShouldCreateRecordDataColumnWithHeaderTextAndHeaderStyle() {
            string headerText = "Column 1";
            CellStyle headerStyle = new();

            template.Columns.Add(headerText, headerStyle);

            Assert.That(template.Columns[0].HeaderText, Is.EqualTo(headerText));
            Assert.That(template.Columns[0].HeaderStyle, Is.EqualTo(headerStyle));
        }

        [Test]
        public void ColumnSpan_AddTwoRowsButOneHasColumnLayersTwo_ShouldReturnTwo() {
            template.Columns
                .Add("Column 1")
                .GetLastColumn().ChildColumns
                    .Add("Child Column1")
                    .Add("Child Column2");

            Assert.That(template.ColumnSpan, Is.EqualTo(2));
        }

        [Test]
        public void RowSpan_AddTwoRowsButOneHasColumnLayersTwo_ShouldReturnTwo() {
            template.Columns
                .Add("Column 1")
                .GetLastColumn().ChildColumns
                    .Add("Child Column1")
                    .Add("Child Column2");

            Assert.That(template.RowSpan, Is.EqualTo(2 + Records.Count()));
        }

        [Test]
        public void DataSource_WhenEnumeratedMultipleTimes_ShouldNotReEnumerateOriginalSource() {
            int enumerationCount = 0;
            IEnumerable<Record> trackingEnumerable = GetTrackingEnumerable();

            template = new RecordSetTemplate<Record>(trackingEnumerable);
            template.Columns.Add("Column 1");

            template.GetLayout();

            Assert.That(enumerationCount, Is.EqualTo(1), "DataSource should only be enumerated once due to caching");

            IEnumerable<Record> GetTrackingEnumerable() {
                enumerationCount++;
                foreach (Record record in Records) {
                    yield return record;
                }
            }
        }

        [Test]
        public void DataSource_WhenSetToReadOnlyList_ShouldUseDirectlyWithoutConversion() {
            IReadOnlyList<Record> readOnlyList = Records.ToList().AsReadOnly();

            template = new RecordSetTemplate<Record>(readOnlyList);
            template.Columns.Add("Column 1");

            TemplateLayout layout = template.GetLayout();

            Assert.That(layout.RowSpan, Is.EqualTo(readOnlyList.Count + template.Columns.RowSpan));
        }

        [Test]
        public void DataSource_WhenChangedAfterCaching_ShouldReEnumerateNewSource() {
            List<Record> firstSource = [new Record { Id = 1 }];
            List<Record> secondSource = [new Record { Id = 2 }, new Record { Id = 3 }];

            template = new RecordSetTemplate<Record>(firstSource);
            template.Columns.Add("Column 1");

            TemplateLayout layout1 = template.GetLayout();
            Assert.That(layout1.RowSpan, Is.EqualTo(firstSource.Count + template.Columns.RowSpan));

            template.DataSource = secondSource;

            TemplateLayout layout2 = template.GetLayout();
            Assert.That(layout2.RowSpan, Is.EqualTo(secondSource.Count + template.Columns.RowSpan));
        }

        private class Record {
            public int Id { get; set; }
        }
    }
}
