using System.Drawing;
using CloudyWing.SpreadsheetExporter.Templates;

namespace CloudyWing.SpreadsheetExporter.Tests.Templates {
    [TestFixture]
    internal class TemplateLayoutTests {
        [Test]
        public void Create_ShouldReturnExpectedTemplateLayout() {
            Cell cell11 = new() {
                Point = new Point(0, 0),
                Size = new Size(2, 2),
                ValueGenerator = (x, y) => "A1"
            };
            Cell cell12 = new() {
                Point = new Point(0, 2),
                Size = new Size(2, 2),
                ValueGenerator = (x, y) => "C1"
            };
            int template1RowSpan = 2;
            ISheetTemplate template1 = Substitute.For<ISheetTemplate>();
            template1.GetLayout().Returns(new TemplateLayout(new[] { cell11, cell12 }, template1RowSpan, new Dictionary<int, double?> { { 0, 10 } }));

            Cell cell21 = new() {
                Point = new Point(0, 0),
                Size = new Size(3, 3),
                ValueGenerator = (x, y) => "A3"
            };
            Cell cell22 = new() {
                Point = new Point(0, 0),
                Size = new Size(3, 3),
                ValueGenerator = (x, y) => "D3"
            };
            ISheetTemplate template2 = Substitute.For<ISheetTemplate>();
            template2.GetLayout().Returns(new TemplateLayout(new[] { cell21, cell22 }, 3, new Dictionary<int, double?> { { 0, 20 } }));

            TemplateLayout layout = TemplateLayout.Create(new[] { template1, template2 });

            Cell fixedCell21 = cell21.ShallowCopy();
            fixedCell21.Point = Point.Add(cell21.Point, new Size(0, template1RowSpan));
            Cell fixedCell22 = cell22.ShallowCopy();
            fixedCell22.Point = Point.Add(cell22.Point, new Size(0, template1RowSpan));

            Assert.Multiple(() => {
                Assert.That(layout.Cells.Count, Is.EqualTo(4));
                Assert.That(layout.Cells[0].Point, Is.EqualTo(cell11.Point));
                Assert.That(layout.Cells[0].Size, Is.EqualTo(cell11.Size));
                Assert.That(layout.Cells[1].Point, Is.EqualTo(cell12.Point));
                Assert.That(layout.Cells[1].Size, Is.EqualTo(cell12.Size));
                Assert.That(layout.Cells[2].Point, Is.EqualTo(fixedCell21.Point));
                Assert.That(layout.Cells[2].Size, Is.EqualTo(fixedCell21.Size));
                Assert.That(layout.Cells[3].Point, Is.EqualTo(fixedCell22.Point));
                Assert.That(layout.Cells[3].Size, Is.EqualTo(fixedCell22.Size));
                Assert.That(layout.RowSpan, Is.EqualTo(5));
                Assert.That(layout.RowHeights.Count, Is.EqualTo(2));
                Assert.That(layout.RowHeights[0], Is.EqualTo(10));
                Assert.That(layout.RowHeights[2], Is.EqualTo(20));
            });
        }

    }
}
