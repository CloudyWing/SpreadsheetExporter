using System.Drawing;
using CloudyWing.SpreadsheetExporter.Templates;

namespace CloudyWing.SpreadsheetExporter.Tests.Templates {
    [TestFixture]
    internal class MergedTemplateTests {
        [Test]
        public void GetContext_ShouldReturnExpectedTemplateContext() {
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
            ITemplate template1 = Substitute.For<ITemplate>();
            template1.GetContext().Returns(new TemplateContext(new[] { cell11, cell12 }, template1RowSpan, new Dictionary<int, double?> { { 0, 10 } }));

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
            ITemplate template2 = Substitute.For<ITemplate>();
            template2.GetContext().Returns(new TemplateContext(new[] { cell21, cell22 }, 3, new Dictionary<int, double?> { { 0, 20 } }));

            MergedTemplate mergedTemplate = new(template1, template2);

            TemplateContext context = mergedTemplate.GetContext();

            Cell fixedCell21 = cell21.ShallowCopy();
            fixedCell21.Point = Point.Add(cell21.Point, new Size(0, template1RowSpan));
            Cell fixedCell22 = cell22.ShallowCopy();
            fixedCell22.Point = Point.Add(cell22.Point, new Size(0, template1RowSpan));

            Assert.That(context.Cells.Count, Is.EqualTo(4));
            Assert.That(context.Cells[0].Point, Is.EqualTo(cell11.Point));
            Assert.That(context.Cells[0].Size, Is.EqualTo(cell11.Size));
            Assert.That(context.Cells[1].Point, Is.EqualTo(cell12.Point));
            Assert.That(context.Cells[1].Size, Is.EqualTo(cell12.Size));
            Assert.That(context.Cells[2].Point, Is.EqualTo(fixedCell21.Point));
            Assert.That(context.Cells[2].Size, Is.EqualTo(fixedCell21.Size));
            Assert.That(context.Cells[3].Point, Is.EqualTo(fixedCell22.Point));
            Assert.That(context.Cells[3].Size, Is.EqualTo(fixedCell22.Size));
            Assert.That(context.RowSpan, Is.EqualTo(5));
            Assert.That(context.RowHeights.Count, Is.EqualTo(2));
            Assert.That(context.RowHeights[0], Is.EqualTo(10));
            Assert.That(context.RowHeights[2], Is.EqualTo(20));
        }
    }
}
