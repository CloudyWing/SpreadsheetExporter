using System.Drawing;
using CloudyWing.SpreadsheetExporter.Templates;

namespace CloudyWing.SpreadsheetExporter.Tests.Templates {
    [TestFixture]
    internal class TemplateContextTests {
        [Test]
        public void Create_ShouldReturnExpectedTemplateContext() {
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

            TemplateContext context = TemplateContext.Create(new[] { template1, template2 });

            Cell fixedCell21 = cell21.ShallowCopy();
            fixedCell21.Point = Point.Add(cell21.Point, new Size(0, template1RowSpan));
            Cell fixedCell22 = cell22.ShallowCopy();
            fixedCell22.Point = Point.Add(cell22.Point, new Size(0, template1RowSpan));

            context.Cells.Count.Should().Be(4);
            context.Cells[0].Should().BeEquivalentTo(cell11);
            context.Cells[1].Should().BeEquivalentTo(cell12);
            context.Cells[2].Should().BeEquivalentTo(fixedCell21);
            context.Cells[3].Should().BeEquivalentTo(fixedCell22);
            context.RowSpan.Should().Be(5);
            context.RowHeights.Count.Should().Be(2);
            context.RowHeights[0].Should().Be(10);
            context.RowHeights[2].Should().Be(20);
        }
    }
}
