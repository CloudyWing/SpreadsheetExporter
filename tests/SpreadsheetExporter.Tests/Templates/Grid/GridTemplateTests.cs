using System.Drawing;
using CloudyWing.SpreadsheetExporter.Templates;
using CloudyWing.SpreadsheetExporter.Templates.Grid;

namespace CloudyWing.SpreadsheetExporter.Tests.Templates.Grid {
    [TestFixture]
    internal class GridTemplateTests {
        [Test]
        public void CreateRow_ShouldAddRowToTemplate() {
            GridTemplate template = new();
            template.CreateRow();

            Assert.That(template.RowSpan, Is.EqualTo(1));
        }

        [Test]
        public void CreateRow_SetRowHeight_ShouldAddRowWithSpecifiedHeightToTemplate() {
            GridTemplate template = new();
            template.CreateRow(10);

            Assert.That(template.GetContext().RowHeights[0], Is.EqualTo(10));
        }

        [Test]
        public void CreateCell_ShouldAddCellToLastRow() {
            GridTemplate template = new();
            template.CreateRow()
                .CreateCell("A1");

            Assert.That(template.GetContext().Cells.Count, Is.EqualTo(1));
        }

        [Test]
        public void CreateCell_SetCellValue_ShouldAddCellWithSpecifiedValue() {
            GridTemplate template = new();
            template.CreateRow()
                .CreateCell("A1");

            Assert.That(template.GetContext().Cells.Single().GetValue(), Is.EqualTo("A1"));
        }

        [Test]
        public void CreateCell_SetColumnSpan_ShouldAddCellWithSpecifiedColumnSpan() {
            GridTemplate template = new();
            template.CreateRow()
                .CreateCell("A1", 2);

            Assert.That(template.GetContext().Cells.Single().Size.Width, Is.EqualTo(2));
        }

        [Test]
        public void CreateCell_SetRowSpan_ShouldAddCellWithSpecifiedRowSpan() {
            GridTemplate template = new();
            template.CreateRow()
                .CreateCell("A1", rowSpan: 2);

            Assert.That(template.GetContext().Cells.Single().Size.Height, Is.EqualTo(2));
        }

        [Test]
        public void CreateCell_SetCellStyle_ShouldAddCellWithSpecifiedStyle() {
            CellStyle cellStyle = new(Font: new CellFont("Arial", 10));

            GridTemplate template = new();
            template.CreateRow()
                .CreateCell("A1", cellStyle: cellStyle);

            Assert.That(template.GetContext().Cells.Single().GetCellStyle(), Is.EqualTo(cellStyle));
        }

        [Test]
        public void CreateCell_SetFormula_ShouldCreateCellWithFormula() {
            GridTemplate template = new();
            template.CreateRow()
                .CreateCell((x, y) => "A1 + B1", 2, 2, SpreadsheetManager.DefaultCellStyles.GridCellStyle);

            TemplateContext context = template.GetContext();
            Cell cell = context.Cells.Single();
            Assert.That(cell.GetFormula(), Is.EqualTo("A1 + B1"));
            Assert.That(cell.GetCellStyle(), Is.EqualTo(SpreadsheetManager.DefaultCellStyles.GridCellStyle));
            Assert.That(cell.Size, Is.EqualTo(new Size(2, 2)));
        }

        [Test]
        public void GetContext_ShouldReturnExpectedTemplateContext() {
            GridTemplate template = new GridTemplate()
                .CreateRow(20)
                .CreateCell("A1", columnSpan: 2, rowSpan: 2)
                .CreateRow()
                .CreateCell("B1", columnSpan: 2, rowSpan: 2);

            TemplateContext context = template.GetContext();

            Assert.That(context.Cells.Count, Is.EqualTo(2));
            Assert.That(context.Cells[0].Point, Is.EqualTo(new Point(0, 0)));
            Assert.That(context.Cells[0].Size, Is.EqualTo(new Size(2, 2)));
            Assert.That(context.Cells[0].GetValue(), Is.EqualTo("A1"));
            Assert.That(context.Cells[1].Point, Is.EqualTo(new Point(2, 1)));
            Assert.That(context.Cells[1].Size, Is.EqualTo(new Size(2, 2)));
            Assert.That(context.Cells[1].GetValue(), Is.EqualTo("B1"));
            Assert.That(context.RowSpan, Is.EqualTo(3));
            Assert.That(context.RowHeights.Count, Is.EqualTo(2));
            Assert.That(context.RowHeights[0], Is.EqualTo(20));
            Assert.That(context.RowHeights[1], Is.EqualTo(null));
        }
    }
}
