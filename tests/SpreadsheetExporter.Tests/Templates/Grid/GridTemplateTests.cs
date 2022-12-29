using System.Drawing;
using CloudyWing.SpreadsheetExporter.Templates;
using CloudyWing.SpreadsheetExporter.Templates.Grid;

namespace CloudyWing.SpreadsheetExporter.Tests.Templates.Grid {
    [TestFixture]
    internal class GridTemplateTests {
        [Test]
        public void CreateRow_ShouldAddRowToTemplate() {
            GridTemplate template = new GridTemplate();
            template.CreateRow();

            template.RowSpan.Should().Be(1);
        }

        [Test]
        public void CreateRow_SetRowHeight_ShouldAddRowWithSpecifiedHeightToTemplate() {
            GridTemplate template = new GridTemplate();
            template.CreateRow(10);

            template.GetContext().RowHeights[0].Should().Be(10);
        }

        [Test]
        public void CreateCell_ShouldAddCellToLastRow() {
            GridTemplate template = new GridTemplate();
            template.CreateRow();
            template.CreateCell("A1");

            template.GetContext().Cells.Count.Should().Be(1);
        }

        [Test]
        public void CreateCell_SetCellValue_ShouldAddCellWithSpecifiedValue() {
            GridTemplate template = new GridTemplate();
            template.CreateRow();
            template.CreateCell("A1");

            template.GetContext().Cells.Single().GetValue().Should().Be("A1");
        }

        [Test]
        public void CreateCell_SetColumnSpan_ShouldAddCellWithSpecifiedColumnSpan() {
            GridTemplate template = new GridTemplate();
            template.CreateRow();
            template.CreateCell("A1", 2);

            template.GetContext().Cells.Single().Size.Width.Should().Be(2);
        }

        [Test]
        public void CreateCell_SetRowSpan_ShouldAddCellWithSpecifiedRowSpan() {
            GridTemplate template = new GridTemplate();
            template.CreateRow();
            template.CreateCell("A1", rowSpan: 2);

            template.GetContext().Cells.Single().Size.Height.Should().Be(2);
        }

        [Test]
        public void CreateCell_SetCellStyle_ShouldAddCellWithSpecifiedStyle() {
            CellStyle cellStyle = new CellStyle(font: new CellFont("Arial", 10));

            GridTemplate template = new GridTemplate();
            template.CreateRow();
            template.CreateCell("A1", cellStyle: cellStyle);

            template.GetContext().Cells.Single().GetCellStyle().Should().BeEquivalentTo(cellStyle);
        }

        [Test]
        public void CreateCell_SetFormula_ShouldCreateCellWithFormula() {
            GridTemplate template = new GridTemplate();
            template.CreateRow();
            template.CreateCell((x, y) => "A1 + B1", 2, 2, SpreadsheetManager.DefaultCellStyles.GridCellStyle);

            TemplateContext context = template.GetContext();
            Cell cell = context.Cells.Single();
            cell.GetFormula().Should().Be("A1 + B1");
            cell.GetCellStyle().Should().Be(SpreadsheetManager.DefaultCellStyles.GridCellStyle);
            cell.Size.Should().Be(new Size(2, 2));
        }

        [Test]
        public void GetContext_ShouldReturnExpectedTemplateContext() {
            GridTemplate template = new GridTemplate()
                .CreateRow(20)
                .CreateCell("A1", columnSpan: 2, rowSpan: 2)
                .CreateRow()
                .CreateCell("B1", columnSpan: 2, rowSpan: 2);

            TemplateContext context = template.GetContext();

            context.Cells.Count.Should().Be(2);
            context.Cells[0].Point.Should().Be(new Point(0, 0));
            context.Cells[0].Size.Should().Be(new Size(2, 2));
            context.Cells[0].GetValue().Should().Be("A1");
            context.Cells[1].Point.Should().Be(new Point(2, 1));
            context.Cells[1].Size.Should().Be(new Size(2, 2));
            context.Cells[1].GetValue().Should().Be("B1");
            context.RowSpan.Should().Be(3);
            context.RowHeights.Count.Should().Be(2);
            context.RowHeights[0].Should().Be(20);
            context.RowHeights[1].Should().Be(16.5);
        }
    }
}
