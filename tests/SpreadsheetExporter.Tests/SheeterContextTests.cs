using System.Drawing;
using CloudyWing.SpreadsheetExporter.Templates;

namespace CloudyWing.SpreadsheetExporter.Tests {
    [TestFixture]
    internal class SheeterContextTests {
        [Test]
        public void Constructor_ShouldInitializeProperties() {
            Sheeter sheeter = new Sheeter("Sheet1") {
                Password = "password"
            };
            sheeter.PageSettings.PaperSize = PaperSize.A4;
            sheeter.SetColumnWidth(1, 50D);
            sheeter.SetColumnWidth(2, 75D);

            ITemplate template = Substitute.For<ITemplate>();
            template.GetContext().Returns(new TemplateContext(Enumerable.Empty<Cell>(), 0, new Dictionary<int, double>()));
            sheeter.AddTemplate(template);

            sheeter.Watermark = new Bitmap(100, 100);

            SheeterContext sut = new SheeterContext(sheeter);

            sut.SheetName.Should().Be(sheeter.SheetName);
            sut.Cells.Should().BeEmpty();
            sut.ColumnWidths.Should().Equal(sheeter.ColumnWidths);
            sut.Password.Should().Be(sheeter.Password);
            sut.IsProtected.Should().BeTrue();
            sut.PageSettings.Should().Be(sheeter.PageSettings);
            sut.HasWatermark.Should().BeTrue();
            sut.Watermark.Should().NotBeNull();
        }

        [Test]
        public void Watermark_ShouldReturnCorrectImage() {
            Sheeter sheeter = new Sheeter("Sheet1");
            sheeter.PageSettings.PaperSize = PaperSize.A4;
            sheeter.Watermark = new Bitmap(50, 50);

            SheeterContext context = new SheeterContext(sheeter);
            Image watermark = context.Watermark;

            watermark.Should().NotBeNull();
            watermark.Width.Should().Be(PaperSize.A4.Width);
            watermark.Height.Should().Be(PaperSize.A4.Height);
        }
    }
}
