using CloudyWing.SpreadsheetExporter.Templates;

namespace CloudyWing.SpreadsheetExporter.Tests {
    [TestFixture]
    internal class SheeterContextTests {
        [Test]
        public void Constructor_ShouldInitializeProperties() {
            Sheeter sheeter = new("Sheet1") {
                Password = "password"
            };
            sheeter.PageSettings.PaperSize = PaperSize.A4;
            sheeter.SetColumnWidth(1, 50D);
            sheeter.SetColumnWidth(2, 75D);

            ITemplate template = Substitute.For<ITemplate>();
            template.GetContext().Returns(new TemplateContext(Enumerable.Empty<Cell>(), 0, new Dictionary<int, double?>()));
            sheeter.AddTemplate(template);

            SheeterContext sut = new(sheeter);

            sut.SheetName.Should().Be(sheeter.SheetName);
            sut.Cells.Should().BeEmpty();
            sut.ColumnWidths.Should().Equal(sheeter.ColumnWidths);
            sut.Password.Should().Be(sheeter.Password);
            sut.IsProtected.Should().BeTrue();
            sut.PageSettings.Should().Be(sheeter.PageSettings);
        }
    }
}
