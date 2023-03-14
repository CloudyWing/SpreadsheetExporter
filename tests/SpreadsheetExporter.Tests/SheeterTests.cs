using CloudyWing.SpreadsheetExporter.Templates;

namespace CloudyWing.SpreadsheetExporter.Tests {
    [TestFixture]
    internal class SheeterTests {
        [Test]
        public void Constructor_ShouldInitializeProperties() {
            string sheetName = "Sheet1";
            Sheeter sheeter = new Sheeter(sheetName);

            sheeter.SheetName.Should().Be(sheetName);
            sheeter.Password.Should().BeNull();
            sheeter.PageSettings.Should().BeEquivalentTo(new PageSettings());
        }

        [Test]
        public void SetColumnWidth_ShouldSetColumnWidthToColumnWidths() {
            Sheeter sheeter = new Sheeter("Sheet1");
            int columnIndex = 1;
            double columnWidth = 100;

            sheeter.SetColumnWidth(columnIndex, columnWidth);

            sheeter.ColumnWidths[columnIndex].Should().Be(columnWidth);
        }

        [Test]
        public void AddTemplate_ShouldAddTemplateToList() {
            Sheeter sheeter = new Sheeter("Sheet1");
            ITemplate template = Substitute.For<ITemplate>();

            sheeter.AddTemplate(template);

            sheeter.Templates.Should().Contain(template);
        }
    }

}
