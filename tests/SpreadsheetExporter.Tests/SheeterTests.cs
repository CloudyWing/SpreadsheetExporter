using CloudyWing.SpreadsheetExporter.Templates;

namespace CloudyWing.SpreadsheetExporter.Tests {
    [TestFixture]
    internal class SheeterTests {
        [Test]
        public void Constructor_ShouldInitializeProperties() {
            string sheetName = "Sheet1";
            Sheeter sheeter = new(sheetName);

            Assert.That(sheeter.SheetName, Is.EqualTo(sheetName));
            Assert.That(sheeter.Password, Is.Null);
            Assert.That(sheeter.PageSettings, Is.Not.Null);
        }

        [Test]
        public void SetColumnWidth_ShouldSetColumnWidthToColumnWidths() {
            Sheeter sheeter = new("Sheet1");
            int columnIndex = 1;
            double columnWidth = 100;

            sheeter.SetColumnWidth(columnIndex, columnWidth);

            Assert.That(sheeter.ColumnWidths[columnIndex], Is.EqualTo(columnWidth));
        }

        [Test]
        public void AddTemplate_ShouldAddTemplateToList() {
            Sheeter sheeter = new("Sheet1");
            ITemplate template = Substitute.For<ITemplate>();

            sheeter.AddTemplate(template);

            Assert.That(sheeter.Templates, Does.Contain(template));
        }
    }

}
