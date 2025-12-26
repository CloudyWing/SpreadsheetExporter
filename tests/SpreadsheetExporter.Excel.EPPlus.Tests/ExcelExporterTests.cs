using CloudyWing.SpreadsheetExporter.Templates.Grid;
using OfficeOpenXml;

namespace CloudyWing.SpreadsheetExporter.Excel.EPPlus.Tests {
    public class Tests {
        [Test]
        public void ContentType_ShouldReturnExpectedContentType() {
            ExcelExporter exporter = new();

            Assert.That(exporter.ContentType, Is.EqualTo("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"));
        }

        [Test]
        public void FileNameExtension_ShouldReturnExpectedFileNameExtension() {
            ExcelExporter exporter = new();

            Assert.That(exporter.FileNameExtension, Is.EqualTo(".xlsx"));
        }

        [Test]
        public void Export_ShouldReturnExpectedByteArray() {
            ExcelExporter exporter = new();
            exporter.CreateSheeter();

            Assert.That(exporter.Export(), Is.Not.Empty);
        }

        [Test]
        public void Export_SheetCreatedEvent_SheetInfoIsSetCorrectly() {
            ExcelExporter exporter = new();
            Sheeter sheeter = exporter.CreateSheeter("Sheet Name");
            GridTemplate template = new();
            string cellValue = "Cell Value";
            template.CreateRow().CreateCell(cellValue);
            sheeter.AddTemplate(template);
            sheeter.SetColumnWidth(0, 100);

            exporter.SheetCreatedEvent += (sender, args) => {
                ExcelWorksheet? sheet = args.SheetObject as ExcelWorksheet;
                SheeterContext sheeterContext = args.SheeterContext;

                Assert.That(sheeterContext.SheetName, Is.EqualTo(sheeter.SheetName), "SheeterContext sheet name should match.");
                Assert.That(sheeterContext.DefaultRowHeight, Is.EqualTo(sheeter.DefaultRowHeight), "SheeterContext default row height should match.");
                Assert.That(sheet, Is.Not.Null, "sheet should be created.");
                Assert.That(sheet!.Name, Is.EqualTo(sheeter.SheetName), "sheet name should match.");
                Assert.That(sheet.Column(1).Width, Is.EqualTo((short)sheeter.ColumnWidths[0]), "column width should match.");
                Assert.That(sheet.GetValue(1, 1), Is.EqualTo(cellValue), "cell value should match.");
            };

            exporter.Export();
        }
    }
}
