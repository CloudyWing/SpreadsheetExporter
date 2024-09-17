using CloudyWing.SpreadsheetExporter.Templates.Grid;
using OfficeOpenXml;

namespace CloudyWing.SpreadsheetExporter.Excel.EPPlus.Tests {
    public class Tests {
        [Test]
        public void ContentType_ShouldReturnExpectedContentType() {
            ExcelExporter exporter = new();

            exporter.ContentType.Should().Be("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
        }

        [Test]
        public void FileNameExtension_ShouldReturnExpectedFileNameExtension() {
            ExcelExporter exporter = new();

            exporter.FileNameExtension.Should().Be(".xlsx");
        }

        [Test]
        public void Export_ShouldReturnExpectedByteArray() {
            ExcelExporter exporter = new();
            exporter.CreateSheeter();

            exporter.Export().Should().NotBeEmpty();
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

                sheeterContext.Should().BeEquivalentTo(new SheeterContext(sheeter));
                sheet.Should().NotBeNull("sheet should be created.");
                sheet!.Name.Should().Be(sheeter.SheetName, "sheet name should match.");
                sheet.Column(1).Width.Should().Be((short)sheeter.ColumnWidths[0], "column width should match.");
                sheet.GetValue(1, 1).Should().Be(cellValue, "cell value should match.");
            };

            exporter.Export();
        }
    }
}
