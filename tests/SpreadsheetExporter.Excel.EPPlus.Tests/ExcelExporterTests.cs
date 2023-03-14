namespace CloudyWing.SpreadsheetExporter.Excel.EPPlus.Tests {
    public class Tests {
        [Test]
        public void ContentType_ShouldReturnExpectedContentType() {
            var exporter = new ExcelExporter();

            exporter.ContentType.Should().Be("application/ms-excel");
        }

        [Test]
        public void FileNameExtension_ShouldReturnExpectedFileNameExtension() {
            ExcelExporter exporter = new ExcelExporter();

            exporter.FileNameExtension.Should().Be(".xlsx");
        }

        [Test]
        public void Export_ShouldReturnExpectedByteArray() {
            ExcelExporter exporter = new ExcelExporter();
            exporter.CreateSheeter();

            exporter.Export().Should().NotBeEmpty();
        }
    }
}
