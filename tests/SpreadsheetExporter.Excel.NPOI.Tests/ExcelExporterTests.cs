using CloudyWing.SpreadsheetExporter.Templates.Grid;
using NPOI.SS.UserModel;

namespace CloudyWing.SpreadsheetExporter.Excel.NPOI.Tests {
    internal class ExcelExporterTests {
        [Test]
        [TestCase(ExcelFormat.ExcelBinaryFileFormat, "application/vnd.ms-excel")]
        [TestCase(ExcelFormat.OfficeOpenXmlDocument, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")]
        public void ContentType_ShouldReturnExpectedValue(ExcelFormat format, string contentType) {
            ExcelExporter exporter = new(format);

            Assert.That(exporter.ContentType, Is.EqualTo(contentType));
        }

        [Test]
        [TestCase(ExcelFormat.ExcelBinaryFileFormat, ".xls")]
        [TestCase(ExcelFormat.OfficeOpenXmlDocument, ".xlsx")]
        public void FileNameExtension_ShouldReturnExpectedValue(ExcelFormat format, string fileNameExtension) {
            ExcelExporter exporter = new(format);

            Assert.That(exporter.FileNameExtension, Is.EqualTo(fileNameExtension));
        }

        [Test]
        [TestCase(ExcelFormat.ExcelBinaryFileFormat, false)]
        [TestCase(ExcelFormat.OfficeOpenXmlDocument, true)]
        public void IsOfficeOpenXmlDocument_ShouldReturnExpectedValue(ExcelFormat format, bool isXlsx) {
            ExcelExporter exporter = new(format);

            Assert.That(exporter.IsOfficeOpenXmlDocument, Is.EqualTo(isXlsx));
        }

        [Test]
        public void Export_ShouldReturnExpectedByteArray() {
            ExcelExporter exporter = new();
            exporter.CreateSheeter();

            Assert.That(exporter.Export(), Is.Not.Empty);
        }

        [Test]
        [TestCase(ExcelFormat.ExcelBinaryFileFormat, true)]
        [TestCase(ExcelFormat.OfficeOpenXmlDocument, true)]
        [TestCase(ExcelFormat.ExcelBinaryFileFormat, false)]
        [TestCase(ExcelFormat.OfficeOpenXmlDocument, false)]
        public void Export_IsClosedNotImplementedException_WithPassword(ExcelFormat format, bool isClosedNotImplementedException) {
            ExcelExporter exporter = new(format) {
                IsClosedNotImplementedException = isClosedNotImplementedException,
                Password = "password"
            };

            Sheeter sheeter = exporter.CreateSheeter();

            Action act = () => exporter.Export();

            if (format == ExcelFormat.ExcelBinaryFileFormat || isClosedNotImplementedException) {
                Assert.DoesNotThrow(() => act());
            } else {
                var ex = Assert.Throws<NotImplementedException>(() => act());
                Assert.That(ex.Message, Is.EqualTo("If no other packages are installed, NPOI currently does not support the output of xlsx file with passwords."));
            }
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
                ISheet? sheet = args.SheetObject as ISheet;
                SheeterContext sheeterContext = args.SheeterContext;

                Assert.That(sheeterContext.SheetName, Is.EqualTo(sheeter.SheetName), "SheeterContext sheet name should match.");
                Assert.That(sheeterContext.DefaultRowHeight, Is.EqualTo(sheeter.DefaultRowHeight), "SheeterContext default row height should match.");
                Assert.That(sheet, Is.Not.Null, "Sheet should be created.");
                Assert.That(sheet!.SheetName, Is.EqualTo(sheeter.SheetName), "Sheet name should match.");
                Assert.That(sheet.GetColumnWidth(0), Is.EqualTo((short)(sheeter.ColumnWidths[0] * 256)), "Column width should match.");
                Assert.That(sheet.GetRow(0).GetCell(0).StringCellValue, Is.EqualTo(cellValue), "Cell value should match.");
            };

            exporter.Export();
        }
    }
}
