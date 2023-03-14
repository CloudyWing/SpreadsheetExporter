using System.Drawing;

namespace CloudyWing.SpreadsheetExporter.Excel.NPOI.Tests {
    internal class ExcelExporterTests {
        [Test]
        [TestCase(ExcelFormat.ExcelBinaryFileFormat, "application/ms-excel")]
        [TestCase(ExcelFormat.OfficeOpenXmlDocument, "application/ms-excel")]
        public void ContentType_ShouldReturnExpectedValue(ExcelFormat format, string contentType) {
            var exporter = new ExcelExporter(format);

            exporter.ContentType.Should().Be(contentType);
        }

        [Test]
        [TestCase(ExcelFormat.ExcelBinaryFileFormat, ".xls")]
        [TestCase(ExcelFormat.OfficeOpenXmlDocument, ".xlsx")]
        public void FileNameExtension_ShouldReturnExpectedValue(ExcelFormat format, string fileNameExtension) {
            ExcelExporter exporter = new ExcelExporter(format);

            exporter.FileNameExtension.Should().Be(fileNameExtension);
        }

        [Test]
        [TestCase(ExcelFormat.ExcelBinaryFileFormat, false)]
        [TestCase(ExcelFormat.OfficeOpenXmlDocument, true)]
        public void IsOfficeOpenXmlDocument_ShouldReturnExpectedValue(ExcelFormat format, bool isXlsx) {
            ExcelExporter exporter = new ExcelExporter(format);

            exporter.IsOfficeOpenXmlDocument.Should().Be(isXlsx);
        }

        [Test]
        public void Export_ShouldReturnExpectedByteArray() {
            ExcelExporter exporter = new ExcelExporter();
            exporter.CreateSheeter();

            exporter.Export().Should().NotBeEmpty();
        }

        [Test]
        [TestCase(ExcelFormat.ExcelBinaryFileFormat, true)]
        [TestCase(ExcelFormat.OfficeOpenXmlDocument, true)]
        [TestCase(ExcelFormat.ExcelBinaryFileFormat, false)]
        [TestCase(ExcelFormat.OfficeOpenXmlDocument, false)]
        public void Export_IsClosedNotImplementedException_WithWatermark(ExcelFormat format, bool isClosedNotImplementedException) {
            ExcelExporter exporter = new ExcelExporter(format) {
                IsClosedNotImplementedException = isClosedNotImplementedException
            };

            Sheeter sheeter = exporter.CreateSheeter();
            sheeter.Watermark = new Bitmap(100, 100);

            Action act = () => exporter.Export();

            if (format == ExcelFormat.OfficeOpenXmlDocument || isClosedNotImplementedException) {
                act.Should().NotThrow<NotImplementedException>();
            } else {
                act.Should().Throw<NotImplementedException>()
                    .WithMessage("NPOI currently does not support the output of xls file with watermarks.");
            }
        }

        [Test]
        [TestCase(ExcelFormat.ExcelBinaryFileFormat, true)]
        [TestCase(ExcelFormat.OfficeOpenXmlDocument, true)]
        [TestCase(ExcelFormat.ExcelBinaryFileFormat, false)]
        [TestCase(ExcelFormat.OfficeOpenXmlDocument, false)]
        public void Export_IsClosedNotImplementedException_WithPassword(ExcelFormat format, bool isClosedNotImplementedException) {
            ExcelExporter exporter = new ExcelExporter(format) {
                IsClosedNotImplementedException = isClosedNotImplementedException,
                Password = "password"
            };

            Sheeter sheeter = exporter.CreateSheeter();

            Action act = () => exporter.Export();

            if (format == ExcelFormat.ExcelBinaryFileFormat || isClosedNotImplementedException) {
                act.Should().NotThrow<NotImplementedException>();
            } else {
                act.Should().Throw<NotImplementedException>()
                    .WithMessage("If no other packages are installed, NPOI currently does not support the output of xlsx file with passwords.");
            }
        }
    }
}
