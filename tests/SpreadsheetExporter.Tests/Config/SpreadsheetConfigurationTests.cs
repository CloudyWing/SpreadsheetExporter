using System.Drawing;
using CloudyWing.SpreadsheetExporter.Config;

namespace CloudyWing.SpreadsheetExporter.Tests.Config {
    [TestFixture]
    public class CellStyleConfigurationTests {
        private static void CreateTestCellStyle(CellStyle cellStyle, out CellStyle headerStyle, out CellStyle fieldStyle) {
            CellFont headerFont = cellStyle.Font with {
                Style = cellStyle.Font.Style | FontStyles.Bold
            };

            headerStyle = cellStyle with {
                Font = headerFont,
                HorizontalAlignment = HorizontalAlignment.Center,
                HasBorder = true
            };

            fieldStyle = cellStyle with {
                HasBorder = true,
                HorizontalAlignment = HorizontalAlignment.Left
            };
        }

        [Test]
        public void Constructor_UseAction_ReturnTrue() {
            CellFont cellFont = new("新細明體", 10, Color.Black, FontStyles.None);
            CellStyle cellStyle = new(HorizontalAlignment.General, VerticalAlignment.Middle, false, false, Color.Empty, cellFont);

            CreateTestCellStyle(cellStyle, out CellStyle headerStyle, out CellStyle fieldStyle);

            CellStyleConfiguration actual = new() {
                CellStyle = cellStyle,
                HeaderStyle = headerStyle,
                FieldStyle = fieldStyle,
                GridCellStyle = cellStyle
            };

            Assert.Multiple(() => {
                Assert.That(actual.CellStyle, Is.EqualTo(cellStyle));
                Assert.That(actual.HeaderStyle, Is.EqualTo(headerStyle));
                Assert.That(actual.FieldStyle, Is.EqualTo(fieldStyle));
            });
        }
    }
}
