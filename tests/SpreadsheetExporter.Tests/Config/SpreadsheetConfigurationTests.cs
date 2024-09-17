using System.Drawing;
using CloudyWing.SpreadsheetExporter.Config;

namespace CloudyWing.SpreadsheetExporter.Tests.Config {
    [TestFixture]
    public class CellStyleConfigurationTests {
        private static void CreateTestCellStyle(CellStyle cellStyle, out CellStyle headerStyle, out CellStyle fieldStyle) {
            CellFont headerFont = cellStyle.Font with {
                Style = cellStyle.Font.Style | FontStyles.IsBold
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

            CellStyleConfiguration actual = new(x => {
                x.CellStyle = cellStyle;
                x.HeaderStyle = headerStyle;
                x.FieldStyle = fieldStyle;
            });

            actual.CellStyle.Should().Be(cellStyle);
            actual.HeaderStyle.Should().Be(headerStyle);
            actual.FieldStyle.Should().Be(fieldStyle);
        }
    }
}
