using System.Drawing;
using CloudyWing.SpreadsheetExporter.Config;

namespace CloudyWing.SpreadsheetExporter.Tests.Config {
    [TestFixture]
    public class CellStyleConfigurationTests {
        private void CreateTestCellStyle(CellStyle cellStyle, out CellStyle headerStyle, out CellStyle fieldStyle) {
            CellFont headerFont = cellStyle.Font
                    .CloneAndSetStyle(cellStyle.Font.Style | FontStyles.IsBold);

            headerStyle = cellStyle
                    .CloneAndSetFont(headerFont)
                    .CloneAndSetHorizontalAlignment(HorizontalAlignment.Center)
                    .CloneAndSetBorder(true);

            fieldStyle = cellStyle
                .CloneAndSetBorder(true)
                .CloneAndSetHorizontalAlignment(HorizontalAlignment.Left);
        }

        [Test]
        public void Constructor_UseAction_ReturnTrue() {
            CellFont cellFont = new CellFont("新細明體", 10, Color.Black, FontStyles.None);
            CellStyle cellStyle = new CellStyle(HorizontalAlignment.General, VerticalAlignment.Middle, false, false, Color.Empty, cellFont);

            CreateTestCellStyle(cellStyle, out CellStyle headerStyle, out CellStyle fieldStyle);

            CellStyleConfiguration actual = new CellStyleConfiguration(x => {
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
