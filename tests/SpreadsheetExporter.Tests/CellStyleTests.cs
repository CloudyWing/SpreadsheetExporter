using System.Drawing;

namespace CloudyWing.SpreadsheetExporter.Tests {
    [TestFixture]
    internal class CellStyleTests {
        [Test]
        public void DefaultConstructor_ShouldInitializeToEmpty() {
            CellStyle defaultStyle = new();

            defaultStyle.Should().Be(CellStyle.Empty);
        }

        [Test]
        public void Constructor_WithParameters_ShouldInitializeProperties() {
            CellFont font = new("Arial", 12);
            CellStyle cellStyle = new(
                HorizontalAlignment.Center,
                VerticalAlignment.Middle,
                true,
                true,
                Color.Blue,
                font,
                "dd/MM/yyyy",
                true
            );

            cellStyle.HorizontalAlignment.Should().Be(HorizontalAlignment.Center);
            cellStyle.VerticalAlignment.Should().Be(VerticalAlignment.Middle);
            cellStyle.HasBorder.Should().Be(true);
            cellStyle.WrapText.Should().Be(true);
            cellStyle.BackgroundColor.Should().Be(Color.Blue);
            cellStyle.Font.Should().Be(font);
            cellStyle.DataFormat.Should().Be("dd/MM/yyyy");
            cellStyle.IsLocked.Should().Be(true);
        }

        [Test]
        public void Empty_ShouldReturnDefaultStyle() {
            CellStyle emptyStyle = CellStyle.Empty;

            emptyStyle.HorizontalAlignment.Should().Be(HorizontalAlignment.General);
            emptyStyle.VerticalAlignment.Should().Be(VerticalAlignment.Top);
            emptyStyle.HasBorder.Should().Be(false);
            emptyStyle.WrapText.Should().Be(false);
            emptyStyle.BackgroundColor.Should().Be(default(Color));
            emptyStyle.Font.Should().Be(default(CellFont));
            emptyStyle.DataFormat.Should().BeNull();
            emptyStyle.IsLocked.Should().Be(false);
        }
    }
}
