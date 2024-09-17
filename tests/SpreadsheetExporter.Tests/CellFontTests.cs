using System.Drawing;

namespace CloudyWing.SpreadsheetExporter.Tests {
    [TestFixture]
    internal class CellFontTests {
        [Test]
        public void DefaultConstructor_ShouldInitializeToEmpty() {
            CellFont defaultFont = new();
            CellFont emptyFont = CellFont.Empty;

            defaultFont.Should().Be(emptyFont);
        }

        [Test]
        public void Constructor_WithParameters_ShouldInitializeProperties() {
            string name = "Arial";
            short size = 12;
            Color color = Color.Red;
            FontStyles style = FontStyles.IsBold;

            CellFont font = new(name, size, color, style);

            font.Name.Should().Be(name);
            font.Size.Should().Be(size);
            font.Color.Should().Be(color);
            font.Style.Should().Be(style);
        }

        [Test]
        public void Empty_ShouldReturnDefaultFont() {
            CellFont emptyFont = CellFont.Empty;

            emptyFont.Name.Should().BeNull();
            emptyFont.Size.Should().Be(0);
            emptyFont.Color.Should().Be(default(Color));
            emptyFont.Style.Should().Be(FontStyles.None);
        }
    }
}
