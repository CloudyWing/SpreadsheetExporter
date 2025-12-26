using System.Drawing;

namespace CloudyWing.SpreadsheetExporter.Tests {
    [TestFixture]
    internal class CellFontTests {
        [Test]
        public void DefaultConstructor_ShouldInitializeToEmpty() {
            CellFont defaultFont = new();
            CellFont emptyFont = CellFont.Empty;

            Assert.That(defaultFont, Is.EqualTo(emptyFont));
        }

        [Test]
        public void Constructor_WithParameters_ShouldInitializeProperties() {
            string name = "Arial";
            short size = 12;
            Color color = Color.Red;
            FontStyles style = FontStyles.IsBold;

            CellFont font = new(name, size, color, style);

            Assert.That(font.Name, Is.EqualTo(name));
            Assert.That(font.Size, Is.EqualTo(size));
            Assert.That(font.Color, Is.EqualTo(color));
            Assert.That(font.Style, Is.EqualTo(style));
        }

        [Test]
        public void Empty_ShouldReturnDefaultFont() {
            CellFont emptyFont = CellFont.Empty;

            Assert.That(emptyFont.Name, Is.Null);
            Assert.That(emptyFont.Size, Is.EqualTo(0));
            Assert.That(emptyFont.Color, Is.EqualTo(default(Color)));
            Assert.That(emptyFont.Style, Is.EqualTo(FontStyles.None));
        }
    }
}
