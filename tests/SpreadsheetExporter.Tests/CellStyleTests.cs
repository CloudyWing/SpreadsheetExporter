using System.Drawing;

namespace CloudyWing.SpreadsheetExporter.Tests {
    [TestFixture]
    internal class CellStyleTests {
        [Test]
        public void DefaultConstructor_ShouldInitializeToEmpty() {
            CellStyle defaultStyle = new();

            Assert.That(defaultStyle, Is.EqualTo(CellStyle.Empty));
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

            Assert.That(cellStyle.HorizontalAlignment, Is.EqualTo(HorizontalAlignment.Center));
            Assert.That(cellStyle.VerticalAlignment, Is.EqualTo(VerticalAlignment.Middle));
            Assert.That(cellStyle.HasBorder, Is.EqualTo(true));
            Assert.That(cellStyle.WrapText, Is.EqualTo(true));
            Assert.That(cellStyle.BackgroundColor, Is.EqualTo(Color.Blue));
            Assert.That(cellStyle.Font, Is.EqualTo(font));
            Assert.That(cellStyle.DataFormat, Is.EqualTo("dd/MM/yyyy"));
            Assert.That(cellStyle.IsLocked, Is.EqualTo(true));
        }

        [Test]
        public void Empty_ShouldReturnDefaultStyle() {
            CellStyle emptyStyle = CellStyle.Empty;

            Assert.That(emptyStyle.HorizontalAlignment, Is.EqualTo(HorizontalAlignment.General));
            Assert.That(emptyStyle.VerticalAlignment, Is.EqualTo(VerticalAlignment.Top));
            Assert.That(emptyStyle.HasBorder, Is.EqualTo(false));
            Assert.That(emptyStyle.WrapText, Is.EqualTo(false));
            Assert.That(emptyStyle.BackgroundColor, Is.EqualTo(default(Color)));
            Assert.That(emptyStyle.Font, Is.EqualTo(default(CellFont)));
            Assert.That(emptyStyle.DataFormat, Is.Null);
            Assert.That(emptyStyle.IsLocked, Is.EqualTo(false));
        }
    }
}
