using System.Drawing;

namespace CloudyWing.SpreadsheetExporter.Tests {
    [TestFixture]
    internal class CellFontTests {
        [Test]
        public void CloneAndSetName_SetName_ShouldCloneAndSetName() {
            CellFont cellFont = new CellFont("Arial");

            CellFont newCellFont = cellFont.CloneAndSetName("Times New Roman");

            newCellFont.Name.Should().Be("Times New Roman");
            newCellFont.Should().NotBeSameAs(cellFont);
            cellFont.Name.Should().Be("Arial");
        }

        [Test]
        public void CloneAndSetSize_SetSize_ShouldCloneAndSetSize() {
            CellFont cellFont = new CellFont("Arial", 12);

            CellFont newCellFont = cellFont.CloneAndSetSize(14);

            newCellFont.Size.Should().Be(14);
            newCellFont.Should().NotBeSameAs(cellFont);
            cellFont.Size.Should().Be(12);
        }

        [Test]
        public void CloneAndSetColor_SetColor_ShouldCloneAndSetColor() {
            CellFont font = new CellFont("Arial", 12, Color.Red, FontStyles.IsBold);
            Color expectedColor = Color.Green;

            CellFont clonedFont = font.CloneAndSetColor(expectedColor);

            clonedFont.Color.Should().Be(expectedColor);
            clonedFont.Name.Should().Be(font.Name);
            clonedFont.Size.Should().Be(font.Size);
            clonedFont.Style.Should().Be(font.Style);
        }

        [Test]
        public void CloneAndSetStyle_SetStyle_ShouldCloneAndSetStyle() {
            CellFont font = new CellFont("Arial", 12, Color.Red, FontStyles.IsBold);
            FontStyles expectedStyle = FontStyles.IsItalic;

            CellFont clonedFont = font.CloneAndSetStyle(expectedStyle);

            clonedFont.Color.Should().Be(font.Color);
            clonedFont.Name.Should().Be(font.Name);
            clonedFont.Size.Should().Be(font.Size);
            clonedFont.Style.Should().Be(expectedStyle);
        }

        [Test]
        public void Equals_SameProperties_ShouldBeTrue() {
            CellFont font1 = new CellFont("Arial", 12, Color.Red, FontStyles.IsBold);
            CellFont font2 = new CellFont("Arial", 12, Color.Red, FontStyles.IsBold);

            font1.Equals(font2).Should().BeTrue();
        }

        [Test]
        public void Equals_DifferentName_ShouldBeFalse() {
            CellFont font1 = new CellFont("Arial", 12, Color.Red, FontStyles.IsBold);
            CellFont font2 = new CellFont("Calibri", 12, Color.Red, FontStyles.IsBold);

            font1.Equals(font2).Should().BeFalse();
        }

        [Test]
        public void Equals_DifferentSize_ShouldBeFalse() {
            CellFont font1 = new CellFont("Arial", 12, Color.Red, FontStyles.IsBold);
            CellFont font2 = new CellFont("Arial", 14, Color.Red, FontStyles.IsBold);

            font1.Equals(font2).Should().BeFalse();
        }

        [Test]
        public void Equals_DifferentColor_ShouldBeFalse() {
            CellFont font1 = new CellFont("Arial", 12, Color.Red, FontStyles.IsBold);
            CellFont font2 = new CellFont("Arial", 12, Color.Green, FontStyles.IsBold);

            font1.Equals(font2).Should().BeFalse();
        }

        [Test]
        public void Equals_DifferentStyle_ShouldBeFalse() {
            CellFont font1 = new CellFont("Arial", 12, Color.Red, FontStyles.IsBold);
            CellFont font2 = new CellFont("Arial", 12, Color.Red, FontStyles.IsItalic);

            font1.Equals(font2).Should().BeFalse();
        }

        [Test]
        public void GetHashCode_SameProperties_ShouldBeEqual() {
            CellFont font1 = new CellFont("Arial", 12, Color.Red, FontStyles.IsBold);
            CellFont font2 = new CellFont("Arial", 12, Color.Red, FontStyles.IsBold);

            int hashCode1 = font1.GetHashCode();
            int hashCode2 = font2.GetHashCode();

            hashCode1.Should().Be(hashCode2);
        }

        [Test]
        public void GetHashCode_DifferentName_ShouldNotBeEqual() {
            CellFont font1 = new CellFont("Arial", 12, Color.Red, FontStyles.IsBold);
            CellFont font2 = new CellFont("Calibri", 12, Color.Red, FontStyles.IsBold);

            int hashCode1 = font1.GetHashCode();
            int hashCode2 = font2.GetHashCode();

            hashCode1.Should().NotBe(hashCode2);
        }

        [Test]
        public void GetHashCode_DifferentSize_ShouldNotBeEqual() {
            CellFont font1 = new CellFont("Arial", 12, Color.Red, FontStyles.IsBold);
            CellFont font2 = new CellFont("Arial", 14, Color.Red, FontStyles.IsBold);


            int hashCode1 = font1.GetHashCode();
            int hashCode2 = font2.GetHashCode();

            hashCode1.Should().NotBe(hashCode2);
        }

        [Test]
        public void GetHashCode_DifferentColor_ShouldNotBeEqual() {
            CellFont font1 = new CellFont("Arial", 12, Color.Red, FontStyles.IsBold);
            CellFont font2 = new CellFont("Arial", 12, Color.Green, FontStyles.IsBold);


            int hashCode1 = font1.GetHashCode();
            int hashCode2 = font2.GetHashCode();

            hashCode1.Should().NotBe(hashCode2);
        }

        [Test]
        public void GetHashCode_DifferentStyle_ShouldNotBeEqual() {
            CellFont font1 = new CellFont("Arial", 12, Color.Red, FontStyles.IsBold);
            CellFont font2 = new CellFont("Arial", 12, Color.Red, FontStyles.IsItalic);


            int hashCode1 = font1.GetHashCode();
            int hashCode2 = font2.GetHashCode();

            hashCode1.Should().NotBe(hashCode2);
        }
    }
}
