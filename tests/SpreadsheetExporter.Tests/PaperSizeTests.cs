using System.Drawing.Printing;

namespace CloudyWing.SpreadsheetExporter.Tests {
    [TestFixture]
    internal class PaperSizeTests {
        [Test]
        public void WidthAndHeight_ValidWithPrinter_ShouldSameSize() {
            foreach (PaperSize item in PaperSize.GetAll()) {
                (int Width, int Height) = GetPaperSize(item.Value);

                item.Width.Should().Be(Width);
                item.Height.Should().Be(Height);
            }
        }

        private (int Width, int Height) GetPaperSize(int value) {
            PrinterSettings settings = new() {
                PrinterName = "Microsoft XPS Document Writer"
            };

            foreach (System.Drawing.Printing.PaperSize printerPaperSize in settings.PaperSizes) {
                if (printerPaperSize.RawKind == value) {
                    return (printerPaperSize.Width, printerPaperSize.Height);
                }
            }

            return (0, 0);
        }

        [Test]
        public void ImplicitConversion_FromPaperSizeToInt_ShouldReturnValue() {
            PaperSize size = PaperSize.Letter;
            int expected = 1;

            int actual = size;

            actual.Should().Be(expected);
        }

        [Test]
        public void ImplicitConversion_FromIntToPaperSize_ShouldReturnPaperSize() {
            int size = 1;
            PaperSize expected = PaperSize.Letter;

            PaperSize actual = (PaperSize)size;

            actual.Should().Be(expected);
        }

        [Test]
        public void ImplicitConversion_FromInvalidIntToPaperSize_ShouldThrowInvalidCastException() {
            int size = int.MaxValue;

            Action action = () => {
                PaperSize actual = (PaperSize)size;
            };

            action.Should().Throw<InvalidCastException>();
        }

        [Test]
        public void EqualOperator_TwoEqualPaperSizes_ShouldReturnTrue() {
            PaperSize size1 = PaperSize.Letter;
            PaperSize size2 = PaperSize.Letter;

            bool actual = size1 == size2;

            actual.Should().BeTrue();
        }

        [Test]
        public void EqualOperator_TwoUnequalPaperSizes_ShouldReturnFalse() {
            PaperSize size1 = PaperSize.Letter;
            PaperSize size2 = PaperSize.LetterSmall;

            bool actual = size1 == size2;

            actual.Should().BeFalse();
        }

        [Test]
        public void EqualOperator_PaperSizeAndInt_ShouldReturnTrue() {
            PaperSize size = PaperSize.Letter;
            int value = 1;

            bool actual1 = size == value;
            bool actual2 = value == size;

            actual1.Should().BeTrue();
            actual2.Should().BeTrue();
        }

        [Test]
        public void EqualOperator_PaperSizeAndInt_ShouldReturnFalse() {
            PaperSize size = PaperSize.Letter;
            int value = 2;

            bool actual1 = size == value;
            bool actual2 = value == size;

            actual1.Should().BeFalse();
            actual2.Should().BeFalse();
        }

        [Test]
        public void NotEqualOperator_TwoEqualPaperSizes_ShouldReturnFalse() {
            PaperSize size1 = PaperSize.Letter;
            PaperSize size2 = PaperSize.Letter;

            bool actual = size1 != size2;

            actual.Should().BeFalse();
        }

        [Test]
        public void NotEqualOperator_TwoPaperSizesWithDifferentValues_ShouldReturnTrue() {
            PaperSize size1 = PaperSize.Letter;
            PaperSize size2 = PaperSize.Tabloid;

            bool result = size1 != size2;

            result.Should().BeTrue();
        }

        [Test]
        public void NotEqualOperator_TwoDifferentPaperSizes_ShouldReturnTrue() {
            PaperSize paperSize1 = PaperSize.Letter;
            PaperSize paperSize2 = PaperSize.LetterSmall;

            (paperSize1 != paperSize2).Should().BeTrue();
        }

        [Test]
        public void NotEqualOperator_NullAndPaperSize_ShouldReturnTrue() {
            PaperSize paperSize = PaperSize.Letter;

            (null != paperSize).Should().BeTrue();
        }

        [Test]
        public void NotEqualOperator_PaperSizeAndNull_ShouldReturnTrue() {
            PaperSize paperSize = PaperSize.Letter;

            (paperSize != null).Should().BeTrue();
        }

        [Test]
        public void Equals_SamePaperSizes_ShouldReturnTrue() {
            PaperSize size1 = PaperSize.Letter;
            PaperSize size2 = PaperSize.Letter;

            bool result = size1.Equals(size2);

            result.Should().BeTrue();
        }

        [Test]
        public void Equals_DifferentPaperSizes_ShouldReturnFalse() {
            PaperSize size1 = PaperSize.Letter;
            PaperSize size2 = PaperSize.LetterSmall;

            bool result = size1.Equals(size2);

            result.Should().BeFalse();
        }
        [Test]
        public void CompareTo_SamePaperSizes_ShouldReturnZero() {
            PaperSize paperSize1 = PaperSize.Letter;
            PaperSize paperSize2 = PaperSize.Letter;

            int result = paperSize1.CompareTo(paperSize2);

            result.Should().Be(0);
        }

        [Test]
        public void CompareTo_TwoDifferentPaperSizes_ShouldReturnOne() {
            PaperSize paperSize1 = PaperSize.Letter;
            PaperSize paperSize2 = PaperSize.Tabloid;

            int result = paperSize1.CompareTo(paperSize2);

            result.Should().Be(-1);
        }

        [Test]
        public void CompareTo_PaperSizeAndInt_ShouldReturnOne() {
            PaperSize paperSize = PaperSize.Letter;
            int value = 3;

            int result = paperSize.CompareTo(value);

            result.Should().Be(-1);
        }

        [Test]
        public void CompareTo_IntAndPaperSize_ShouldReturnOne() {
            int value = 1;
            PaperSize paperSize = PaperSize.Tabloid;

            int result = value.CompareTo(paperSize);

            result.Should().Be(-1);
        }

        [Test]
        public void Equals_PaperSizeAndNull_ShouldReturnFalse() {
            PaperSize size1 = PaperSize.Letter;

            bool result = size1.Equals(null);

            result.Should().BeFalse();
        }

        [Test]
        public void GetHashCode_SamePaperSizes_ShouldBeEqual() {
            PaperSize size1 = PaperSize.Letter;
            PaperSize size2 = PaperSize.Letter;

            int hashCode1 = size1.GetHashCode();
            int hashCode2 = size2.GetHashCode();

            hashCode1.Should().Be(hashCode2);
        }

        [Test]
        public void GetHashCode_DifferentPaperSizes_ShouldNotBeEqual() {
            PaperSize size1 = PaperSize.Letter;
            PaperSize size2 = PaperSize.LetterSmall;

            int hashCode1 = size1.GetHashCode();
            int hashCode2 = size2.GetHashCode();

            hashCode1.Should().NotBe(hashCode2);
        }
    }
}
