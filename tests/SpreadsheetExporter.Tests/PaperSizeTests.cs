namespace CloudyWing.SpreadsheetExporter.Tests {
    [TestFixture]
    internal class PaperSizeTests {
        [Test]
        public void ImplicitConversion_FromPaperSizeToInt_ShouldReturnValue() {
            PaperSize size = PaperSize.Letter;
            int expected = 1;

            int actual = size;

            Assert.That(actual, Is.EqualTo(expected));
        }

        [Test]
        public void ImplicitConversion_FromIntToPaperSize_ShouldReturnPaperSize() {
            int size = 1;
            PaperSize expected = PaperSize.Letter;

            PaperSize actual = (PaperSize)size;

            Assert.That(actual, Is.EqualTo(expected));
        }

        [Test]
        public void ImplicitConversion_FromInvalidIntToPaperSize_ShouldThrowInvalidCastException() {
            int size = int.MaxValue;

            Action action = () => {
                PaperSize actual = (PaperSize)size;
            };

            Assert.Throws<InvalidCastException>(() => action());
        }

        [Test]
        public void EqualOperator_TwoEqualPaperSizes_ShouldReturnTrue() {
            PaperSize size1 = PaperSize.Letter;
            PaperSize size2 = PaperSize.Letter;

            bool actual = Equals(size1, size2);

            Assert.That(actual, Is.True);
        }

        [Test]
        public void EqualOperator_TwoUnequalPaperSizes_ShouldReturnFalse() {
            PaperSize size1 = PaperSize.Letter;
            PaperSize size2 = PaperSize.LetterSmall;

            bool actual = Equals(size1, size2);

            Assert.That(actual, Is.False);
        }

        [Test]
        public void EqualOperator_PaperSizeAndInt_ShouldReturnTrue() {
            PaperSize size = PaperSize.Letter;
            int value = 1;

            bool actual1 = size.Equals((PaperSize)value);
            bool actual2 = ((PaperSize)value).Equals(size);

            Assert.Multiple(() => {
                Assert.That(actual1, Is.True);
                Assert.That(actual2, Is.True);
            });
        }

        [Test]
        public void EqualOperator_PaperSizeAndInt_ShouldReturnFalse() {
            PaperSize size = PaperSize.Letter;
            int value = 2;

            bool actual1 = size.Equals((PaperSize)value);
            bool actual2 = ((PaperSize)value).Equals(size);

            Assert.Multiple(() => {
                Assert.That(actual1, Is.False);
                Assert.That(actual2, Is.False);
            });
        }

        [Test]
        public void NotEqualOperator_TwoEqualPaperSizes_ShouldReturnFalse() {
            PaperSize size1 = PaperSize.Letter;
            PaperSize size2 = PaperSize.Letter;

            bool actual = !Equals(size1, size2);

            Assert.That(actual, Is.False);
        }

        [Test]
        public void NotEqualOperator_TwoPaperSizesWithDifferentValues_ShouldReturnTrue() {
            PaperSize size1 = PaperSize.Letter;
            PaperSize size2 = PaperSize.Tabloid;

            bool result = !Equals(size1, size2);

            Assert.That(result, Is.True);
        }

        [Test]
        public void NotEqualOperator_TwoDifferentPaperSizes_ShouldReturnTrue() {
            PaperSize paperSize1 = PaperSize.Letter;
            PaperSize paperSize2 = PaperSize.LetterSmall;

            Assert.That(!Equals(paperSize1, paperSize2), Is.True);
        }

        [Test]
        public void NotEqualOperator_NullAndPaperSize_ShouldReturnTrue() {
            PaperSize paperSize = PaperSize.Letter;

            Assert.That(paperSize, Is.Not.Null);
        }

        [Test]
        public void NotEqualOperator_PaperSizeAndNull_ShouldReturnTrue() {
            PaperSize paperSize = PaperSize.Letter;

            Assert.That(paperSize, Is.Not.Null);
        }

        [Test]
        public void Equals_SamePaperSizes_ShouldReturnTrue() {
            PaperSize size1 = PaperSize.Letter;
            PaperSize size2 = PaperSize.Letter;

            bool result = size1.Equals(size2);

            Assert.That(result, Is.True);
        }

        [Test]
        public void Equals_DifferentPaperSizes_ShouldReturnFalse() {
            PaperSize size1 = PaperSize.Letter;
            PaperSize size2 = PaperSize.LetterSmall;

            bool result = size1.Equals(size2);

            Assert.That(result, Is.False);
        }
        [Test]
        public void CompareTo_SamePaperSizes_ShouldReturnZero() {
            PaperSize paperSize1 = PaperSize.Letter;
            PaperSize paperSize2 = PaperSize.Letter;

            int result = paperSize1.CompareTo(paperSize2);

            Assert.That(result, Is.EqualTo(0));
        }

        [Test]
        public void CompareTo_TwoDifferentPaperSizes_ShouldReturnOne() {
            PaperSize paperSize1 = PaperSize.Letter;
            PaperSize paperSize2 = PaperSize.Tabloid;

            int result = paperSize1.CompareTo(paperSize2);

            Assert.That(result, Is.EqualTo(-1));
        }

        [Test]
        public void CompareTo_PaperSizeAndInt_ShouldReturnOne() {
            PaperSize paperSize = PaperSize.Letter;
            int value = 3;

            int result = paperSize.CompareTo(value);

            Assert.That(result, Is.EqualTo(-1));
        }

        [Test]
        public void CompareTo_IntAndPaperSize_ShouldReturnOne() {
            int value = 1;
            PaperSize paperSize = PaperSize.Tabloid;

            int result = value.CompareTo(paperSize);

            Assert.That(result, Is.EqualTo(-1));
        }

        [Test]
        public void Equals_PaperSizeAndNull_ShouldReturnFalse() {
            PaperSize size1 = PaperSize.Letter;

            bool result = size1.Equals(null);

            Assert.That(result, Is.False);
        }

        [Test]
        public void GetHashCode_SamePaperSizes_ShouldBeEqual() {
            PaperSize size1 = PaperSize.Letter;
            PaperSize size2 = PaperSize.Letter;

            int hashCode1 = size1.GetHashCode();
            int hashCode2 = size2.GetHashCode();

            Assert.That(hashCode1, Is.EqualTo(hashCode2));
        }

        [Test]
        public void GetHashCode_DifferentPaperSizes_ShouldNotBeEqual() {
            PaperSize size1 = PaperSize.Letter;
            PaperSize size2 = PaperSize.LetterSmall;

            int hashCode1 = size1.GetHashCode();
            int hashCode2 = size2.GetHashCode();

            Assert.That(hashCode1, Is.Not.EqualTo(hashCode2));
        }
    }
}
