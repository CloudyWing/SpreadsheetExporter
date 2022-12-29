using System.Drawing;

namespace CloudyWing.SpreadsheetExporter.Tests {
    [TestFixture]
    internal class CellStyleTests {
        [Test]
        public void CloneAndSetHorizontalAlignment_SetAlignment_ShouldCloneAndSetAlignment() {
            CellStyle style = new CellStyle();
            HorizontalAlignment align = HorizontalAlignment.Right;

            CellStyle result = style.CloneAndSetHorizontalAlignment(align);

            result.HorizontalAlignment.Should().Be(align);
            result.VerticalAlignment.Should().Be(style.VerticalAlignment);
            result.HasBorder.Should().Be(style.HasBorder);
            result.WrapText.Should().Be(style.WrapText);
            result.BackgroundColor.Should().Be(style.BackgroundColor);
            result.Font.Should().Be(style.Font);
            result.DataFormat.Should().Be(style.DataFormat);
            result.IsLocked.Should().Be(style.IsLocked);
        }

        [Test]
        public void CloneAndSetVerticalAlignment_SetAlignment_ShouldCloneAndSetAlignment() {
            CellStyle style = new CellStyle();
            VerticalAlignment align = VerticalAlignment.Top;

            CellStyle result = style.CloneAndSetVerticalAlignment(align);

            result.HorizontalAlignment.Should().Be(style.HorizontalAlignment);
            result.VerticalAlignment.Should().Be(align);
            result.HasBorder.Should().Be(style.HasBorder);
            result.WrapText.Should().Be(style.WrapText);
            result.BackgroundColor.Should().Be(style.BackgroundColor);
            result.Font.Should().Be(style.Font);
            result.DataFormat.Should().Be(style.DataFormat);
            result.IsLocked.Should().Be(style.IsLocked);
        }

        [Test]
        public void CloneAndSetHasBorder_SetHasBorder_ShouldCloneAndSetHasBorder() {
            CellStyle style = new CellStyle();
            bool hasBorder = true;

            CellStyle result = style.CloneAndSetBorder(hasBorder);

            result.HorizontalAlignment.Should().Be(style.HorizontalAlignment);
            result.VerticalAlignment.Should().Be(style.VerticalAlignment);
            result.HasBorder.Should().Be(hasBorder);
            result.WrapText.Should().Be(style.WrapText);
            result.BackgroundColor.Should().Be(style.BackgroundColor);
            result.Font.Should().Be(style.Font);
            result.DataFormat.Should().Be(style.DataFormat);
            result.IsLocked.Should().Be(style.IsLocked);
        }

        [Test]
        public void CloneAndSetWrapText_SetWrapText_ShouldCloneAndSetWrapText() {
            CellStyle style = new CellStyle();
            bool wrapText = true;

            CellStyle result = style.CloneAndSetWrapText(wrapText);

            result.HorizontalAlignment.Should().Be(style.HorizontalAlignment);
            result.VerticalAlignment.Should().Be(style.VerticalAlignment);
            result.HasBorder.Should().Be(style.HasBorder);
            result.WrapText.Should().Be(wrapText);
            result.BackgroundColor.Should().Be(style.BackgroundColor);
            result.Font.Should().Be(style.Font);
            result.DataFormat.Should().Be(style.DataFormat);
            result.IsLocked.Should().Be(style.IsLocked);
        }

        [Test]
        public void CloneAndSetBackgroundColor_SetBackgroundColor_ShouldCloneAndSetBackgroundColor() {
            CellStyle style = new CellStyle();
            Color backgroundColor = Color.Red;

            CellStyle result = style.CloneAndSetBackgroundColor(backgroundColor);

            result.HorizontalAlignment.Should().Be(style.HorizontalAlignment);
            result.VerticalAlignment.Should().Be(style.VerticalAlignment);
            result.HasBorder.Should().Be(style.HasBorder);
            result.WrapText.Should().Be(style.WrapText);
            result.BackgroundColor.Should().Be(backgroundColor);
            result.Font.Should().Be(style.Font);
            result.DataFormat.Should().Be(style.DataFormat);
            result.IsLocked.Should().Be(style.IsLocked);
        }

        [Test]
        public void CloneAndSetFont_SetFont_ShouldCloneAndSetFont() {
            CellStyle style = new CellStyle();
            CellFont font = new CellFont("Arial", 12, Color.Red, FontStyles.IsBold);

            CellStyle result = style.CloneAndSetFont(font);

            result.HorizontalAlignment.Should().Be(style.HorizontalAlignment);
            result.VerticalAlignment.Should().Be(style.VerticalAlignment);
            result.HasBorder.Should().Be(style.HasBorder);
            result.WrapText.Should().Be(style.WrapText);
            result.BackgroundColor.Should().Be(style.BackgroundColor);
            result.Font.Should().Be(font);
            result.DataFormat.Should().Be(style.DataFormat);
            result.IsLocked.Should().Be(style.IsLocked);
        }

        [Test]
        public void CloneAndSetDataFormat_SetDataFormat_ShouldCloneAndSetDataFormat() {
            CellStyle style = new CellStyle();
            string dataFormat = "dd/mm/yyyy";

            CellStyle result = style.CloneAndSetDataFormat(dataFormat);

            result.HorizontalAlignment.Should().Be(style.HorizontalAlignment);
            result.VerticalAlignment.Should().Be(style.VerticalAlignment);
            result.HasBorder.Should().Be(style.HasBorder);
            result.WrapText.Should().Be(style.WrapText);
            result.BackgroundColor.Should().Be(style.BackgroundColor);
            result.Font.Should().Be(style.Font);
            result.DataFormat.Should().Be(dataFormat);
            result.IsLocked.Should().Be(style.IsLocked);
        }

        [Test]
        public void CloneAndSetIsLocked_SetIsLocked_ShouldCloneAndSetIsLocked() {
            CellStyle style = new CellStyle();
            bool isLocked = true;

            CellStyle result = style.CloneAndSetLocked(isLocked);

            result.HorizontalAlignment.Should().Be(style.HorizontalAlignment);
            result.VerticalAlignment.Should().Be(style.VerticalAlignment);
            result.HasBorder.Should().Be(style.HasBorder);
            result.WrapText.Should().Be(style.WrapText);
            result.BackgroundColor.Should().Be(style.BackgroundColor);
            result.Font.Should().Be(style.Font);
            result.DataFormat.Should().Be(style.DataFormat);
            result.IsLocked.Should().Be(isLocked);
        }

        [Test]
        public void Equals_SameProperties_ShouldBeTrue() {
            CellStyle style1 = new CellStyle();
            CellStyle style2 = new CellStyle();

            bool result = style1.Equals(style2);

            result.Should().BeTrue();
        }

        [Test]
        public void Equals_DifferentHorizontalAlignment_ShouldBeFalse() {
            CellStyle style1 = new CellStyle(HorizontalAlignment.Left);
            CellStyle style2 = new CellStyle(HorizontalAlignment.Right);

            bool result = style1.Equals(style2);

            result.Should().BeFalse();
        }

        [Test]
        public void Equals_DifferentVerticalAlignment_ShouldBeFalse() {
            CellStyle style1 = new CellStyle(valign: VerticalAlignment.Top);
            CellStyle style2 = new CellStyle(valign: VerticalAlignment.Bottom);

            bool result = style1.Equals(style2);

            result.Should().BeFalse();
        }

        public void Equals_DifferentHasBorder_ShouldBeFalse() {
            CellStyle style1 = new CellStyle(hasBorder: true);
            CellStyle style2 = new CellStyle(hasBorder: false);

            bool result = style1.Equals(style2);

            result.Should().BeFalse();
        }

        [Test]
        public void Equals_DifferentWrapText_ShouldBeFalse() {
            CellStyle style1 = new CellStyle(wrapText: true);
            CellStyle style2 = new CellStyle(wrapText: false);

            bool result = style1.Equals(style2);

            result.Should().BeFalse();
        }

        [Test]
        public void Equals_DifferentBackgroundColor_ShouldBeFalse() {
            CellStyle style1 = new CellStyle(backgroundColor: Color.AliceBlue);
            CellStyle style2 = new CellStyle(backgroundColor: Color.AntiqueWhite);

            bool result = style1.Equals(style2);

            result.Should().BeFalse();
        }

        [Test]
        public void Equals_DifferentFont_ShouldBeFalse() {
            CellStyle style1 = new CellStyle(font: new CellFont("Arial", 12));
            CellStyle style2 = new CellStyle(font: new CellFont("Calibri", 11));

            bool result = style1.Equals(style2);

            result.Should().BeFalse();
        }

        [Test]
        public void Equals_DifferentDataFormat_ShouldBeFalse() {
            CellStyle style1 = new CellStyle(dataFormat: "0");
            CellStyle style2 = new CellStyle(dataFormat: "0.00");

            bool result = style1.Equals(style2);

            result.Should().BeFalse();
        }

        [Test]
        public void Equals_DifferentIsLocked_ShouldBeFalse() {
            CellStyle style1 = new CellStyle(isLocked: true);
            CellStyle style2 = new CellStyle(isLocked: false);

            bool result = style1.Equals(style2);

            result.Should().BeFalse();
        }

        [Test]
        public void GetHashCode_SameProperties_ShouldBeEqual() {
            CellStyle style1 = new CellStyle(
                halign: HorizontalAlignment.Left,
                valign: VerticalAlignment.Top,
                hasBorder: true,
                wrapText: true,
                backgroundColor: Color.Red,
                font: new CellFont("Arial", 12, Color.Black, FontStyles.IsBold),
                dataFormat: "yyyy-MM-dd",
                isLocked: true
            );

            CellStyle style2 = new CellStyle(
                halign: HorizontalAlignment.Left,
                valign: VerticalAlignment.Top,
                hasBorder: true,
                wrapText: true,
                backgroundColor: Color.Red,
                font: new CellFont("Arial", 12, Color.Black, FontStyles.IsBold),
                dataFormat: "yyyy-MM-dd",
                isLocked: true
            );

            style1.GetHashCode().Should().Be(style2.GetHashCode());
        }

        [Test]
        public void GetHashCode_DifferentHorizontalAlignment_ShouldNotBeEqual() {
            CellStyle style1 = new CellStyle(
                halign: HorizontalAlignment.Left,
                valign: VerticalAlignment.Top,
                hasBorder: true,
                wrapText: true,
                backgroundColor: Color.Red,
                font: new CellFont("Arial", 12, Color.Black, FontStyles.IsBold),
                dataFormat: "yyyy-MM-dd",
                isLocked: true
            );

            CellStyle style2 = new CellStyle(
                halign: HorizontalAlignment.Center,
                valign: VerticalAlignment.Top,
                hasBorder: true,
                wrapText: true,
                backgroundColor: Color.Red,
                font: new CellFont("Arial", 12, Color.Black, FontStyles.IsBold),
                dataFormat: "yyyy-MM-dd",
                isLocked: true
            );

            style1.GetHashCode().Should().NotBe(style2.GetHashCode());
        }

        [Test]
        public void GetHashCode_DifferentVerticalAlignment_ShouldNotBeEqual() {
            CellStyle style1 = new CellStyle(
                halign: HorizontalAlignment.Center,
                valign: VerticalAlignment.Top,
                hasBorder: true,
                wrapText: true,
                backgroundColor: Color.Red,
                font: new CellFont("Arial", 12, Color.Black, FontStyles.IsBold),
                dataFormat: "yyyy-MM-dd",
                isLocked: true
            );

            CellStyle style2 = new CellStyle(
                halign: HorizontalAlignment.Center,
                valign: VerticalAlignment.Middle,
                hasBorder: true,
                wrapText: true,
                backgroundColor: Color.Red,
                font: new CellFont("Arial", 12, Color.Black, FontStyles.IsBold),
                dataFormat: "yyyy-MM-dd",
                isLocked: true
            );

            style1.GetHashCode().Should().NotBe(style2.GetHashCode());
        }

        [Test]
        public void GetHashCode_DifferentHasBorder_ShouldNotBeEqual() {
            CellStyle style1 = new CellStyle(
                halign: HorizontalAlignment.Center,
                valign: VerticalAlignment.Middle,
                hasBorder: true,
                wrapText: true,
                backgroundColor: Color.Red,
                font: new CellFont("Arial", 12, Color.Black, FontStyles.IsBold),
                dataFormat: "yyyy-MM-dd",
                isLocked: true
                );

            CellStyle style2 = new CellStyle(
                halign: HorizontalAlignment.Center,
                valign: VerticalAlignment.Middle,
                hasBorder: false,
                wrapText: true,
                backgroundColor: Color.Red,
                font: new CellFont("Arial", 12, Color.Black, FontStyles.IsBold),
                dataFormat: "yyyy-MM-dd",
                isLocked: true
            );

            style1.GetHashCode().Should().NotBe(style2.GetHashCode());
        }

        [Test]
        public void GetHashCode_DifferentWrapText_ShouldNotBeEqual() {
            CellStyle style1 = new CellStyle(
                halign: HorizontalAlignment.Left,
                valign: VerticalAlignment.Top,
                hasBorder: true,
                wrapText: true,
                backgroundColor: Color.Red,
                font: new CellFont("Arial", 12, Color.Black, FontStyles.IsBold),
                dataFormat: "yyyy-MM-dd",
                isLocked: true
            );

            CellStyle style2 = new CellStyle(
                halign: HorizontalAlignment.Left,
                valign: VerticalAlignment.Top,
                hasBorder: true,
                wrapText: false,
                backgroundColor: Color.Red,
                font: new CellFont("Arial", 12, Color.Black, FontStyles.IsBold),
                dataFormat: "yyyy-MM-dd",
                isLocked: true
            );

            style1.GetHashCode().Should().NotBe(style2.GetHashCode());
        }

        [Test]
        public void GetHashCode_DifferentBackgroundColor_ShouldNotBeEqual() {
            CellStyle style1 = new CellStyle(
                halign: HorizontalAlignment.Left,
                valign: VerticalAlignment.Top,
                hasBorder: true,
                wrapText: true,
                backgroundColor: Color.Red,
                font: new CellFont("Arial", 12, Color.Black, FontStyles.IsBold),
                dataFormat: "yyyy-MM-dd",
                isLocked: true
            );

            CellStyle style2 = new CellStyle(
                halign: HorizontalAlignment.Left,
                valign: VerticalAlignment.Top,
                hasBorder: true,
                wrapText: true,
                backgroundColor: Color.Green,
                font: new CellFont("Arial", 12, Color.Black, FontStyles.IsBold),
                dataFormat: "yyyy-MM-dd",
                isLocked: true
            );

            style1.GetHashCode().Should().NotBe(style2.GetHashCode());
        }

        [Test]
        public void GetHashCode_DifferentFont_ShouldNotBeEqual() {
            CellStyle style1 = new CellStyle(
                halign: HorizontalAlignment.Left,
                valign: VerticalAlignment.Top,
                hasBorder: true,
                wrapText: true,
                backgroundColor: Color.Red,
                font: new CellFont("Arial", 12, Color.Black, FontStyles.IsBold),
                dataFormat: "yyyy-MM-dd",
                isLocked: true
            );

            CellStyle style2 = new CellStyle(
                halign: HorizontalAlignment.Left,
                valign: VerticalAlignment.Top,
                hasBorder: true,
                wrapText: true,
                backgroundColor: Color.Red,
                font: new CellFont("Verdana", 12, Color.Black, FontStyles.IsBold),
                dataFormat: "yyyy-MM-dd",
                isLocked: true
            );

            style1.GetHashCode().Should().NotBe(style2.GetHashCode());
        }

        [Test]
        public void GetHashCode_DifferentDataFormat_ShouldNotBeEqual() {
            CellStyle style1 = new CellStyle(
                halign: HorizontalAlignment.Left,
                valign: VerticalAlignment.Top,
                hasBorder: true,
                wrapText: true,
                backgroundColor: Color.Red,
                font: new CellFont("Arial", 12, Color.Black, FontStyles.IsBold),
                dataFormat: "yyyy-MM-dd",
                isLocked: true
            );

            CellStyle style2 = new CellStyle(
                halign: HorizontalAlignment.Left,
                valign: VerticalAlignment.Top,
                hasBorder: true,
                wrapText: true,
                backgroundColor: Color.Red,
                font: new CellFont("Arial", 12, Color.Black, FontStyles.IsBold),
                dataFormat: "yyyy-MM-dd hh:mm:ss",
                isLocked: true
            );

            style1.GetHashCode().Should().NotBe(style2.GetHashCode());
        }

        [Test]
        public void GetHashCode_DifferentIsLocked_ShouldNotBeEqual() {
            CellStyle style1 = new CellStyle(
                halign: HorizontalAlignment.Left,
                valign: VerticalAlignment.Top,
                hasBorder: true,
                wrapText: true,
                backgroundColor: Color.Red,
                font: new CellFont("Arial", 12, Color.Black, FontStyles.IsBold),
                dataFormat: "yyyy-MM-dd",
                isLocked: true
            );

            CellStyle style2 = new CellStyle(
                halign: HorizontalAlignment.Left,
                valign: VerticalAlignment.Top,
                hasBorder: true,
                wrapText: true,
                backgroundColor: Color.Red,
                font: new CellFont("Arial", 12, Color.Black, FontStyles.IsBold),
                dataFormat: "yyyy-MM-dd",
                isLocked: false
            );

            style1.GetHashCode().Should().NotBe(style2.GetHashCode());
        }
    }
}
