using System;
using System.Collections.Generic;
using System.Drawing;

namespace CloudyWing.SpreadsheetExporter {
    /// <summary>
    /// The cell style.
    /// </summary>
    /// <seealso cref="IEquatable{CellStyle}" />
    public struct CellStyle : IEquatable<CellStyle> {
        /// <summary>
        /// The cell style equals to <c>new CellStyle()</c>.
        /// </summary>
        public static CellStyle Empty = new();

        /// <summary>
        /// Initializes a new instance of the <see cref="CellStyle"/> struct.
        /// </summary>
        /// <param name="halign">The halign.</param>
        /// <param name="valign">The valign.</param>
        /// <param name="hasBorder">if set to <c>true</c> [has border].</param>
        /// <param name="wrapText">if set to <c>true</c> [wrap text].</param>
        /// <param name="backgroundColor">Color of the background. The default is <c>Color.Empty</c>.</param>
        /// <param name="font">The font. The default is <c>CellFont.Empty</c>.</param>
        /// <param name="dataFormat">The data format.</param>
        /// <param name="isLocked">if set to <c>true</c> [is locked].</param>
        public CellStyle(
            HorizontalAlignment halign = HorizontalAlignment.Center, VerticalAlignment valign = VerticalAlignment.Middle,
            bool hasBorder = false, bool wrapText = false, Color? backgroundColor = null,
            CellFont? font = null, string dataFormat = null, bool isLocked = false
        ) : this() {
            backgroundColor ??= Color.Empty;
            font ??= new CellFont();
            HorizontalAlignment = halign;
            VerticalAlignment = valign;
            HasBorder = hasBorder;
            WrapText = wrapText;
            BackgroundColor = (Color)backgroundColor;
            Font = (CellFont)font;
            DataFormat = dataFormat;
            IsLocked = isLocked;
        }

        /// <summary>
        /// Gets the horizontal alignment.
        /// </summary>
        /// <value>
        /// The horizontal alignment.
        /// </value>
        public HorizontalAlignment HorizontalAlignment { get; private set; }

        /// <summary>
        /// Gets the vertical alignment.
        /// </summary>
        /// <value>
        /// The vertical alignment.
        /// </value>
        public VerticalAlignment VerticalAlignment { get; private set; }

        /// <summary>
        /// Gets a value indicating whether this instance has border.
        /// </summary>
        /// <value>
        ///   <c>true</c> if this instance has border; otherwise, <c>false</c>.</value>
        public bool HasBorder { get; private set; }

        /// <summary>
        /// Gets a value indicating whether [wrap text].
        /// </summary>
        /// <value>
        ///   <c>true</c> if [wrap text]; otherwise, <c>false</c>.</value>
        public bool WrapText { get; private set; }

        /// <summary>
        /// Gets the color of the background.
        /// </summary>
        /// <value>
        /// The color of the background.
        /// </value>
        public Color BackgroundColor { get; private set; }

        /// <summary>
        /// Gets the font.
        /// </summary>
        /// <value>
        /// The font.
        /// </value>
        public CellFont Font { get; private set; }

        /// <summary>
        /// Gets the data format.
        /// </summary>
        /// <value>
        /// The data format.
        /// </value>
        public string DataFormat { get; private set; }

        /// <summary>
        /// Gets a value indicating whether this instance is locked.
        /// </summary>
        /// <value>
        ///   <c>true</c> if this instance is locked; otherwise, <c>false</c>.</value>
        public bool IsLocked { get; private set; }

        /// <summary>
        /// Clones and set horizontal alignment of new instance.
        /// </summary>
        /// <param name="align">The align.</param>
        /// <returns>The cloned new cell style.</returns>
        public CellStyle CloneAndSetHorizontalAlignment(HorizontalAlignment align) {
            return this with { HorizontalAlignment = align };
        }

        /// <summary>
        /// Clones and set vertical alignment of new instance.
        /// </summary>
        /// <param name="valign">The valign.</param>
        /// <returns>The cloned new cell style.</returns>
        public CellStyle CloneAndSetVerticalAlignment(VerticalAlignment valign) {
            return this with { VerticalAlignment = valign };
        }

        /// <summary>
        /// Clones and set border of new instance.
        /// </summary>
        /// <param name="hasBolder">if set to <c>true</c> [has bolder].</param>
        /// <returns>The cloned new cell style.</returns>
        public CellStyle CloneAndSetBorder(bool hasBolder) {
            return this with { HasBorder = hasBolder };
        }

        /// <summary>
        /// Clones and set wrap text of new instance.
        /// </summary>
        /// <param name="wrapText">if set to <c>true</c> [wrap text].</param>
        /// <returns>The cloned new cell style.</returns>
        public CellStyle CloneAndSetWrapText(bool wrapText) {
            return this with { WrapText = wrapText };
        }

        /// <summary>
        /// Clones and set background color of new instance.
        /// </summary>
        /// <param name="backgroundColor">Color of the background.</param>
        /// <returns>The cloned new cell style.</returns>
        public CellStyle CloneAndSetBackgroundColor(Color backgroundColor) {
            return this with { BackgroundColor = backgroundColor };
        }

        /// <summary>
        /// Clones and set font of new instance.
        /// </summary>
        /// <param name="font">The font.</param>
        /// <returns>The cloned new cell style.</returns>
        public CellStyle CloneAndSetFont(CellFont font) {
            return this with { Font = font };
        }

        /// <summary>
        /// Clones and set data format of new instance.
        /// </summary>
        /// <param name="dataForamt">The data foramt.</param>
        /// <returns>The cloned new cell style.</returns>
        public CellStyle CloneAndSetDataFormat(string dataForamt) {
            return this with { DataFormat = dataForamt };
        }

        /// <summary>
        /// Clones and set lockedof of new instance.
        /// </summary>
        /// <param name="isLocked">if set to <c>true</c> [is locked].</param>
        /// <returns>The cloned new cell style.</returns>
        public CellStyle CloneAndSetLocked(bool isLocked) {
            return this with { IsLocked = isLocked };
        }

        /// <inheritdoc/>
        public override bool Equals(object obj) {
            return obj is CellStyle style && Equals(style);
        }

        /// <inheritdoc/>
        public bool Equals(CellStyle other) {
            return HorizontalAlignment == other.HorizontalAlignment
                && VerticalAlignment == other.VerticalAlignment
                && HasBorder == other.HasBorder
                && WrapText == other.WrapText
                && EqualityComparer<Color>.Default.Equals(BackgroundColor, other.BackgroundColor)
                && EqualityComparer<CellFont>.Default.Equals(Font, other.Font)
                && DataFormat == other.DataFormat
                && IsLocked == other.IsLocked;
        }

        /// <inheritdoc/>
        public override int GetHashCode() {
            int hashCode = 253911835;
            hashCode = (hashCode * -1521134295) + HorizontalAlignment.GetHashCode();
            hashCode = (hashCode * -1521134295) + VerticalAlignment.GetHashCode();
            hashCode = (hashCode * -1521134295) + HasBorder.GetHashCode();
            hashCode = (hashCode * -1521134295) + WrapText.GetHashCode();
            hashCode = (hashCode * -1521134295) + BackgroundColor.GetHashCode();
            hashCode = (hashCode * -1521134295) + Font.GetHashCode();
            hashCode = (hashCode * -1521134295) + EqualityComparer<string>.Default.GetHashCode(DataFormat);
            hashCode = (hashCode * -1521134295) + IsLocked.GetHashCode();
            return hashCode;
        }

        /// <summary>
        /// Implements the operator op_Equality.
        /// </summary>
        /// <param name="left">The left.</param>
        /// <param name="right">The right.</param>
        /// <returns>
        /// The result of the operator.
        /// </returns>
        public static bool operator ==(CellStyle left, CellStyle right) {
            return left.Equals(right);
        }

        /// <summary>
        /// Implements the operator op_Inequality.
        /// </summary>
        /// <param name="left">The left.</param>
        /// <param name="right">The right.</param>
        /// <returns>
        /// The result of the operator.
        /// </returns>
        public static bool operator !=(CellStyle left, CellStyle right) {
            return !(left == right);
        }
    }
}
