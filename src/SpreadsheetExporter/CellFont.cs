using System;
using System.Collections.Generic;
using System.Drawing;

namespace CloudyWing.SpreadsheetExporter {
    /// <summary>
    /// The cell font.
    /// </summary>
    public struct CellFont : IEquatable<CellFont> {
        /// <summary>
        /// The cell font equals to <c>new CellFont()</c>.
        /// </summary>
        public static CellFont Empty = new();

        /// <summary>
        /// Initializes a new instance of the <see cref="CellFont" /> struct.
        /// </summary>
        /// <param name="name">The font name.</param>
        /// <param name="size">The size.</param>
        /// <param name="color">The color.</param>
        /// <param name="style">The style.</param>
        public CellFont(
            string name, short size = 0, Color? color = null, FontStyles style = FontStyles.None
        ) {
            Name = name;
            Size = size;
            Color = color ?? Color.Empty;
            Style = style;
        }

        /// <summary>
        /// Gets the font name.
        /// </summary>
        /// <value>
        /// The font name.
        /// </value>
        public string Name { get; private set; }

        /// <summary>
        /// Gets the size.
        /// </summary>
        /// <value>
        /// The size.
        /// </value>
        public short Size { get; private set; }

        /// <summary>
        /// Gets the color. The default is <c>Color.Empty</c>.
        /// </summary>
        /// <value>
        /// The color.
        /// </value>
        public Color Color { get; private set; } = Color.Empty;

        /// <summary>
        /// Gets the style.
        /// </summary>
        /// <value>
        /// The style.
        /// </value>
        public FontStyles Style { get; private set; }

        /// <summary>
        /// Clones and set the name of new instance.
        /// </summary>
        /// <param name="name">The name.</param>
        /// <returns>The cloned new cell.</returns>
        public CellFont CloneAndSetName(string name) {
            return this with { Name = name };
        }

        /// <summary>
        /// Clones and set the size of new instance.
        /// </summary>
        /// <param name="size">The size.</param>
        /// <returns>The cloned new cell.</returns>
        public CellFont CloneAndSetSize(short size) {
            return this with { Size = size };
        }

        /// <summary>
        /// Clones and set the color of new instance.
        /// </summary>
        /// <param name="color">The color.</param>
        /// <returns>The cloned new cell.</returns>
        public CellFont CloneAndSetColor(Color color) {
            return this with { Color = color };
        }

        /// <summary>
        /// Clones and set the style of new instance.
        /// </summary>
        /// <param name="style">The style.</param>
        /// <returns>The cloned new cell.</returns>
        public CellFont CloneAndSetStyle(FontStyles style) {
            return this with { Style = style };
        }

        /// <inheritdoc/>
        public override bool Equals(object obj) {
            return obj is CellFont font && Equals(font);
        }

        /// <inheritdoc/>
        public bool Equals(CellFont other) {
            return Name == other.Name
                   && Size == other.Size
                   && EqualityComparer<Color>.Default.Equals(Color, other.Color)
                   && Style == other.Style;
        }

        /// <inheritdoc/>
        public override int GetHashCode() {
            int hashCode = 253911835;
            hashCode = (hashCode * -1521134295) + EqualityComparer<string>.Default.GetHashCode(Name);
            hashCode = (hashCode * -1521134295) + Size.GetHashCode();
            hashCode = (hashCode * -1521134295) + Color.GetHashCode();
            hashCode = (hashCode * -1521134295) + Style.GetHashCode();
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
        public static bool operator ==(CellFont left, CellFont right) {
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
        public static bool operator !=(CellFont left, CellFont right) {
            return !(left == right);
        }
    }
}
