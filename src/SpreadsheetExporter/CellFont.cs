using System.Drawing;

namespace CloudyWing.SpreadsheetExporter {
    /// <summary>
    /// The cell font.
    /// </summary>
    /// <param name="Name">The font name.</param>
    /// <param name="Size">The size.</param>
    /// <param name="Color">The color.</param>
    /// <param name="Style">The style.</param>
    public record struct CellFont(
        string Name = null,
        short Size = 0,
        Color Color = default,
        FontStyles Style = FontStyles.None
    ) {
        /// <summary>
        /// The cell font equals to <c>new CellFont()</c>.
        /// </summary>
        public static CellFont Empty { get; } = new();
    }
}
