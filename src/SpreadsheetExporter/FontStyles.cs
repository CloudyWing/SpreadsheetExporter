using System;

namespace CloudyWing.SpreadsheetExporter {
    /// <summary>The font styles.</summary>
    [Flags]
    public enum FontStyles {
        /// <summary>The none.</summary>
        None = 0,

        /// <summary>The is bold.</summary>
        IsBold = 1,

        /// <summary>The is italic.</summary>
        IsItalic = 2,

        /// <summary>The has underline.</summary>
        HasUnderline = 4,

        /// <summary>The is strikeout.</summary>
        IsStrikeout = 8,
    }
}
