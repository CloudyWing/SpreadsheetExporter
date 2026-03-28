using System;

namespace CloudyWing.SpreadsheetExporter;

/// <summary>
/// Specifies font style flags that can be combined.
/// </summary>
[Flags]
public enum FontStyles {
    /// <summary>
    /// No font styles applied.
    /// </summary>
    None = 0,

    /// <summary>
    /// Bold text.
    /// </summary>
    Bold = 1,

    /// <summary>
    /// Italic text.
    /// </summary>
    Italic = 2,

    /// <summary>
    /// Underlined text.
    /// </summary>
    Underline = 4,

    /// <summary>
    /// Strikeout text.
    /// </summary>
    Strikeout = 8,
}
