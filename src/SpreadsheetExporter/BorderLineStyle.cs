namespace CloudyWing.SpreadsheetExporter;

/// <summary>
/// Specifies the line style of a cell's outside border.
/// </summary>
public enum BorderLineStyle {
    /// <summary>
    /// No explicit line style. <see cref="CellStyle.HasBorder"/> decides whether a thin border is drawn.
    /// </summary>
    None,

    /// <summary>
    /// A thin solid line.
    /// </summary>
    Thin,

    /// <summary>
    /// A medium solid line.
    /// </summary>
    Medium,

    /// <summary>
    /// A thick solid line.
    /// </summary>
    Thick,

    /// <summary>
    /// A thin dashed line.
    /// </summary>
    Dashed,

    /// <summary>
    /// A thin dotted line.
    /// </summary>
    Dotted,

    /// <summary>
    /// A double line.
    /// </summary>
    Double
}
