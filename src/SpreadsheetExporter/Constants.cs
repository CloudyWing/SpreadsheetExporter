using System;

namespace CloudyWing.SpreadsheetExporter;

/// <summary>
/// The constants.
/// </summary>
public static class Constants {
    /// <summary>
    /// The automatic fit column width.
    /// </summary>
    public const int AutoFitColumnWidth = -1;

    /// <summary>
    /// The hidden column.
    /// </summary>
    public const int HiddenColumn = 0;

    /// <summary>
    /// The automatic fit row height.
    /// </summary>
    public const int AutoFitRowHeight = AutoFitColumnWidth;

    /// <summary>
    /// The automatic fit row height.
    /// </summary>
    [Obsolete("Use AutoFitRowHeight instead. This property will be removed in a future version.")]
    public const int AutoFiteRowHeight = AutoFitRowHeight;

    /// <summary>
    /// The hidden row.
    /// </summary>
    public const int HiddenRow = HiddenColumn;
}
