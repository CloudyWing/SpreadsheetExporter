namespace CloudyWing.SpreadsheetExporter;

/// <summary>
/// Provides constant values used throughout the spreadsheet export process.
/// </summary>
public static class Constants {
    /// <summary>
    /// Specifies that the column width should be automatically fitted to the content.
    /// </summary>
    public const int AutoFitColumnWidth = -1;

    /// <summary>
    /// Specifies that the column should be hidden.
    /// </summary>
    public const int HiddenColumn = 0;

    /// <summary>
    /// Specifies that the row height should be automatically fitted to the content.
    /// </summary>
    public const int AutoFitRowHeight = AutoFitColumnWidth;

    /// <summary>
    /// Specifies that the row should be hidden.
    /// </summary>
    public const int HiddenRow = HiddenColumn;
}
