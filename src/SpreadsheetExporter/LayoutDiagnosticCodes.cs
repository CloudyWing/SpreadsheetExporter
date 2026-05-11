namespace CloudyWing.SpreadsheetExporter;

/// <summary>
/// Defines layout diagnostic codes produced by spreadsheet validation.
/// </summary>
public static class LayoutDiagnosticCodes {
    /// <summary>
    /// Indicates that a cell range has invalid coordinates or size.
    /// </summary>
    public const string InvalidCellRange = "SE-LAYOUT-001";

    /// <summary>
    /// Indicates that two cell ranges overlap.
    /// </summary>
    public const string OverlappingCellRange = "SE-LAYOUT-002";

    /// <summary>
    /// Indicates that a cell defines both a value and a formula.
    /// </summary>
    public const string ValueAndFormulaConflict = "SE-LAYOUT-003";

    /// <summary>
    /// Indicates that a row height or column width value is invalid.
    /// </summary>
    public const string InvalidDimension = "SE-LAYOUT-004";
}
