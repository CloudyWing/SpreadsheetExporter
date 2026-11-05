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

    /// <summary>
    /// Indicates that a sheet name violates the Excel naming rules: it is blank, longer than 31 characters,
    /// contains any of the characters <c>[ ] : * ? / \</c>, or begins or ends with an apostrophe.
    /// </summary>
    public const string InvalidSheetName = "SE-LAYOUT-005";

    /// <summary>
    /// Indicates that a sheet name duplicates another sheet name in the same document.
    /// Sheet names are compared case-insensitively.
    /// </summary>
    public const string DuplicateSheetName = "SE-LAYOUT-006";

    /// <summary>
    /// Indicates that the renderer does not support a feature used by the document.
    /// </summary>
    public const string UnsupportedRendererCapability = "SE-RENDERER-001";
}
