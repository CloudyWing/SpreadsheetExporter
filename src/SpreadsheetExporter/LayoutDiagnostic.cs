namespace CloudyWing.SpreadsheetExporter;

/// <summary>
/// Represents a diagnostic found while validating spreadsheet layout.
/// </summary>
/// <param name="Code">The diagnostic code.</param>
/// <param name="Message">The diagnostic message.</param>
/// <param name="Range">The related cell range, or <see langword="null" /> when not applicable.</param>
/// <param name="Source">The related cell source, or <see langword="null" /> when not applicable.</param>
public sealed record LayoutDiagnostic(
    string Code,
    string Message,
    CellRange? Range,
    CellSource? Source
);
