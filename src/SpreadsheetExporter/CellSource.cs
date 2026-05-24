namespace CloudyWing.SpreadsheetExporter;

/// <summary>
/// Represents the source location of a planned cell within a spreadsheet document.
/// </summary>
/// <param name="SheetIndex">The zero-based sheet index.</param>
/// <param name="SheetName">The sheet name.</param>
/// <param name="TemplateIndex">The zero-based template index within the sheet.</param>
/// <param name="TemplateType">The template type name.</param>
/// <param name="Path">The optional source path.</param>
public sealed record CellSource(
    int SheetIndex,
    string SheetName,
    int TemplateIndex,
    string TemplateType,
    string? Path
);
