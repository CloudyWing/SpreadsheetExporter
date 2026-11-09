using System.Collections.Generic;

namespace CloudyWing.SpreadsheetExporter.Templates;

/// <summary>
/// Provides column widths defined by a template.
/// </summary>
internal interface IColumnWidthProvider {
    /// <summary>
    /// Gets the column widths indexed by zero-based column index.
    /// </summary>
    /// <returns>The column widths defined by the template.</returns>
    IReadOnlyDictionary<int, double> GetColumnWidths();
}
