using System.Collections.Generic;

namespace CloudyWing.SpreadsheetExporter;

/// <summary>
/// Defines the contract for rendering sheet contexts into a binary spreadsheet format.
/// </summary>
public interface ISpreadsheetRenderer {
    /// <summary>
    /// Gets the MIME content-type of the output format.
    /// </summary>
    /// <value>The MIME content-type string.</value>
    string ContentType { get; }

    /// <summary>
    /// Gets the file name extension of the output format, including the leading dot.
    /// </summary>
    /// <value>The file extension, e.g. ".xlsx".</value>
    string FileNameExtension { get; }

    /// <summary>
    /// Renders the specified sheet layouts into a byte array.
    /// </summary>
    /// <param name="layouts">The collection of sheet layouts to render.</param>
    /// <returns>The rendered spreadsheet as a byte array.</returns>
    byte[] Render(IEnumerable<SheetLayout> layouts);
}
