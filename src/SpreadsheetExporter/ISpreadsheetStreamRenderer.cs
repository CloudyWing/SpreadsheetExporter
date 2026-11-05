using System.Collections.Generic;
using System.IO;

namespace CloudyWing.SpreadsheetExporter;

/// <summary>
/// Defines an optional renderer contract that writes the rendered spreadsheet directly to a stream.
/// </summary>
/// <remarks>
/// <see cref="SpreadsheetDocument"/> uses this contract when available to avoid creating an intermediate
/// byte array; renderers that implement only <see cref="ISpreadsheetRenderer"/> keep working through
/// <see cref="ISpreadsheetRenderer.Render(IEnumerable{SheetLayout})"/>.
/// </remarks>
public interface ISpreadsheetStreamRenderer : ISpreadsheetRenderer {
    /// <summary>
    /// Renders the specified sheet layouts and writes the result to the specified stream.
    /// </summary>
    /// <param name="layouts">The collection of sheet layouts to render.</param>
    /// <param name="stream">The writable stream that receives the rendered spreadsheet. The stream is not closed.</param>
    void Render(IEnumerable<SheetLayout> layouts, Stream stream);
}
