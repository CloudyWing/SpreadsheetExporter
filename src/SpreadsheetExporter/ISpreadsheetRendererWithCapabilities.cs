namespace CloudyWing.SpreadsheetExporter;

/// <summary>
/// Defines optional capability metadata for spreadsheet renderers.
/// </summary>
public interface ISpreadsheetRendererWithCapabilities : ISpreadsheetRenderer {
    /// <summary>
    /// Gets the renderer capabilities.
    /// </summary>
    RendererCapabilities Capabilities { get; }
}
