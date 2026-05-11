namespace CloudyWing.SpreadsheetExporter;

/// <summary>
/// Represents optional capabilities supported by a spreadsheet renderer.
/// </summary>
public sealed record RendererCapabilities {
    /// <summary>
    /// Gets a value indicating whether the renderer supports cell styles.
    /// </summary>
    public bool SupportsStyles { get; init; }

    /// <summary>
    /// Gets a value indicating whether the renderer supports merged cells.
    /// </summary>
    public bool SupportsMergedCells { get; init; }

    /// <summary>
    /// Gets a value indicating whether the renderer supports formulas.
    /// </summary>
    public bool SupportsFormulas { get; init; }

    /// <summary>
    /// Gets a value indicating whether the renderer supports data validation.
    /// </summary>
    public bool SupportsDataValidation { get; init; }

    /// <summary>
    /// Gets a value indicating whether the renderer supports freeze panes.
    /// </summary>
    public bool SupportsFreezePanes { get; init; }

    /// <summary>
    /// Gets a value indicating whether the renderer supports auto filter.
    /// </summary>
    public bool SupportsAutoFilter { get; init; }

    /// <summary>
    /// Gets a value indicating whether the renderer supports multiple sheets.
    /// </summary>
    public bool SupportsMultipleSheets { get; init; }

    /// <summary>
    /// Gets a value indicating whether the renderer supports worksheet protection.
    /// </summary>
    public bool SupportsWorksheetProtection { get; init; }

    /// <summary>
    /// Gets a value indicating whether the renderer supports page settings.
    /// </summary>
    public bool SupportsPageSettings { get; init; }
}
