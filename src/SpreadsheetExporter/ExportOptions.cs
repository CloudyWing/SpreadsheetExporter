namespace CloudyWing.SpreadsheetExporter;

/// <summary>
/// Represents options that control validation performed by <see cref="SpreadsheetDocument.Export(ExportOptions)"/>.
/// </summary>
public sealed record ExportOptions {
    /// <summary>
    /// Gets the default options, which export without validation, matching <see cref="SpreadsheetDocument.Export()"/>.
    /// </summary>
    public static ExportOptions Default { get; } = new();

    /// <summary>
    /// Gets a value indicating whether layout diagnostics are checked before rendering.
    /// </summary>
    /// <value>
    /// <see langword="true"/> to throw when <see cref="SpreadsheetDocument.GetLayoutDiagnostics"/> would report diagnostics;
    /// otherwise, <see langword="false"/>. Default is <see langword="false"/>.
    /// </value>
    public bool ValidateLayout { get; init; }

    /// <summary>
    /// Gets a value indicating whether renderer capability diagnostics are checked before rendering.
    /// </summary>
    /// <value>
    /// <see langword="true"/> to throw when <see cref="SpreadsheetDocument.GetRendererCapabilityDiagnostics"/> would report
    /// diagnostics; otherwise, <see langword="false"/>. Default is <see langword="false"/>.
    /// </value>
    public bool ValidateCapabilities { get; init; }
}
