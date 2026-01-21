namespace CloudyWing.SpreadsheetExporter.Templates;

/// <summary>
/// The template interface.
/// </summary>
public interface ITemplate {
    /// <summary>
    /// Gets the context.
    /// </summary>
    /// <returns>The template context.</returns>
    TemplateContext GetContext();
}
