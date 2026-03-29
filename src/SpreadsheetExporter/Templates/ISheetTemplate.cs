namespace CloudyWing.SpreadsheetExporter.Templates;

/// <summary>
/// Defines the contract for a sheet template that produces cell layout information.
/// </summary>
public interface ISheetTemplate {
    /// <summary>
    /// Gets the column span occupied by this template.
    /// </summary>
    /// <value>The number of columns.</value>
    int ColumnSpan { get; }

    /// <summary>
    /// Gets the row span occupied by this template.
    /// </summary>
    /// <value>The number of rows.</value>
    int RowSpan { get; }

    /// <summary>
    /// Gets the computed cell layout for this template.
    /// </summary>
    /// <returns>The <see cref="TemplateLayout"/> for this template.</returns>
    TemplateLayout GetLayout();
}
