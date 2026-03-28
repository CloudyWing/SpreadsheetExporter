using System;
using System.Collections.Generic;
using System.Linq;

namespace CloudyWing.SpreadsheetExporter.Templates;

/// <summary>
/// The merged template. Merges multiple templates into a single template.
/// </summary>
/// <seealso cref="ISheetTemplate" />
public class MergedTemplate : ISheetTemplate {
    private readonly IEnumerable<ISheetTemplate> templates;
    private TemplateLayout? cachedLayout;

    /// <summary>
    /// Initializes a new instance of the <see cref="MergedTemplate"/> class.
    /// </summary>
    /// <param name="templates">The templates to merge.</param>
    /// <exception cref="ArgumentNullException"><paramref name="templates"/> is <see langword="null"/>.</exception>
    public MergedTemplate(IEnumerable<ISheetTemplate> templates) {
        ArgumentNullException.ThrowIfNull(templates);
        this.templates = templates;
    }

    /// <summary>
    /// Initializes a new instance of the <see cref="MergedTemplate"/> class.
    /// </summary>
    /// <param name="templates">The templates to merge.</param>
    public MergedTemplate(params ISheetTemplate[] templates)
        : this(templates as IEnumerable<ISheetTemplate>) { }

    /// <summary>
    /// Gets the column span occupied by this template.
    /// </summary>
    /// <value>The number of columns.</value>
    public int ColumnSpan => GetCachedLayout().Cells.Count == 0
        ? 0
        : GetCachedLayout().Cells.Max(c => c.Point.X + c.Size.Width);

    /// <summary>
    /// Gets the row span occupied by this template.
    /// </summary>
    /// <value>The number of rows.</value>
    public int RowSpan => GetCachedLayout().RowSpan;

    /// <summary>
    /// Gets the computed cell layout for this template.
    /// </summary>
    /// <returns>The <see cref="TemplateLayout"/> for this template.</returns>
    public TemplateLayout GetLayout() {
        return GetCachedLayout();
    }

    private TemplateLayout GetCachedLayout() {
        return cachedLayout ??= TemplateLayout.Create(templates);
    }
}
