using System.Collections.Generic;

namespace CloudyWing.SpreadsheetExporter.Templates {
    /// <summary>The template interface.</summary>
    public interface ITemplate {
        /// <summary>Gets the height of rows.</summary>
        /// <value>The height of rows.</value>
        IReadOnlyDictionary<int, double> RowHeights { get; }

        /// <summary>Gets the context.</summary>
        /// <returns>The template context.</returns>
        TemplateContext GetContext();
    }
}
