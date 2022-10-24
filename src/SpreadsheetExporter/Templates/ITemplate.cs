using System.Collections.Generic;

namespace CloudyWing.SpreadsheetExporter.Templates {
    public interface ITemplate {
        IReadOnlyDictionary<int, double> RowHeights { get; }

        TemplateContext GetContext();
    }
}
