using System.Collections.Generic;
using System.Linq;
using CloudyWing.SpreadsheetExporter.Extensions;

namespace CloudyWing.SpreadsheetExporter {
    public class TemplateContext {
        public TemplateContext(
            IEnumerable<Cell> cells, int columnSpan, int rowSpan, IReadOnlyDictionary<int, double> rowHeights
        ) {
            Cells = cells.ToList().AsReadOnly();
            ColumnSpan = columnSpan;
            RowSpan = rowSpan;
            RowHeights = rowHeights.AsReadOnly();
        }

        /// <summary>
        /// Gets the cells.
        /// </summary>
        /// <value>
        /// The cells.
        /// </value>
        public IReadOnlyList<Cell> Cells { get; }

        /// <summary>
        /// Gets the column span.
        /// </summary>
        /// <value>
        /// The column span.
        /// </value>
        public int ColumnSpan { get; }

        /// <summary>
        /// Gets the row span.
        /// </summary>
        /// <value>
        /// The row span.
        /// </value>
        public int RowSpan { get; }

        /// <summary>
        /// Gets the row heights.
        /// </summary>
        /// <value>
        /// The row heights.
        /// </value>
        public IReadOnlyDictionary<int, double> RowHeights { get; }
    }
}
