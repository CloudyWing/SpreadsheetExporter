using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using CloudyWing.SpreadsheetExporter.Extensions;

namespace CloudyWing.SpreadsheetExporter.Templates {
    /// <summary>The template context.</summary>
    public class TemplateContext {
        /// <summary>Initializes a new instance of the <see cref="TemplateContext" /> class.</summary>
        /// <param name="cells">The cells.</param>
        /// <param name="rowSpan">The row span.</param>
        /// <param name="rowHeights">The row heights.</param>
        public TemplateContext(
            IEnumerable<Cell> cells, int rowSpan, IReadOnlyDictionary<int, double> rowHeights
        ) {
            Cells = cells.ToList().AsReadOnly();
            RowSpan = rowSpan;
            RowHeights = rowHeights.AsReadOnly();
        }

        /// <summary>Gets the cells.</summary>
        /// <value>The cells.</value>
        public IReadOnlyList<Cell> Cells { get; }

        /// <summary>Gets the row span.</summary>
        /// <value>The row span.</value>
        public int RowSpan { get; }

        /// <summary>Gets the height of rows.</summary>
        /// <value>The height of rows.</value>
        public IReadOnlyDictionary<int, double> RowHeights { get; }

        internal static TemplateContext Create(IEnumerable<ITemplate> templates) {
            if (templates.Count() == 1) {
                return templates.Single().GetContext();
            }

            List<Cell> cells = new();
            Dictionary<int, double> rowHeights = new();
            int rowSpan = 0;
            int rowIndex = 0;

            foreach (TemplateContext context in templates.Select(x => x.GetContext())) {
                foreach (KeyValuePair<int, double> pair in context.RowHeights) {
                    rowHeights.Add(rowIndex + pair.Key, pair.Value);
                }

                foreach (Cell cell in context.Cells) {
                    Cell fixedCell = cell.ShallowCopy();
                    fixedCell.Point = new System.Drawing.Point() {
                        X = cell.Point.X,
                        Y = cell.Point.Y + rowIndex
                    };

                    Debug.Assert(fixedCell.Point.Y == cell.Point.Y + rowIndex);

                    cells.Add(fixedCell);
                }

                rowIndex += context.RowSpan;
                rowSpan += context.RowSpan;
            }

            return new TemplateContext(cells, rowSpan, rowHeights);
        }
    }
}
