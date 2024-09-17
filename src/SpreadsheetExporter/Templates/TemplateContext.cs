﻿using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using CloudyWing.SpreadsheetExporter.Extensions;

namespace CloudyWing.SpreadsheetExporter.Templates {
    /// <summary>
    /// The template context.
    /// </summary>
    /// <param name="cells">The cells.</param>
    /// <param name="rowSpan">The row span.</param>
    /// <param name="rowHeights">The row heights.</param>
    public class TemplateContext(
        IEnumerable<Cell> cells, int rowSpan, IReadOnlyDictionary<int, double?> rowHeights
    ) {
        /// <summary>
        /// Gets the cells.
        /// </summary>
        /// <value>
        /// The cells.
        /// </value>
        public IReadOnlyList<Cell> Cells { get; } = cells.AsReadOnly();

        /// <summary>
        /// Gets the row span.
        /// </summary>
        /// <value>
        /// The row span.
        /// </value>
        public int RowSpan { get; } = rowSpan;

        /// <summary>
        /// Gets the height of rows.
        /// </summary>
        /// <value>
        /// The height of rows.
        /// </value>
        public IReadOnlyDictionary<int, double?> RowHeights { get; } = rowHeights.AsReadOnly();

        internal static TemplateContext Create(IEnumerable<ITemplate> templates) {
            List<ITemplate> templatesList = templates.ToList();

            if (templatesList.Count == 1) {
                return templatesList.Single().GetContext();
            }

            List<Cell> cells = [];
            Dictionary<int, double?> rowHeights = [];
            int rowSpan = 0;
            int rowIndex = 0;

            foreach (TemplateContext context in templatesList.Select(x => x.GetContext())) {
                foreach (KeyValuePair<int, double?> pair in context.RowHeights) {
                    rowHeights[rowIndex + pair.Key] = pair.Value;
                }

                foreach (Cell cell in context.Cells) {
                    Cell fixedCell = cell.ShallowCopy();
                    fixedCell.Point = new System.Drawing.Point {
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
