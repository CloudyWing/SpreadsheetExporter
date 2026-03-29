using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using CloudyWing.SpreadsheetExporter.Util;

namespace CloudyWing.SpreadsheetExporter.Templates;

/// <summary>
/// Represents the computed cell layout produced by a template, ready for rendering.
/// </summary>
/// <param name="cells">The cells.</param>
/// <param name="rowSpan">The row span.</param>
/// <param name="rowHeights">The row heights.</param>
public class TemplateLayout(
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

    internal static TemplateLayout Create(IEnumerable<ISheetTemplate> templates) {
        List<ISheetTemplate> templatesList = templates.ToList();

        if (templatesList.Count == 1) {
            return templatesList.Single().GetLayout();
        }

        List<Cell> cells = [];
        Dictionary<int, double?> rowHeights = [];
        int rowSpan = 0;
        int rowIndex = 0;

        foreach (TemplateLayout layout in templatesList.Select(x => x.GetLayout())) {
            foreach (KeyValuePair<int, double?> pair in layout.RowHeights) {
                rowHeights[rowIndex + pair.Key] = pair.Value;
            }

            foreach (Cell cell in layout.Cells) {
                Cell fixedCell = cell.ShallowCopy();
                fixedCell.Point = new Point {
                    X = cell.Point.X,
                    Y = cell.Point.Y + rowIndex
                };

                Debug.Assert(fixedCell.Point.Y == cell.Point.Y + rowIndex);

                cells.Add(fixedCell);
            }

            rowIndex += layout.RowSpan;
            rowSpan += layout.RowSpan;
        }

        return new TemplateLayout(cells, rowSpan, rowHeights);
    }
}
