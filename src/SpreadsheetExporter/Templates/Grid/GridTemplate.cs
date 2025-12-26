using System;
using System.Collections;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Drawing;
using System.Linq;

namespace CloudyWing.SpreadsheetExporter.Templates.Grid;

/// <summary>
/// The grid template. Create cell information using <c>CreateRow()</c> and <c>CreateCell()</c>.
/// </summary>
public class GridTemplate : ITemplate {
    private readonly List<CellCollection> rows = [];
    private readonly HashSet<Point> points = [];
    private readonly Dictionary<int, double?> rowHeights = [];

    /// <summary>
    /// Gets the column span.
    /// </summary>
    /// <value>
    /// The column span.
    /// </value>
    public int ColumnSpan => points.Count == 0 ? 0 : points.Max(x => x.X) + 1;

    /// <summary>
    /// Gets the row span.
    /// </summary>
    /// <value>
    /// The row span.
    /// </value>
    public int RowSpan => points.Count == 0 ? 0 : points.Max(x => x.Y) + 1;

    /// <summary>
    /// Creates the row.
    /// </summary>
    /// <param name="height">The height.</param>
    /// <returns>The self.</returns>
    public GridTemplate CreateRow(double? height = null) {
        // 避免建立最後一筆 Row，卻沒加入 Cell 導致用作標算列數會不正確，所以加一個 -x 座標
        points.Add(new Point(-1, rows.Count));
        rowHeights.Add(rows.Count, height);
        rows.Add(new CellCollection(this));

        return this;
    }

    /// <summary>
    /// Creates the cell.
    /// </summary>
    /// <param name="value">The value.</param>
    /// <param name="columnSpan">The column span.</param>
    /// <param name="rowSpan">The row span.</param>
    /// <param name="cellStyle">The cell style. The default is <c>SpreadsheetManager.DefaultCellStyles.GridCellStyle</c>.</param>
    /// <returns>The self.</returns>
    /// <exception cref="ArgumentOutOfRangeException">
    /// <paramref name="columnSpan"/> must be greater than 0.
    /// </exception>
    /// <exception cref="ArgumentOutOfRangeException">
    /// <paramref name="rowSpan"/> must be greater than 0.
    /// </exception>
    public GridTemplate CreateCell(
        object value, int columnSpan = 1, int rowSpan = 1, CellStyle? cellStyle = null
    ) {
        return CreateCell(
            cell => {
                cell.ValueGenerator = (cellIndex, rowIndex) => value;
            },
            columnSpan, rowSpan, cellStyle
        );
    }

    /// <summary>
    /// Create cell that contain formula.
    /// </summary>
    /// <param name="formulaGenerator">The formula generator. Pass the cell index and row index to the generator. The  index start at 0.</param>
    /// <param name="columnSpan">The column span.</param>
    /// <param name="rowSpan">The row span.</param>
    /// <param name="cellStyle">The cell style. The default is <c>SpreadsheetManager.DefaultCellStyles.GridCellStyle</c>.</param>
    /// <returns>The self.</returns>
    /// <exception cref="ArgumentOutOfRangeException">
    /// <paramref name="columnSpan"/> must be greater than 0.
    /// </exception>
    /// <exception cref="ArgumentOutOfRangeException">
    /// <paramref name="rowSpan"/> must be greater than 0.
    /// </exception>
    public GridTemplate CreateCell(
        Func<int, int, string> formulaGenerator, int columnSpan = 1, int rowSpan = 1, CellStyle? cellStyle = null
    ) {
        return CreateCell(
            cell => {
                cell.FormulaGenerator = formulaGenerator;
            },
            columnSpan, rowSpan, cellStyle
        );
    }

    /// <summary>
    /// Creates a cell with custom configuration using a configurator action.
    /// This method provides full flexibility for configuring cell properties including value, formula, data validation, and style.
    /// </summary>
    /// <param name="configurator">The action to configure the cell. The action receives the cell instance to configure.</param>
    /// <param name="columnSpan">The column span.</param>
    /// <param name="rowSpan">The row span.</param>
    /// <param name="cellStyle">The cell style. The default is <c>SpreadsheetManager.DefaultCellStyles.GridCellStyle</c>.</param>
    /// <returns>The self.</returns>
    /// <exception cref="ArgumentOutOfRangeException">
    /// <paramref name="columnSpan"/> must be greater than 0.
    /// </exception>
    /// <exception cref="ArgumentOutOfRangeException">
    /// <paramref name="rowSpan"/> must be greater than 0.
    /// </exception>
    public GridTemplate CreateCell(
        Action<Cell> configurator, int columnSpan = 1, int rowSpan = 1, CellStyle? cellStyle = null
    ) {
        ValidateCellParameters(columnSpan, rowSpan);
        cellStyle ??= SpreadsheetManager.DefaultCellStyles.GridCellStyle;

        if (rows.Count == 0) {
            CreateRow();
        }

        Cell cell = new Cell {
            Size = new Size(columnSpan, rowSpan),
            CellStyleGenerator = (cellIndex, rowIndex) => (CellStyle)cellStyle
        };

        configurator?.Invoke(cell);

        CellCollection lastRow = rows.Last();
        lastRow.Add(cell);

        return this;
    }

    private static void ValidateCellParameters(int columnSpan, int rowSpan) {
        if (columnSpan <= 0) {
            throw new ArgumentOutOfRangeException(nameof(columnSpan), "Must be greater than 0.");
        }
        if (rowSpan <= 0) {
            throw new ArgumentOutOfRangeException(nameof(rowSpan), "Must be greater than 0.");
        }
    }

    /// <inheritdoc/>
    public TemplateContext GetContext() {
        return new TemplateContext(
            GetCells(), RowSpan, new ReadOnlyDictionary<int, double?>(rowHeights)
        );
    }

    private IEnumerable<Cell> GetCells() {
        foreach (CellCollection cells in rows) {
            foreach (Cell cell in cells) {
                yield return cell;
            }
        }
    }

    private class CellCollection : IEnumerable<Cell>, IEnumerable {
        private readonly GridTemplate grid;
        private readonly List<Cell> items = [];
        private int nextAvailableX = 0;

        internal CellCollection(GridTemplate grid) {
            this.grid = grid;
        }

        public void Add(Cell item) {
            items.Add(item);

            item.Point = new Point(nextAvailableX, grid.rows.Count - 1);

            while (grid.points.Contains(item.Point)) {
                item.Point += new Size(1, 0);
            }
            grid.points.Add(item.Point);
            nextAvailableX = item.Point.X + item.Size.Width;

            for (int colOffset = 0; colOffset < item.Size.Width; colOffset++) {
                for (int rowOffset = 0; rowOffset < item.Size.Height; rowOffset++) {
                    // 因為是自己，加入過，所以不用理會
                    if (colOffset == 0 && rowOffset == 0) {
                        continue;
                    }
                    Point point = item.Point + new Size(colOffset, rowOffset);
                    if (!grid.points.Add(point)) {
                        int actualX = item.Point.X + colOffset;
                        int actualY = item.Point.Y + rowOffset;
                        throw new ArgumentException(
                            $"Cell merge conflict: The point ({actualX}, {actualY}) is already occupied. " +
                            $"Current cell at ({item.Point.X}, {item.Point.Y}) with size ({item.Size.Width}, {item.Size.Height}) overlaps with existing cells."
                        );
                    }
                }
            }
        }

        IEnumerator<Cell> IEnumerable<Cell>.GetEnumerator() {
            return items.GetEnumerator();
        }

        IEnumerator IEnumerable.GetEnumerator() {
            return ((IEnumerable)items).GetEnumerator();
        }
    }
}
