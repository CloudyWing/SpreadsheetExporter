using System;
using System.Collections;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Drawing;
using System.Linq;

namespace CloudyWing.SpreadsheetExporter.Templates.Grid {
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
        /// columnSpan - Must be greater than 0.
        /// or
        /// rowSpan - Must be greater than 0.
        /// </exception>
        public GridTemplate CreateCell(
            object value, int columnSpan = 1, int rowSpan = 1, CellStyle? cellStyle = null
        ) {
            ValidateCellParameters(columnSpan, rowSpan);
            cellStyle ??= SpreadsheetManager.DefaultCellStyles.GridCellStyle;

            if (rows.Count == 0) {
                CreateRow();
            }

            AddCell(new Cell {
                ValueGenerator = (cellIndex, rowIndex) => value,
                Size = new Size(columnSpan, rowSpan),
                CellStyleGenerator = (cellIndex, rowIndex) => (CellStyle)cellStyle
            });

            return this;
        }

        /// <summary>
        /// Create cell that contain formula.
        /// </summary>
        /// <param name="formulaGenerator">The formula generator. Pass the cell index and row index to the generator. The  index start at 0.</param>
        /// <param name="columnSpan">The column span.</param>
        /// <param name="rowSpan">The row span.</param>
        /// <param name="cellStyle">The cell style. The default is <c>SpreadsheetManager.DefaultCellStyles.GridCellStyle</c>.</param>
        /// <returns>The self.</returns>
        /// <exception cref="ArgumentOutOfRangeException">columnSpan - Must be greater than 0.
        /// or
        /// rowSpan - Must be greater than 0.</exception>
        public GridTemplate CreateCell(
            Func<int, int, string> formulaGenerator, int columnSpan = 1, int rowSpan = 1, CellStyle? cellStyle = null
        ) {
            ValidateCellParameters(columnSpan, rowSpan);
            cellStyle ??= SpreadsheetManager.DefaultCellStyles.GridCellStyle;

            if (rows.Count == 0) {
                CreateRow();
            }

            AddCell(new Cell {
                Size = new Size(columnSpan, rowSpan),
                CellStyleGenerator = (cellIndex, rowIndex) => (CellStyle)cellStyle,
                FormulaGenerator = formulaGenerator
            });

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

        private void AddCell(Cell cell) {
            CellCollection lastRow = rows.Last();
            lastRow.Add(cell);
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

            internal CellCollection(GridTemplate grid) {
                this.grid = grid;
            }

            public void Add(Cell item) {
                items.Add(item);

                item.Point = new Point(0, grid.rows.Count - 1);

                while (grid.points.Contains(item.Point)) {
                    item.Point += new Size(1, 0);
                }
                grid.points.Add(item.Point);

                for (int colOffset = 0; colOffset < item.Size.Width; colOffset++) {
                    for (int rowOffset = 0; rowOffset < item.Size.Height; rowOffset++) {
                        // 因為是自己，加入過，所以不用理會
                        if (colOffset == 0 && rowOffset == 0) {
                            continue;
                        }
                        Point point = item.Point + new Size(colOffset, rowOffset);
                        if (!grid.points.Add(point)) {
                            throw new ArgumentException($"The point ({colOffset}, {rowOffset}) already exists.");
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
}
