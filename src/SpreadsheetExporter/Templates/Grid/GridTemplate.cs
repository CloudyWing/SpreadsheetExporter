using System;
using System.Collections;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Drawing;
using System.Linq;

namespace CloudyWing.SpreadsheetExporter.Templates.Grid {
    /// <summary>The grid template. Create cell information using <c>CreateRow()</c> and <c>CreateCell</c>.</summary>
    public class GridTemplate : ITemplate {
        private readonly IList<CellCollection> rows = new List<CellCollection>();
        private readonly IList<Point> points = new List<Point>();
        private readonly IDictionary<int, double> rowHeights = new Dictionary<int, double>();

        /// <summary>Gets the column span.</summary>
        /// <value>The column span.</value>
        public int ColumnSpan => points.Count == 0 ? 0 : points.Max(x => x.X) + 1;

        /// <summary>Gets the row span.</summary>
        /// <value>The row span.</value>
        public int RowSpan => points.Count == 0 ? 0 : points.Max(x => x.Y) + 1;

        /// <summary>Gets the cells.</summary>
        /// <value>The cells.</value>
        public IEnumerable<Cell> Cells {
            get {
                foreach (CellCollection cells in rows) {
                    foreach (Cell cell in cells) {
                        yield return cell;
                    }
                }
            }
        }

        public IReadOnlyDictionary<int, double> RowHeights => new ReadOnlyDictionary<int, double>(rowHeights);

        /// <summary>Creates the row.</summary>
        /// <param name="height">The height.</param>
        public void CreateRow(double height = 16.5d) {
            // 避免建立最後一筆 Row，卻沒加入 Cell 導致用作標算列數會不正確，所以加一個 -x 座標
            points.Add(new Point(-1, rows.Count));
            rowHeights.Add(rows.Count, height);
            rows.Add(new CellCollection(this));
        }

        /// <summary>Creates the cell.</summary>
        /// <param name="value">The value.</param>
        /// <param name="columnSpan">The column span.</param>
        /// <param name="rowSpan">The row span.</param>
        /// <param name="cellStyle">The cell style.</param>
        /// <returns>The cell.</returns>
        /// <exception cref="ArgumentOutOfRangeException">columnSpan - Must be greater than 0.
        /// or
        /// rowSpan - Must be greater than 0.</exception>
        public Cell CreateCell(
            object value, int columnSpan = 1, int rowSpan = 1, CellStyle? cellStyle = null
        ) {
            if (columnSpan <= 0) {
                throw new ArgumentOutOfRangeException(nameof(columnSpan), "Must be greater than 0.");
            }
            if (rowSpan <= 0) {
                throw new ArgumentOutOfRangeException(nameof(rowSpan), "Must be greater than 0.");
            }
            cellStyle ??= SpreadsheetManager.Configuration.CellStyle;

            if (rows.Count == 0) {
                CreateRow();
            }

            CellCollection lastRow = rows.Last();
            Cell cell = new() {
                Value = value,
                Size = new Size(columnSpan, rowSpan),
                CellStyle = (CellStyle)cellStyle
            };
            lastRow.Add(cell);

            return cell;
        }

        /// <summary>Create cell that contain formula.</summary>
        /// <param name="formulaGenerator">The formula generator. Pass the row index and cell index to the generator. The  index start at 0.</param>
        /// <param name="columnSpan">The column span.</param>
        /// <param name="rowSpan">The row span.</param>
        /// <param name="cellStyle">The cell style.</param>
        /// <returns>The cell.</returns>
        /// <exception cref="ArgumentOutOfRangeException">columnSpan - Must be greater than 0.
        /// or
        /// rowSpan - Must be greater than 0.</exception>
        public Cell CreateCell(
            Func<int, int, string> formulaGenerator, int columnSpan = 1, int rowSpan = 1, CellStyle? cellStyle = null
        ) {
            if (columnSpan <= 0) {
                throw new ArgumentOutOfRangeException(nameof(columnSpan), "Must be greater than 0.");
            }
            if (rowSpan <= 0) {
                throw new ArgumentOutOfRangeException(nameof(rowSpan), "Must be greater than 0.");
            }
            cellStyle ??= SpreadsheetManager.Configuration.CellStyle;

            if (rows.Count == 0) {
                CreateRow();
            }

            CellCollection lastRow = rows.Last();
            Cell cell = new() {
                Size = new Size(columnSpan, rowSpan),
                CellStyle = (CellStyle)cellStyle,
                Formula = formulaGenerator(rows.Count - 1, lastRow.Count())
            };
            lastRow.Add(cell);

            return cell;
        }

        public TemplateContext GetContext() {
            return new TemplateContext(
                Cells, ColumnSpan, RowSpan, RowHeights.ToDictionary(pair => pair.Key, pair => pair.Value)
            );
        }

        private class CellCollection : IEnumerable<Cell>, IEnumerable {
            private readonly GridTemplate grid;
            private readonly IList<Cell> items = new List<Cell>();

            internal CellCollection(GridTemplate grid) {
                this.grid = grid;
            }

            public void Add(Cell item) {
                items.Add(item);

                item.Point = new Point(0, grid.rows.Count - 1);

                while (IsPointExists(item.Point)) {
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
                        if (IsPointExists(point)) {
                            throw new ArgumentException($"The point ({colOffset}, {rowOffset}) already exist.");
                        }
                        grid.points.Add(point);
                    }
                }
            }

            private bool IsPointExists(Point point) {
                return grid.points.Contains(point);
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
