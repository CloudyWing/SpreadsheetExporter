using System;
using System.Drawing;

namespace CloudyWing.SpreadsheetExporter {
    /// <summary>The spreadsheet cell.</summary>
    public class Cell {
        /// <summary>Gets or sets the cell content value generator. Pass the cell index and row index to the generator. The  index start at 0.</summary>
        /// <value>The cell content value.</value>
        public Func<int, int, object> ValueGenerator { get; set; }

        /// <summary>Gets or sets the cell style generator. Pass the cell index and row index to the generator. The  index start at 0.</summary>
        /// <value>The cell style.</value>
        public Func<int, int, CellStyle> CellStyleGenerator { get; set; }

        /// <summary>Gets or sets the formula generator. Pass the cell index and row index to the generator. The  index start at 0.</summary>
        /// <value>The formula generator.</value>
        public Func<int, int, string> FormulaGenerator { get; set; }

        /// <summary>Gets or sets the point.</summary>
        /// <value>The point.</value>
        public Point Point { get; set; } // 本來考慮把 setter 改成 internal，但考量有開放自訂 Template，所以還是維持 public

        /// <summary>Gets or sets the cell size.</summary>
        /// <value>The cell size.</value>
        public Size Size { get; set; } = new Size(1, 1);

        /// <summary>Gets the content value.</summary>
        /// <returns>The content value.</returns>
        public object GetValue() {
            return ValueGenerator?.Invoke(Point.X, Point.Y);
        }

        /// <summary>Gets the cell style.</summary>
        /// <returns>The cell style.</returns>
        public CellStyle GetCellStyle() {
            return CellStyleGenerator?.Invoke(Point.X, Point.Y) ?? CellStyle.Empty;
        }

        /// <summary>Gets the formula. if formula starts with <c>=</c>, automatically removed.</summary>
        /// <returns>The the formula.</returns>
        public string GetFormula() {
            return FormulaGenerator?.Invoke(Point.X, Point.Y)?.TrimStart(' ').TrimStart('=');
        }

        /// <summary>Shallows the copy.</summary>
        /// <returns>The spreadsheet cell.</returns>
        public Cell ShallowCopy() {
            return (Cell)MemberwiseClone();
        }
    }
}
