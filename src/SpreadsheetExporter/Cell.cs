using System.Drawing;

namespace CloudyWing.SpreadsheetExporter {
    /// <summary>The spreadsheet cell.</summary>
    public class Cell {
        /// <summary>Gets or sets the cell content value.</summary>
        /// <value>The cell content value.</value>
        public object Value { get; set; }

        /// <summary>Gets or sets the point.</summary>
        /// <value>The point.</value>
        public Point Point { get; set; }

        /// <summary>Gets or sets the cell size.</summary>
        /// <value>The cell size.</value>
        public Size Size { get; set; } = new Size(1, 1);

        /// <summary>Gets or sets the cell style.</summary>
        /// <value>The cell style.</value>
        public CellStyle CellStyle { get; set; }

        /// <summary>Gets or sets the formula.</summary>
        /// <value>The formula.</value>
        public string Formula { get; set; }

        /// <summary>Shallows the copy.</summary>
        /// <returns>The spreadsheet cell.</returns>
        public Cell ShallowCopy() {
            return (Cell)MemberwiseClone();
        }
    }
}
