using System.Drawing;

namespace CloudyWing.SpreadsheetExporter.Templates.RecordSet {
    /// <summary>
    /// The data column base.
    /// </summary>
    /// <typeparam name="T">The type of the record.</typeparam>
    public abstract class DataColumnBase<T> {
        private DataColumnCollection<T> childColumns;

        /// <summary>
        /// Gets or sets the header text.
        /// </summary>
        /// <value>
        /// The header text.
        /// </value>
        public string HeaderText { get; set; }

        /// <summary>
        /// Gets or sets the header style.
        /// </summary>
        /// <value>
        /// The header style.
        /// </value>
        public CellStyle HeaderStyle { get; set; }

        /// <summary>
        /// Gets the point.
        /// </summary>
        /// <value>
        /// The point.
        /// </value>
        public Point Point { get; internal set; }

        /// <summary>
        /// Gets the column span.
        /// </summary>
        /// <value>
        /// The column span.
        /// </value>
        public int ColumnSpan => ChildColumns.Count == 0 ? 1 : ChildColumns.ColumnSpan;

        /// <summary>
        /// Gets the row span.
        /// </summary>
        /// <value>
        /// The row span.
        /// </value>
        public int RowSpan => ParentColumns.RowSpan - ChildColumns.RowSpan;

        /// <summary>
        /// Gets the child columns.
        /// </summary>
        /// <value>
        /// The child columns.
        /// </value>
        public DataColumnCollection<T> ChildColumns {
            get {
                return childColumns ??= new DataColumnCollection<T>(this);
            }
        }

        /// <summary>
        /// Gets the parent columns.
        /// </summary>
        /// <value>
        /// The parent columns.
        /// </value>
        public DataColumnCollection<T> ParentColumns { get; internal set; }

        /// <summary>
        /// Gets the column layers.
        /// </summary>
        /// <value>
        /// The column layers.
        /// </value>
        /// <remarks>
        /// 自己和子層共幾層 Column，用來計算 RowSpan
        /// </remarks>
        public int ColumnLayers => ChildColumns.Count == 0
            ? 1 : ChildColumns.RowSpan + 1;

        /// <summary>
        /// Gets the field value.
        /// </summary>
        /// <param name="context">The context.</param>
        /// <returns>Returns the field value.</returns>
        public abstract object GetFieldValue(RecordContext<T> context);

        /// <summary>
        /// Gets the field formula.
        /// </summary>
        /// <param name="context">The context.</param>
        /// <returns>Returns the field formula.</returns>
        public abstract string GetFieldFormula(RecordContext<T> context);

        /// <summary>
        /// Gets the field style.
        /// </summary>
        /// <param name="context">The context.</param>
        /// <returns>Returns the field style.</returns>
        public abstract CellStyle GetFieldStyle(RecordContext<T> context);
    }
}
