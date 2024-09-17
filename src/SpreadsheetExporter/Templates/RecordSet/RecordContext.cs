namespace CloudyWing.SpreadsheetExporter.Templates.RecordSet {
    /// <summary>
    /// The record context.
    /// </summary>
    /// <typeparam name="T">The type of the record.</typeparam>
    public class RecordContext<T> {
        /// <summary>
        /// Initializes a new instance of the <see cref="RecordContext{T}" /> class.
        /// </summary>
        /// <param name="cellIndex">Index of the cell.</param>
        /// <param name="rowIndex">Index of the row.</param>
        /// <param name="record">The record.</param>
        public RecordContext(int cellIndex, int rowIndex, T record) {
            CellIndex = cellIndex;
            RowIndex = rowIndex;
            Record = record;
        }

        /// <summary>
        /// Gets the index of the cell. The index start at 0.
        /// </summary>
        /// <value>
        /// The index of the cell.
        /// </value>
        public int CellIndex { get; }

        /// <summary>
        /// Gets the index of the row. The index start at 0.
        /// </summary>
        /// <value>
        /// The index of the row.
        /// </value>
        public int RowIndex { get; }

        /// <summary>
        /// Gets the record.
        /// </summary>
        /// <value>
        /// The record.
        /// </value>
        public T Record { get; }
    }
}
