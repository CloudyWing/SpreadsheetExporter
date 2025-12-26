namespace CloudyWing.SpreadsheetExporter.Templates.RecordSet;
/// <summary>
/// The record context.
/// </summary>
/// <typeparam name="T">The type of the record.</typeparam>
/// <param name="cellIndex">Index of the cell.</param>
/// <param name="rowIndex">Index of the row.</param>
/// <param name="record">The record.</param>
public class RecordContext<T>(int cellIndex, int rowIndex, T record) {
    /// <summary>
    /// Gets the index of the cell. The index start at 0.
    /// </summary>
    /// <value>
    /// The index of the cell.
    /// </value>
    public int CellIndex { get; } = cellIndex;

    /// <summary>
    /// Gets the index of the row. The index start at 0.
    /// </summary>
    /// <value>
    /// The index of the row.
    /// </value>
    public int RowIndex { get; } = rowIndex;

    /// <summary>
    /// Gets the record.
    /// </summary>
    /// <value>
    /// The record.
    /// </value>
    public T Record { get; } = record;
}
