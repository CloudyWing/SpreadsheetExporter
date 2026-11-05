using System;
using System.Data;

namespace CloudyWing.SpreadsheetExporter.Templates.DataTable;

/// <summary>
/// Provides the data row, position, and value of a <see cref="DataTableTemplate"/> data cell.
/// </summary>
public sealed class DataRowContext {
    /// <summary>
    /// Initializes a new instance of the <see cref="DataRowContext"/> class.
    /// </summary>
    /// <param name="cellIndex">The zero-based column index of the cell.</param>
    /// <param name="rowIndex">The zero-based row index of the cell.</param>
    /// <param name="dataRowIndex">The zero-based index of the row in <see cref="System.Data.DataTable.Rows"/>.</param>
    /// <param name="dataRow">The data row.</param>
    /// <param name="columnName">The name of the bound data column.</param>
    /// <param name="value">The cell value, or <see langword="null"/> when the data row contains <see cref="DBNull"/>.</param>
    /// <exception cref="ArgumentNullException">
    /// <paramref name="dataRow"/> or <paramref name="columnName"/> is <see langword="null"/>.
    /// </exception>
    public DataRowContext(int cellIndex, int rowIndex, int dataRowIndex, DataRow dataRow, string columnName, object? value) {
        ArgumentNullException.ThrowIfNull(dataRow);
        ArgumentNullException.ThrowIfNull(columnName);

        CellIndex = cellIndex;
        RowIndex = rowIndex;
        DataRowIndex = dataRowIndex;
        DataRow = dataRow;
        ColumnName = columnName;
        Value = value;
    }

    /// <summary>
    /// Gets the zero-based column index of the cell, as passed to the cell generator.
    /// </summary>
    public int CellIndex { get; }

    /// <summary>
    /// Gets the zero-based row index of the cell, as passed to the cell generator.
    /// </summary>
    public int RowIndex { get; }

    /// <summary>
    /// Gets the zero-based index of the row in <see cref="System.Data.DataTable.Rows"/>.
    /// </summary>
    public int DataRowIndex { get; }

    /// <summary>
    /// Gets the data row that contains the cell value.
    /// </summary>
    public DataRow DataRow { get; }

    /// <summary>
    /// Gets the name of the bound data column.
    /// </summary>
    public string ColumnName { get; }

    /// <summary>
    /// Gets the cell value, or <see langword="null"/> when the data row contains <see cref="DBNull"/>.
    /// </summary>
    public object? Value { get; }
}
