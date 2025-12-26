using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Linq;

namespace CloudyWing.SpreadsheetExporter.Templates.DataTable;

/// <summary>
/// The DataTable template. Create cell information using DataTable as data source.
/// </summary>
/// <seealso cref="ITemplate" />
/// <param name="dataTable">The data table.</param>
/// <exception cref="ArgumentNullException">dataTable</exception>
public class DataTableTemplate(System.Data.DataTable dataTable) : ITemplate {
    /// <summary>
    /// Gets or sets the data table.
    /// </summary>
    /// <value>
    /// The data table.
    /// </value>
    public System.Data.DataTable DataTable { get; set; } = dataTable ?? throw new ArgumentNullException(nameof(dataTable));

    /// <summary>
    /// Gets the columns.
    /// </summary>
    /// <value>
    /// The columns.
    /// </value>
    public DataTableColumnCollection Columns { get; } = InitializeColumns(dataTable);

    private static DataTableColumnCollection InitializeColumns(System.Data.DataTable dataTable) {
        DataTableColumnCollection columns = [];
        foreach (DataColumn column in dataTable.Columns) {
            columns.Add(new DataTableColumn {
                ColumnName = column.ColumnName,
                HeaderText = column.ColumnName
            });
        }
        return columns;
    }

    /// <summary>
    /// Gets or sets the height of the header.
    /// </summary>
    /// <value>
    /// The height of the header.
    /// </value>
    public double? HeaderHeight { get; set; }

    /// <summary>
    /// Gets or sets the height of the record.
    /// </summary>
    /// <value>
    /// The height of the record.
    /// </value>
    public double? RecordHeight { get; set; }

    /// <summary>
    /// Gets the column span.
    /// </summary>
    /// <value>
    /// The column span.
    /// </value>
    public int ColumnSpan => Columns.Count;

    /// <summary>
    /// Gets the row span.
    /// </summary>
    /// <value>
    /// The row span.
    /// </value>
    public int RowSpan => DataTable.Rows.Count + 1;

    private IEnumerable<Cell> GetHeaderCells() {
        for (int i = 0; i < Columns.Count; i++) {
            DataTableColumn column = Columns[i];
            yield return new Cell {
                ValueGenerator = (cellIndex, rowIndex) => column.HeaderText,
                CellStyleGenerator = (cellIndex, rowIndex) => column.HeaderStyle,
                Point = new Point(i, 0),
                Size = new Size(1, 1)
            };
        }
    }

    private IEnumerable<Cell> GetRecordCells() {
        int rowIndex = 1;
        foreach (DataRow row in DataTable.Rows) {
            for (int colIndex = 0; colIndex < Columns.Count; colIndex++) {
                DataTableColumn column = Columns[colIndex];
                object value = row[column.ColumnName];

                if (value == DBNull.Value) {
                    value = null;
                }

                yield return new Cell {
                    ValueGenerator = (cellIndex, rowIdx) => column.FieldValueGenerator?.Invoke(value) ?? value,
                    CellStyleGenerator = (cellIndex, rowIdx) => column.FieldStyleGenerator?.Invoke(value) ?? SpreadsheetManager.DefaultCellStyles.FieldStyle,
                    Point = new Point(colIndex, rowIndex),
                    Size = new Size(1, 1),
                    FormulaGenerator = (cellIndex, rowIdx) => column.FieldFormulaGenerator?.Invoke(value)
                };
            }
            rowIndex++;
        }
    }

    /// <inheritdoc/>
    public TemplateContext GetContext() {
        return new TemplateContext(GetCells(), RowSpan, GetRowHeights());
    }

    private IEnumerable<Cell> GetCells() {
        return GetHeaderCells().Union(GetRecordCells());
    }

    private Dictionary<int, double?> GetRowHeights() {
        Dictionary<int, double?> dic = [];

        dic[0] = HeaderHeight;

        for (int i = 0; i < DataTable.Rows.Count; i++) {
            dic[i + 1] = RecordHeight;
        }

        return dic;
    }
}
