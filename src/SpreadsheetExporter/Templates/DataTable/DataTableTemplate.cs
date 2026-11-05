using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Linq;

namespace CloudyWing.SpreadsheetExporter.Templates.DataTable;

/// <summary>
/// The DataTable template. Create cell information using DataTable as data source.
/// </summary>
/// <seealso cref="ISheetTemplate" />
/// <param name="dataTable">The data table.</param>
/// <exception cref="ArgumentNullException">dataTable</exception>
public class DataTableTemplate(System.Data.DataTable dataTable) : ISheetTemplate, ITableTemplate {
    /// <summary>
    /// Gets or sets the data table.
    /// </summary>
    /// <value>
    /// The data table.
    /// </value>
    /// <remarks>
    /// When the new data table does not contain every column referenced by <see cref="Columns"/>,
    /// <see cref="Columns"/> is rebuilt in place from the new data table's columns and previous column settings are discarded.
    /// Otherwise <see cref="Columns"/> is kept unchanged.
    /// </remarks>
    public System.Data.DataTable DataTable {
        get;
        set {
            field = value;
            if (value is not null && !ContainsAllColumns(value)) {
                RebuildColumns(value);
            }
        }
    } = dataTable ?? throw new ArgumentNullException(nameof(dataTable));

    /// <summary>
    /// Gets the columns.
    /// </summary>
    /// <value>
    /// The columns.
    /// </value>
    public DataTableColumnCollection Columns { get; } = InitializeColumns(dataTable);

    private static DataTableColumnCollection InitializeColumns(System.Data.DataTable dataTable) {
        DataTableColumnCollection columns = [];
        AddDefaultColumns(columns, dataTable);

        return columns;
    }

    private static void AddDefaultColumns(DataTableColumnCollection columns, System.Data.DataTable dataTable) {
        foreach (DataColumn column in dataTable.Columns) {
            columns.Add(new DataTableColumn {
                ColumnName = column.ColumnName,
                HeaderText = column.ColumnName
            });
        }
    }

    private bool ContainsAllColumns(System.Data.DataTable dataTable) {
        return Columns.All(column => column.ColumnName is null || dataTable.Columns.Contains(column.ColumnName));
    }

    private void RebuildColumns(System.Data.DataTable dataTable) {
        Columns.Clear();
        AddDefaultColumns(Columns, dataTable);
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

    private IEnumerable<Cell> GetRecordCells(IReadOnlyList<DataColumn> dataColumns) {
        int rowIndex = 1;
        foreach (DataRow row in DataTable.Rows) {
            for (int colIndex = 0; colIndex < Columns.Count; colIndex++) {
                DataTableColumn column = Columns[colIndex];
                object? value = row[dataColumns[colIndex]];

                if (value == DBNull.Value) {
                    value = null;
                }

                Cell cell = new() {
                    CellStyleGenerator = (cellIndex, rowIdx) => column.FieldStyleGenerator?.Invoke(value)
                        ?? SpreadsheetManager.DefaultCellStyles.FieldStyle,
                    Point = new Point(colIndex, rowIndex),
                    Size = new Size(1, 1)
                };

                if (column.FieldFormulaGenerator is not null) {
                    cell.FormulaGenerator = (cellIndex, rowIdx) => column.FieldFormulaGenerator.Invoke(value);
                } else {
                    cell.ValueGenerator = (cellIndex, rowIdx) => column.FieldValueGenerator?.Invoke(value) ?? value;
                }

                if (column.FieldDataValidationGenerator is not null) {
                    cell.DataValidationGenerator = (cellIndex, rowIdx) =>
                        column.FieldDataValidationGenerator.Invoke(value);
                }

                yield return cell;
            }
            rowIndex++;
        }
    }

    /// <inheritdoc/>
    public TemplateLayout GetLayout() {
        ValidateColumns();
        IReadOnlyList<DataColumn> dataColumns = ResolveDataColumns();

        return new TemplateLayout(GetCells(dataColumns), RowSpan, GetRowHeights());
    }

    private void ValidateColumns() {
        foreach (DataTableColumn column in Columns) {
            if (column.FieldValueGenerator is not null && column.FieldFormulaGenerator is not null) {
                throw new InvalidOperationException(
                    $"Cannot use both FieldValueGenerator and FieldFormulaGenerator on DataTable column '{column.ColumnName}'. "
                    + "A cell can only have a value or a formula, not both."
                );
            }
        }
    }

    private List<DataColumn> ResolveDataColumns() {
        List<DataColumn> dataColumns = [];

        // 沒有資料列時不會讀取欄位值，維持不驗證欄位名稱
        if (DataTable.Rows.Count == 0) {
            return dataColumns;
        }

        foreach (DataTableColumn column in Columns) {
            DataColumn? dataColumn = column.ColumnName is null ? null : DataTable.Columns[column.ColumnName];
            if (dataColumn is null) {
                throw new ArgumentException(
                    $"DataTable '{DataTable.TableName}' does not contain column '{column.ColumnName}'. "
                    + $"Available columns: {string.Join(", ", DataTable.Columns.Cast<DataColumn>().Select(x => x.ColumnName))}"
                );
            }

            dataColumns.Add(dataColumn);
        }

        return dataColumns;
    }

    private IEnumerable<Cell> GetCells(IReadOnlyList<DataColumn> dataColumns) {
        return GetHeaderCells().Concat(GetRecordCells(dataColumns));
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
