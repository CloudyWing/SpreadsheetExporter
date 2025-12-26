using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;

namespace CloudyWing.SpreadsheetExporter.Templates.RecordSet;

/// <summary>
/// The recordset template. Create cell information using set data source and data column.
/// </summary>
/// <seealso cref="ITemplate" />
/// <typeparam name="T">The type of the record.</typeparam>
/// <param name="dataSource">The data source.</param>
/// <exception cref="ArgumentNullException">dataSource</exception>
public class RecordSetTemplate<T>(IEnumerable<T> dataSource) : ITemplate {
    /// <summary>
    /// Gets or sets the data source.
    /// </summary>
    /// <value>
    /// The data source.
    /// </value>
    public IEnumerable<T> DataSource { get; set; } = dataSource ?? throw new ArgumentNullException(nameof(dataSource));

    /// <summary>
    /// Gets the columns.
    /// </summary>
    /// <value>
    /// The columns.
    /// </value>
    public DataColumnCollection<T> Columns { get; } = new DataColumnCollection<T>(null);

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
    public int ColumnSpan => Columns.ColumnSpan;

    /// <summary>
    /// Gets the row span.
    /// </summary>
    /// <value>
    /// The row span.
    /// </value>
    public int RowSpan => DataSource.Count() + Columns.RowSpan;

    private IEnumerable<Cell> GetHearderCells(DataColumnCollection<T> columns) {
        foreach (DataColumnBase<T> column in columns) {
            yield return new Cell {
                ValueGenerator = (cellIndex, rowIndex) => column.HeaderText,
                CellStyleGenerator = (cellIndex, rowIndex) => column.HeaderStyle,
                Point = column.Point,
                Size = new Size(column.ColumnSpan, column.RowSpan)
            };

            foreach (Cell childCell in GetHearderCells(column.ChildColumns)) {
                yield return childCell;
            }
        }
    }

    private IEnumerable<Cell> GetRecordCells() {
        Point point = new(0, Columns.RowSpan);
        foreach (T record in DataSource) {
            foreach (DataColumnBase<T> column in Columns.DataSourceColumns) {
                yield return new Cell {
                    ValueGenerator = (cellIndex, rowIndex) => column.GetFieldValue(new RecordContext<T>(cellIndex, rowIndex, record)),
                    CellStyleGenerator = (cellIndex, rowIndex) => column.GetFieldStyle(new RecordContext<T>(cellIndex, rowIndex, record)),
                    Point = point,
                    Size = new Size(column.ColumnSpan, 1),
                    FormulaGenerator = (cellIndex, rowIndex) => column.GetFieldFormula(new RecordContext<T>(cellIndex, rowIndex, record))
                };
                point.X += column.ColumnSpan;
            }
            point = new Point(0, point.Y + 1);
        }
    }

    /// <inheritdoc/>
    public TemplateContext GetContext() {
        Columns.CalculatePoints(Point.Empty);
        return new TemplateContext(GetCells(), RowSpan, GetRowHeights());
    }

    private IEnumerable<Cell> GetCells() {
        return GetHearderCells(Columns).Union(GetRecordCells());
    }

    private Dictionary<int, double?> GetRowHeights() {
        Dictionary<int, double?> dic = [];
        int i;
        for (i = 0; i < Columns.RowSpan; i++) {
            dic[i] = HeaderHeight;
        }
        foreach (T _ in DataSource) {
            dic[i++] = RecordHeight;
        }
        return dic;
    }
}
