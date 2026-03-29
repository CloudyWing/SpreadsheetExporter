using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;

namespace CloudyWing.SpreadsheetExporter.Templates.RecordSet;

/// <summary>
/// The recordset template. Create cell information using set data source and data column.
/// </summary>
/// <seealso cref="ISheetTemplate" />
/// <typeparam name="T">The type of the record.</typeparam>
public class RecordSetTemplate<T> : ISheetTemplate {
    private IEnumerable<T> dataSourceField;

    /// <summary>
    /// Initializes a new instance of the <see cref="RecordSetTemplate{T}"/> class.
    /// </summary>
    /// <param name="dataSource">The data source.</param>
    /// <exception cref="ArgumentNullException"><paramref name="dataSource"/> is <see langword="null"/>.</exception>
    public RecordSetTemplate(IEnumerable<T> dataSource) {
        ArgumentNullException.ThrowIfNull(dataSource);
        dataSourceField = dataSource;
    }

    internal RecordSetTemplate() : this([]) {
    }

    /// <summary>
    /// Gets or sets the data source.
    /// </summary>
    /// <value>The data source.</value>
    public IEnumerable<T> DataSource {
        get => dataSourceField;
        set {
            dataSourceField = value;
            CachedDataSource = null;
        }
    }

    private IReadOnlyList<T>? CachedDataSource {
        get {
            field ??= DataSource is IReadOnlyList<T> readOnlyList
                    ? readOnlyList
                    : DataSource.ToList().AsReadOnly();

            return field;
        }
        set;
    }

    /// <summary>
    /// Gets the columns.
    /// </summary>
    /// <value>The columns.</value>
    public RecordSetColumnCollection<T> Columns { get; } = new RecordSetColumnCollection<T>(null);

    /// <summary>
    /// Gets or sets the height of the header.
    /// </summary>
    /// <value>The height of the header.</value>
    public double? HeaderHeight { get; set; }

    /// <summary>
    /// Gets or sets the height of the record.
    /// </summary>
    /// <value>The height of the record.</value>
    public double? RecordHeight { get; set; }

    /// <summary>
    /// Gets the column span.
    /// </summary>
    /// <value>The column span.</value>
    public int ColumnSpan => Columns.ColumnSpan;

    /// <summary>
    /// Gets the row span.
    /// </summary>
    /// <value>The row span.</value>
    public int RowSpan => CachedDataSource!.Count + Columns.RowSpan;

    private static IEnumerable<Cell> GetHeaderCells(RecordSetColumnCollection<T> columns) {
        foreach (RecordSetColumnBase<T> column in columns) {
            yield return new Cell {
                ValueGenerator = (cellIndex, rowIndex) => column.HeaderText,
                CellStyleGenerator = (cellIndex, rowIndex) => column.HeaderStyle,
                Point = column.Point,
                Size = new Size(column.ColumnSpan, column.RowSpan)
            };

            foreach (Cell childCell in GetHeaderCells(column.ChildColumns)) {
                yield return childCell;
            }
        }
    }

    private IEnumerable<Cell> GetRecordCells() {
        Point point = new(0, Columns.RowSpan);
        foreach (T record in CachedDataSource!) {
            foreach (RecordSetColumnBase<T> column in Columns.LeafColumns) {
                yield return new Cell {
                    ValueGenerator = (cellIndex, rowIndex) => column.GetFieldValue(new RecordContext<T>(cellIndex, rowIndex, record)),
                    CellStyleGenerator = (cellIndex, rowIndex) => column.GetFieldStyle(new RecordContext<T>(cellIndex, rowIndex, record)),
                    Point = point,
                    Size = new Size(column.ColumnSpan, 1),
                    FormulaGenerator = (cellIndex, rowIndex) => column.GetFieldFormula(new RecordContext<T>(cellIndex, rowIndex, record)),
                    DataValidationGenerator = (cellIndex, rowIndex) => column.GetFieldDataValidation(new RecordContext<T>(cellIndex, rowIndex, record))
                };
                point.X += column.ColumnSpan;
            }
            point = new Point(0, point.Y + 1);
        }
    }

    /// <inheritdoc/>
    public TemplateLayout GetLayout() {
        Columns.CalculatePoints(Point.Empty);
        return new TemplateLayout(GetCells(), RowSpan, GetRowHeights());
    }

    private IEnumerable<Cell> GetCells() {
        return RecordSetTemplate<T>.GetHeaderCells(Columns).Union(GetRecordCells());
    }

    private Dictionary<int, double?> GetRowHeights() {
        Dictionary<int, double?> dic = [];
        int i;
        for (i = 0; i < Columns.RowSpan; i++) {
            dic[i] = HeaderHeight;
        }
        foreach (T _ in CachedDataSource!) {
            dic[i++] = RecordHeight;
        }
        return dic;
    }

}
