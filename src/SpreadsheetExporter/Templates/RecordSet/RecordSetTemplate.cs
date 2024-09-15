using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;

namespace CloudyWing.SpreadsheetExporter.Templates.RecordSet {
    /// <summary>The recordset template. Create cell information using set data source and data column.</summary>
    /// <typeparam name="T">The type of the record.</typeparam>
    /// <seealso cref="ITemplate" />
    public class RecordSetTemplate<T> : ITemplate {
        /// <summary>Initializes a new instance of the <see cref="RecordSetTemplate{T}" /> class.</summary>
        /// <param name="dataSource">The data source.</param>
        /// <exception cref="ArgumentNullException">dataSource</exception>
        public RecordSetTemplate(IEnumerable<T> dataSource) {
            DataSource = dataSource ?? throw new ArgumentNullException(nameof(dataSource));
        }

        /// <summary>Gets or sets the data source.</summary>
        /// <value>The data source.</value>
        public IEnumerable<T> DataSource { get; set; }

        /// <summary>Gets the columns.</summary>
        /// <value>The columns.</value>
        public DataColumnCollection<T> Columns { get; } = new DataColumnCollection<T>(null);

        /// <summary>Gets or sets the height of the header.</summary>
        /// <value>The height of the header.</value>
        public double? HeaderHeight { get; set; }

        /// <summary>Gets or sets the height of the record.</summary>
        /// <value>The height of the record.</value>
        public double? RecordHeight { get; set; }

        /// <summary>Gets the column span.</summary>
        /// <value>The column span.</value>
        public int ColumnSpan => Columns.ColumnSpan;

        /// <summary>Gets the row span.</summary>
        /// <value>The row span.</value>
        public int RowSpan => DataSource.Count() + Columns.RowSpan;


        private IEnumerable<Cell> GetHearderCells(DataColumnCollection<T> cols) {
            List<Cell> cells = new();
            foreach (DataColumnBase<T> col in cols) {
                cells.Add(new Cell() {
                    ValueGenerator = (cellIndex, rowIndex) => col.HeaderText,
                    CellStyleGenerator = (cellIndex, rowIndex) => col.HeaderStyle,
                    Point = col.Point,
                    Size = new Size(col.ColumnSpan, col.RowSpan)
                });

                cells.AddRange(GetHearderCells(col.ChildColumns));
            }
            return cells;
        }

        private IEnumerable<Cell> GetRecordCells() {
            Point p = new(0, Columns.RowSpan);

            int i = 0;
            foreach (T record in DataSource) {
                foreach (DataColumnBase<T> col in Columns.DataSourceColumns) {
                    yield return new Cell {
                        ValueGenerator = (cellIndex, rowIndex) => col.GetFieldValue(new(cellIndex, rowIndex, record)),
                        CellStyleGenerator = (cellIndex, rowIndex) => col.GetFieldStyle(new(cellIndex, rowIndex, record)),
                        Point = p,
                        Size = new Size(col.ColumnSpan, 1),
                        FormulaGenerator = (cellIndex, rowIndex) => col.GetFieldFormula(new(cellIndex, rowIndex, record))
                    };
                    p += new Size(col.ColumnSpan, 0);
                }

                p = new Point(0, p.Y + 1);
                i++;
            }
        }

        /// <inheritdoc/>
        public TemplateContext GetContext() {
            return new TemplateContext(GetCells(), RowSpan, GetRowHeights());
        }

        private IEnumerable<Cell> GetCells() {
            return GetHearderCells(Columns).Union(GetRecordCells());
        }

        private Dictionary<int, double?> GetRowHeights() {
            Dictionary<int, double?> dic = new();
            int i;

            for (i = 0; i < Columns.RowSpan; i++) {
                dic.Add(i, HeaderHeight);
            }

            foreach (T item in DataSource) {
                dic.Add(i++, RecordHeight);
            }

            return dic;
        }
    }
}
