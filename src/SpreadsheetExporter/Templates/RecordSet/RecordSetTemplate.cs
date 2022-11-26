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
        /// <exception cref="System.ArgumentNullException">dataSource</exception>
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
        public double HeaderHeight { get; set; } = 16.5d;

        /// <summary>Gets or sets the height of the item.</summary>
        /// <value>The height of the item.</value>
        public double ItemHeight { get; set; } = 16.5d;

        /// <summary>Gets the column span.</summary>
        /// <value>The column span.</value>
        public int ColumnSpan => Columns.ColumnSpan;

        /// <summary>Gets the row span.</summary>
        /// <value>The row span.</value>
        public int RowSpan => DataSource.Count() + Columns.RowSpan;

        /// <summary>Gets the cells.</summary>
        /// <value>The cells.</value>
        public IEnumerable<Cell> Cells => GetHearderCells(Columns).Union(GetRecordCells());

        private IEnumerable<Cell> GetHearderCells(DataColumnCollection<T> cols) {
            List<Cell> cells = new();
            foreach (DataColumnBase<T> col in cols) {
                cells.Add(new Cell() {
                    Value = col.HeaderText,
                    CellStyle = col.HeaderStyle,
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
                RecordContext<T> recordContext = new(Columns.RowSpan, i, record);

                foreach (DataColumnBase<T> col in Columns.DataSourceColumns) {
                    yield return new Cell {
                        Value = col.GetFieldValue(recordContext),
                        CellStyle = col.GetFieldStyle(recordContext),
                        Point = p,
                        Size = new Size(col.ColumnSpan, 1),
                        Formula = col.GetFieldFormula(recordContext)
                    };
                    p += new Size(col.ColumnSpan, 0);
                }

                p = new Point(0, p.Y + 1);
                i++;
            }
        }

        /// <inheritdoc/>
        public IReadOnlyDictionary<int, double> RowHeights {
            get {
                Dictionary<int, double> dic = new();
                int i;

                for (i = 0; i < Columns.RowSpan; i++) {
                    dic.Add(i, HeaderHeight);
                }

                foreach (T item in DataSource) {
                    dic.Add(i++, ItemHeight);
                }

                return dic;
            }
        }

        /// <inheritdoc/>
        public TemplateContext GetContext() {
            return new TemplateContext(Cells, RowSpan, RowHeights);
        }
    }
}
