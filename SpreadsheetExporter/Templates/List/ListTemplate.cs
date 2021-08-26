using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using CloudyWing.SpreadsheetExporter.Util;

namespace CloudyWing.SpreadsheetExporter.Templates.List {
    /// <summary>
    /// 利用設定DataSource和DataColumn來產出資料清單式樣板
    /// </summary>
    /// <typeparam name="T">每筆資料的型別</typeparam>
    public class ListTemplate<T> : ITemplate {
        /// <summary>
        /// Gets or sets the data source.
        /// </summary>
        /// <value>
        /// The data source.
        /// </value>
        public IEnumerable<T> DataSource { get; set; }

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
        public double HeaderHeight { get; set; } = 16.5d;

        /// <summary>
        /// Gets or sets the height of the item.
        /// </summary>
        /// <value>
        /// The height of the item.
        /// </value>
        public double ItemHeight { get; set; } = 16.5d;

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

        /// <summary>
        /// Gets the cells.
        /// </summary>
        /// <value>
        /// The cells.
        /// </value>
        public IEnumerable<Cell> Cells => GetHearderCells(Columns).Union(GetItemCells());

        private IEnumerable<Cell> GetHearderCells(DataColumnCollection<T> cols) {
            List<Cell> cells = new List<Cell>();
            foreach (DataColumn<T> col in cols) {
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

        private IEnumerable<Cell> GetItemCells() {
            Point p = new Point(0, Columns.RowSpan);

            int i = 0;
            foreach (T valueObj in DataSource) {
                foreach (DataColumn<T> col in Columns.DataSourceColumns) {
                    yield return new Cell {
                        Value = col.GetContentValue(valueObj),
                        CellStyle = col.GetDataCellStyle(valueObj),
                        Point = p,
                        Size = new Size(col.ColumnSpan, 1),
                        Formula = col.ItemFormulaFunctor is null ? null : col.ItemFormulaFunctor(i)
                    };
                    p += new Size(col.ColumnSpan, 0);
                }

                p = new Point(0, p.Y + 1);
                i++;
            }
        }

        /// <summary>
        /// Gets the row heights.
        /// </summary>
        /// <value>
        /// The row heights.
        /// </value>
        public IReadOnlyDictionary<int, double> RowHeights {
            get {
                Dictionary<int, double> dic = new Dictionary<int, double>();
                int i;

                for (i = 0; i < Columns.RowSpan; i++) {
                    dic.Add(i, HeaderHeight);
                }

                foreach (var item in DataSource) {
                    dic.Add(i++, ItemHeight);
                }

                return dic;
            }
        }

        /// <summary>
        /// Gets the context.
        /// </summary>
        /// <returns></returns>
        public TemplateContext GetContext() {
            Validate();
            return new TemplateContext(Cells, ColumnSpan, RowSpan, RowHeights);
        }

        private void Validate() {
            if (Columns.Count > 0) {
                if (DataSource is null) {
                    throw new NullReferenceException("未指定資料來源。");
                }

                if (DataSource.Count() == 0) {
                    return;
                }

                IDictionary<string, object> firstData = DictionaryUtils.ConvertFrom(DataSource.First());

                foreach (DataColumn<T> col in Columns.DataSourceColumns) {
                    if (!string.IsNullOrWhiteSpace(col.DataKey) && !firstData.ContainsKey(col.DataKey)) {
                        throw new ArgumentException($"資料來源未包含Property「{col.DataKey}」。");
                    }
                }
            }
        }
    }
}
