using System;
using System.Collections.Generic;
using System.Drawing;
using CloudyWing.SpreadsheetExporter.Util;

namespace CloudyWing.SpreadsheetExporter.Templates.List {
    /// <summary>
    /// Excel匯出資料各標題欄位設定
    /// </summary>
    public class DataColumn<T> {
        private DataColumnCollection<T> childColumns;
        private Point point;

        /// <summary>
        /// 顯示文字
        /// </summary>
        public string HeaderText { get; set; }

        /// <summary>
        /// 對應DataSource的Property Name
        /// </summary>
        public string DataKey { get; set; }

        /// <summary>
        /// 標題的儲存格樣式
        /// </summary>
        public CellStyle HeaderStyle { get; set; }

        /// <summary>
        /// 資料儲存格樣式，若有設定ItemStyleFunctor，則無效果
        /// </summary>
        public CellStyle ItemStyle { get; set; }

        /// <summary>
        /// 資料儲存格樣式，優先權高於ItemStyle
        /// </summary>
        public Func<object, T, CellStyle> ItemStyleFunctor { get; set; }

        /// <summary>
        /// 修正顯示值的委派
        /// </summary>
        public Func<object, T, object> ContentRender { get; set; }

        /// <summary>
        /// 座標
        /// </summary>
        public Point Point {
            get {
                return point;
            }
            internal set {
                point = value;
                ChildColumns.ResetColumnsPoint(value + new Size(0, RowSpan));
            }
        }

        /// <summary>
        /// 產生Excel公式，委派參數為從0開始計算的資料列索引
        /// </summary>
        public Func<int, string> ItemFormulaFunctor { get; set; }

        /// <summary>
        /// 欄跨度
        /// </summary>
        public int ColumnSpan => ChildColumns.Count == 0 ? 1 : ChildColumns.ColumnSpan;

        /// <summary>
        /// 列跨度
        /// </summary>
        public int RowSpan => ParentColumns.RowSpan - ChildColumns.RowSpan;

        /// <remarks>
        /// 自己和子層共幾層Column，用來計算RowSpan
        /// </remarks>
        internal int ColumnLayers => ChildColumns.Count == 0 ? 1 : ChildColumns.RowSpan + 1;

        /// <summary>
        /// 子標題欄位設定的集合
        /// </summary>
        public DataColumnCollection<T> ChildColumns {
            get {
                if (childColumns is null) {
                    childColumns = new DataColumnCollection<T>(this);
                }
                return childColumns;
            }
        }

        /// <summary>
        /// 自身標題欄位所屬的父集合
        /// </summary>
        public DataColumnCollection<T> ParentColumns { get; internal set; }

        internal virtual object GetContentValue(T valueObject) {
            IDictionary<string, object> data = DictionaryUtils.ConvertFrom(valueObject);

            if (string.IsNullOrWhiteSpace(DataKey)) {
                return ContentRender is null ? "" : ContentRender(null, valueObject);
            }
            object value = data[DataKey];

            return ContentRender is null ? value : ContentRender(value, valueObject);
        }

        internal virtual CellStyle GetDataCellStyle(T valueObject) {
            IDictionary<string, object> data = DictionaryUtils.ConvertFrom(valueObject);
            object value = string.IsNullOrWhiteSpace(DataKey) ? null : data[DataKey];

            return ItemStyleFunctor is null ? ItemStyle : ItemStyleFunctor(value, valueObject);
        }
    }

    public class DataColumn<TProperty, TObject> : DataColumn<TObject> {
        /// <summary>
        /// 資料儲存格樣式
        /// </summary>
        public new Func<TProperty, TObject, CellStyle> ItemStyleFunctor { get; set; }

        /// <summary>
        /// 修正顯示值的委派
        /// </summary>
        public new Func<TProperty, TObject, object> ContentRender { get; set; }

        internal override object GetContentValue(TObject valueObject) {
            IDictionary<string, object> data = DictionaryUtils.ConvertFrom(valueObject);

            if (string.IsNullOrWhiteSpace(DataKey)) {
                return ContentRender is null ? "" : ContentRender(default, valueObject);
            }
            object value = data[DataKey];

            if (ContentRender is null) {
                return value;
            }

            return ChangeValueTypeForFunc(ContentRender, value, valueObject);
        }

        internal override CellStyle GetDataCellStyle(TObject valueObject) {
            if (ItemStyleFunctor is null) {
                return ItemStyle;
            }

            if (string.IsNullOrWhiteSpace(DataKey)) {
                return ItemStyleFunctor(default, valueObject);
            }

            IDictionary<string, object> data = DictionaryUtils.ConvertFrom(valueObject);
            object value = data[DataKey];

            return ChangeValueTypeForFunc(ItemStyleFunctor, value, valueObject);
        }

        private TResult ChangeValueTypeForFunc<TResult>(Func<TProperty, TObject, TResult> func, object value, TObject valueObject) {
            if (value is null) {
                return func(default, valueObject);
            }

            Type fromType = Nullable.GetUnderlyingType(value.GetType()) ?? value.GetType();
            Type toType = Nullable.GetUnderlyingType(typeof(TProperty)) ?? typeof(TProperty);

            if (toType.IsPrimitive && typeof(IConvertible).IsAssignableFrom(fromType)) {
                return func((TProperty)Convert.ChangeType(value, toType), valueObject);
            }

            return func((TProperty)value, valueObject);
        }
    }
}
