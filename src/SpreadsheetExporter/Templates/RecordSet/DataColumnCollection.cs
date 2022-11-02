using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Drawing;
using System.Linq;
using System.Linq.Expressions;

namespace CloudyWing.SpreadsheetExporter.Templates.RecordSet {
    /// <summary>The data column collection.</summary>
    /// <typeparam name="T">The type of the record.</typeparam>
    public class DataColumnCollection<T> : Collection<DataColumnBase<T>> {
        private readonly DataColumnBase<T> parentItem;

        internal DataColumnCollection(DataColumnBase<T> cell) {
            parentItem = cell;
        }

        /// <summary>Gets the column span.</summary>
        /// <value>The column span.</value>
        public int ColumnSpan => Count == 0 ? 0 : this.Sum(x => x.ColumnSpan);

        /// <summary>Gets the row span.</summary>
        /// <value>The row span.</value>
        public int RowSpan => Count == 0 ? 0 : this.Max(x => x.ColumnLayers);

        /// <summary>Gets the root columns.</summary>
        /// <value>The root columns.</value>
        public DataColumnCollection<T> RootColumns {
            get {
                DataColumnCollection<T> items = this;
                while (items.parentItem != null && items.parentItem.ParentColumns != null) {
                    items = items.parentItem.ParentColumns;
                }
                return items;
            }
        }

        /// <summary>
        /// 從根節點往下重設座標
        /// </summary>
        internal void ResetRootPoint() {
            RootColumns.ResetColumnsPoint(Point.Empty);
        }

        /// <summary>
        /// 重設底下所有DataColumn的座標
        /// </summary>
        /// <param name="point"></param>
        internal void ResetColumnsPoint(Point point) {
            Size offset = new();

            foreach (DataColumnBase<T> item in this) {
                item.Point = point + offset;
                offset.Width += item.ColumnSpan;
            }
        }

        /// <summary>Adds the data column to the end of the DataColumnCollection&lt;T&gt;.</summary>
        /// <param name="headerText">The header text.</param>
        /// <param name="providerSetter">The provider setter.</param>
        /// <param name="headerStyle">The header style. The dafault is SpreadsheetManager.Configuration.ListHeaderStyle.</param>
        /// <param name="fieldStyleGenerator">The field style generator. The dafault is <c>(context) => SpreadsheetManager.Configuration.ListHeaderStyle</c>.</param>
        public void Add(
            string headerText, Action<GeneratorProvider<RecordContext<T>, T>> providerSetter,
            CellStyle? headerStyle = null, Func<RecordContext<T>, CellStyle> fieldStyleGenerator = null
        ) {
            GeneratorProvider<RecordContext<T>, T> provider = new();
            providerSetter?.Invoke(provider);

            RecordDataColumn<T> dataColumn = new() {
                HeaderText = headerText,
                HeaderStyle = headerStyle ?? SpreadsheetManager.Configuration.ListHeaderStyle,
                FieldStyleGenerator = fieldStyleGenerator ?? ((context) => SpreadsheetManager.Configuration.ListTextStyle)
            };

            provider.SetGeneratorForColumn(dataColumn);

            Add(dataColumn);
        }

        /// <summary>Adds the data column to the end of the DataColumnCollection&lt;T&gt;.</summary>
        /// <typeparam name="TField">The type of the field.</typeparam>
        /// <param name="headerText">The header text.</param>
        /// <param name="fieldKey">The field key.</param>
        /// <param name="headerStyle">The header style. The dafault is SpreadsheetManager.Configuration.ListHeaderStyle.</param>
        /// <param name="fieldStyleGenerator">The field style generator. The dafault is <c>(context) => SpreadsheetManager.Configuration.ListHeaderStyle</c>.</param>
        public void Add<TField>(
            string headerText, string fieldKey,
            CellStyle? headerStyle = null, Func<FieldContext<TField, T>, CellStyle> fieldStyleGenerator = null
        ) {
            Add(new DataColumn<TField, T>(fieldKey) {
                HeaderText = headerText,
                HeaderStyle = headerStyle ?? SpreadsheetManager.Configuration.ListHeaderStyle,
                FieldStyleGenerator = fieldStyleGenerator ?? ((context) => SpreadsheetManager.Configuration.ListTextStyle)
            });
        }

        /// <summary>Adds the data column to the end of the DataColumnCollection&lt;T&gt;.</summary>
        /// <typeparam name="TField">The type of the field.</typeparam>
        /// <param name="headerText">The header text.</param>
        /// <param name="fieldKeyExpression">The field key expression.</param>
        /// <param name="headerStyle">The header style. The dafault is SpreadsheetManager.Configuration.ListHeaderStyle.</param>
        /// <param name="fieldStyleGenerator">The field style generator. The dafault is <c>(context) => SpreadsheetManager.Configuration.ListHeaderStyle</c>.</param>
        public void Add<TField>(
            string headerText, Expression<Func<T, TField>> fieldKeyExpression,
            CellStyle? headerStyle = null, Func<FieldContext<TField, T>, CellStyle> fieldStyleGenerator = null
        ) {
            Add(new DataColumn<TField, T>(fieldKeyExpression) {
                HeaderText = headerText,
                HeaderStyle = headerStyle ?? SpreadsheetManager.Configuration.ListHeaderStyle,
                FieldStyleGenerator = fieldStyleGenerator ?? ((context) => SpreadsheetManager.Configuration.ListTextStyle)
            });
        }

        /// <summary>Adds the data column to the end of the DataColumnCollection&lt;T&gt;.</summary>
        /// <typeparam name="TField">The type of the field.</typeparam>
        /// <param name="headerText">The header text.</param>
        /// <param name="fieldKey">The field key.</param>
        /// <param name="providerSetter">The provider setter.</param>
        /// <param name="headerStyle">The header style. The dafault is SpreadsheetManager.Configuration.ListHeaderStyle.</param>
        /// <param name="fieldStyleGenerator">The field style generator. The dafault is <c>(context) =&gt; SpreadsheetManager.Configuration.ListTextStyle</c>.</param>
        public void Add<TField>(
            string headerText, string fieldKey, Action<GeneratorProvider<FieldContext<TField, T>, T>> providerSetter,
            CellStyle? headerStyle = null, Func<FieldContext<TField, T>, CellStyle> fieldStyleGenerator = null
        ) {
            GeneratorProvider<FieldContext<TField, T>, T> provider = new();
            providerSetter?.Invoke(provider);

            DataColumn<TField, T> dataColumn = new(fieldKey) {
                HeaderText = headerText,
                HeaderStyle = headerStyle ?? SpreadsheetManager.Configuration.ListHeaderStyle,
                FieldStyleGenerator = fieldStyleGenerator ?? ((context) => SpreadsheetManager.Configuration.ListTextStyle)
            };

            provider.SetGeneratorForColumn(dataColumn);

            Add(dataColumn);
        }

        /// <summary>Adds the data column to the end of the DataColumnCollection&lt;T&gt;.</summary>
        /// <typeparam name="TField">The type of the field.</typeparam>
        /// <param name="headerText">The header text.</param>
        /// <param name="fieldKeyExpression">The field key expression.</param>
        /// <param name="providerSetter">The provider setter.</param>
        /// <param name="headerStyle">The header style. The dafault is SpreadsheetManager.Configuration.ListHeaderStyle.</param>
        /// <param name="fieldStyleGenerator">The field style generator. The dafault is <c>(context) =&gt; SpreadsheetManager.Configuration.ListTextStyle</c>.</param>
        public void Add<TField>(
            string headerText, Expression<Func<T, TField>> fieldKeyExpression,
            Action<GeneratorProvider<FieldContext<TField, T>, T>> providerSetter,
            CellStyle? headerStyle = null, Func<FieldContext<TField, T>, CellStyle> fieldStyleGenerator = null
        ) {
            GeneratorProvider<FieldContext<TField, T>, T> provider = new();
            providerSetter?.Invoke(provider);

            DataColumn<TField, T> dataColumn = new(fieldKeyExpression) {
                HeaderText = headerText,
                HeaderStyle = headerStyle ?? SpreadsheetManager.Configuration.ListHeaderStyle,
                FieldStyleGenerator = fieldStyleGenerator ?? ((context) => SpreadsheetManager.Configuration.ListTextStyle)
            };

            provider.SetGeneratorForColumn(dataColumn);

            Add(dataColumn);
        }

        /// <summary>Adds the child data column at the end of the last data column.</summary>
        /// <param name="childColumn">The child data column.</param>
        /// <exception cref="NullReferenceException"></exception>
        public void AddChildToLast(DataColumnBase<T> childColumn) {
            DataColumnBase<T> column = this.LastOrDefault();
            if (column is null) {
                throw new NullReferenceException($"No {nameof(DataColumnBase<T>)} have been created.");
            }

            column.ChildColumns.Add(childColumn);
        }

        /// <summary>Adds the child data column at the end of the last data column.</summary>
        /// <param name="headerText">The header text.</param>
        /// <param name="providerSetter">The provider setter.</param>
        /// <param name="headerStyle">The header style. The dafault is SpreadsheetManager.Configuration.ListHeaderStyle.</param>
        /// <param name="fieldStyleGenerator">The field style generator. The dafault is <c>(context) =&gt; SpreadsheetManager.Configuration.ListTextStyle</c>.</param>
        public void AddChildToLast(
            string headerText, Action<GeneratorProvider<RecordContext<T>, T>> providerSetter,
            CellStyle? headerStyle = null, Func<RecordContext<T>, CellStyle> fieldStyleGenerator = null
        ) {
            GeneratorProvider<RecordContext<T>, T> provider = new();
            providerSetter?.Invoke(provider);

            RecordDataColumn<T> dataColumn = new() {
                HeaderText = headerText,
                HeaderStyle = headerStyle ?? SpreadsheetManager.Configuration.ListHeaderStyle,
                FieldStyleGenerator = fieldStyleGenerator ?? ((context) => SpreadsheetManager.Configuration.ListTextStyle)
            };

            provider.SetGeneratorForColumn(dataColumn);

            AddChildToLast(dataColumn);
        }

        /// <summary>Adds the child data column at the end of the last data column.</summary>
        /// <typeparam name="TField">The type of the field.</typeparam>
        /// <param name="headerText">The header text.</param>
        /// <param name="fieldKey">The field key.</param>
        /// <param name="headerStyle">The header style. The dafault is SpreadsheetManager.Configuration.ListHeaderStyle.</param>
        /// <param name="fieldStyleGenerator">The field style generator. The dafault is <c>(context) => SpreadsheetManager.Configuration.ListHeaderStyle</c>.</param>
        public void AddChildToLast<TField>(
            string headerText, string fieldKey,
            CellStyle? headerStyle = null, Func<FieldContext<TField, T>, CellStyle> fieldStyleGenerator = null
        ) {
            AddChildToLast(new DataColumn<TField, T>(fieldKey) {
                HeaderText = headerText,
                HeaderStyle = headerStyle ?? SpreadsheetManager.Configuration.ListHeaderStyle,
                FieldStyleGenerator = fieldStyleGenerator ?? ((context) => SpreadsheetManager.Configuration.ListTextStyle)
            });
        }

        /// <summary>Adds the child data column at the end of the last data column.</summary>
        /// <typeparam name="TField">The type of the field.</typeparam>
        /// <param name="headerText">The header text.</param>
        /// <param name="fieldKeyExpression">The field key expression.</param>
        /// <param name="headerStyle">The header style. The dafault is SpreadsheetManager.Configuration.ListHeaderStyle.</param>
        /// <param name="fieldStyleGenerator">The field style generator. The dafault is <c>(context) => SpreadsheetManager.Configuration.ListHeaderStyle</c>.</param>
        public void AddChildToLast<TField>(
            string headerText, Expression<Func<T, TField>> fieldKeyExpression,
            CellStyle? headerStyle = null, Func<FieldContext<TField, T>, CellStyle> fieldStyleGenerator = null
        ) {
            AddChildToLast(new DataColumn<TField, T>(fieldKeyExpression) {
                HeaderText = headerText,
                HeaderStyle = headerStyle ?? SpreadsheetManager.Configuration.ListHeaderStyle,
                FieldStyleGenerator = fieldStyleGenerator ?? ((context) => SpreadsheetManager.Configuration.ListTextStyle)
            });
        }

        /// <summary>Adds the child data column at the end of the last data column.</summary>
        /// <typeparam name="TField">The type of the field.</typeparam>
        /// <param name="headerText">The header text.</param>
        /// <param name="fieldKey">The field key.</param>
        /// <param name="providerSetter">The provider setter.</param>
        /// <param name="headerStyle">The header style. The dafault is SpreadsheetManager.Configuration.ListHeaderStyle.</param>
        /// <param name="fieldStyleGenerator">The field style generator. The dafault is <c>(context) =&gt; SpreadsheetManager.Configuration.ListTextStyle</c>.</param>
        public void AddChildToLast<TField>(
            string headerText, string fieldKey, Action<GeneratorProvider<FieldContext<TField, T>, T>> providerSetter,
            CellStyle? headerStyle = null, Func<FieldContext<TField, T>, CellStyle> fieldStyleGenerator = null
        ) {
            GeneratorProvider<FieldContext<TField, T>, T> provider = new();
            providerSetter?.Invoke(provider);

            DataColumn<TField, T> dataColumn = new(fieldKey) {
                HeaderText = headerText,
                HeaderStyle = headerStyle ?? SpreadsheetManager.Configuration.ListHeaderStyle,
                FieldStyleGenerator = fieldStyleGenerator ?? ((context) => SpreadsheetManager.Configuration.ListTextStyle)
            };

            provider.SetGeneratorForColumn(dataColumn);

            AddChildToLast(dataColumn);
        }

        /// <summary>Adds the child data column at the end of the last data column.</summary>
        /// <typeparam name="TField">The type of the field.</typeparam>
        /// <param name="headerText">The header text.</param>
        /// <param name="fieldKeyExpression">The field key expression.</param>
        /// <param name="providerSetter">The provider setter.</param>
        /// <param name="headerStyle">The header style. The dafault is SpreadsheetManager.Configuration.ListHeaderStyle.</param>
        /// <param name="fieldStyleGenerator">The field style generator. The dafault is <c>(context) =&gt; SpreadsheetManager.Configuration.ListTextStyle</c>.</param>
        public void AddChildToLast<TField>(
            string headerText, Expression<Func<T, TField>> fieldKeyExpression,
            Action<GeneratorProvider<FieldContext<TField, T>, T>> providerSetter,
            CellStyle? headerStyle = null, Func<FieldContext<TField, T>, CellStyle> fieldStyleGenerator = null
        ) {
            GeneratorProvider<FieldContext<TField, T>, T> provider = new();
            providerSetter?.Invoke(provider);

            DataColumn<TField, T> dataColumn = new(fieldKeyExpression) {
                HeaderText = headerText,
                HeaderStyle = headerStyle ?? SpreadsheetManager.Configuration.ListHeaderStyle,
                FieldStyleGenerator = fieldStyleGenerator ?? ((context) => SpreadsheetManager.Configuration.ListTextStyle)
            };

            provider.SetGeneratorForColumn(dataColumn);

            AddChildToLast(dataColumn);
        }

        protected override void InsertItem(int index, DataColumnBase<T> item) {
            if (item.ParentColumns != null) {
                throw new ArgumentException($"{nameof(DataColumnBase<T>)} is already contained by another {nameof(DataColumnCollection<T>)}.", nameof(item));
            }

            base.InsertItem(index, item);
            item.ParentColumns = this;
            ResetRootPoint();
        }

        protected override void RemoveItem(int index) {
            Items[index].ParentColumns = null;
            base.RemoveItem(index);
            ResetRootPoint();
        }

        protected override void SetItem(int index, DataColumnBase<T> item) {
            if (item.ParentColumns != null) {
                throw new ArgumentException($"{nameof(DataColumnBase<T>)} is already contained by another {nameof(DataColumnCollection<T>)}.", nameof(item));
            }

            Items[index].ParentColumns = null;
            base.SetItem(index, item);
            item.ParentColumns = this;
            ResetRootPoint();
        }

        protected override void ClearItems() {
            foreach (DataColumnBase<T> item in Items) {
                item.ParentColumns = null;
            }
            base.ClearItems();
            ResetRootPoint();
        }

        /// <summary>Get the column containing the properties of the data source.</summary>
        /// <value>The data columns.</value>
        public IEnumerable<DataColumnBase<T>> DataSourceColumns {
            get {
                foreach (DataColumnBase<T> item in this) {
                    if (item.ChildColumns.Count == 0) {
                        yield return item;
                    } else {
                        foreach (DataColumnBase<T> dataHeader in item.ChildColumns.DataSourceColumns) {
                            yield return dataHeader;
                        }
                    }
                }
            }
        }

        public sealed class GeneratorProvider<TContext, TRecord> where TContext : RecordContext<TRecord> {
            private ProviderType type = ProviderType.None;
            private Func<TContext, object> valueGenerator;
            private Func<TContext, string> formulaGenerator;

            /// <summary>Uses the value.</summary>
            /// <param name="generator">The generator.</param>
            public void UseValue(Func<TContext, object> generator) {
                type = ProviderType.Value;
                valueGenerator = generator;
            }

            /// <summary>Uses the formula.</summary>
            /// <param name="generator">The generator.</param>
            public void UseFormula(Func<TContext, string> generator) {
                type = ProviderType.Formula;
                formulaGenerator = generator;
            }

            internal void SetGeneratorForColumn(RecordDataColumn<TRecord> dataColumn) {
                switch (type) {
                    case ProviderType.Value:
                        dataColumn.FieldValueGenerator = valueGenerator as Func<RecordContext<TRecord>, object>;
                        break;
                    case ProviderType.Formula:
                        dataColumn.FieldFormulaGenerator = formulaGenerator as Func<RecordContext<TRecord>, string>;
                        break;
                    default:
                        throw new ArgumentException();
                }
            }

            internal void SetGeneratorForColumn<TField>(DataColumn<TField, TRecord> dataColumn) {
                switch (type) {
                    case ProviderType.Value:
                        dataColumn.FieldValueGenerator = valueGenerator as Func<FieldContext<TField, TRecord>, object>;
                        break;
                    case ProviderType.Formula:
                        dataColumn.FieldFormulaGenerator = formulaGenerator as Func<FieldContext<TField, TRecord>, string>;
                        break;
                    default:
                        throw new ArgumentException();
                }
            }

            private enum ProviderType {
                None,
                Value,
                Formula
            }
        }
    }
}
