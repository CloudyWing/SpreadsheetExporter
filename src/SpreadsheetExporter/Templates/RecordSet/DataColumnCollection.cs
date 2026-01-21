using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Drawing;
using System.Linq;
using System.Linq.Expressions;

namespace CloudyWing.SpreadsheetExporter.Templates.RecordSet;

/// <summary>
/// The data column collection.
/// </summary>
/// <typeparam name="T">The type of the record.</typeparam>
public class DataColumnCollection<T> : Collection<DataColumnBase<T>> {
    private readonly DataColumnBase<T> parentItem;

    internal DataColumnCollection(DataColumnBase<T> cell) {
        parentItem = cell;
    }

    /// <summary>
    /// Gets the column span.
    /// </summary>
    /// <value>
    /// The column span.
    /// </value>
    public int ColumnSpan => Count == 0 ? 0 : this.Sum(x => x.ColumnSpan);

    /// <summary>
    /// Gets the row span.
    /// </summary>
    /// <value>
    /// The row span.
    /// </value>
    public int RowSpan => Count == 0 ? 0 : this.Max(x => x.ColumnLayers);

    /// <summary>
    /// Gets the root columns.
    /// </summary>
    /// <value>
    /// The root columns.
    /// </value>
    public DataColumnCollection<T> RootColumns {
        get {
            DataColumnCollection<T> items = this;
            while (items.parentItem?.ParentColumns is not null) {
                items = items.parentItem.ParentColumns;
            }
            return items;
        }
    }

    /// <summary>
    /// Calculates the points.
    /// </summary>
    internal void CalculatePoints(Point point) {
        Size offset = new();

        foreach (DataColumnBase<T> item in this) {
            item.ChildColumns?
                .CalculatePoints(point + new Size(offset.Width, item.RowSpan));
            item.Point = point + offset;
            offset.Width += item.ColumnSpan;
        }
    }

    /// <summary>
    /// Adds the specified header text.
    /// </summary>
    /// <param name="dataColumn">The data column.</param>
    /// <returns>The self.</returns>
    public new DataColumnCollection<T> Add(DataColumnBase<T> dataColumn) {
        base.Add(dataColumn);
        return this;
    }

    /// <summary>
    /// Adds the specified header text.
    /// </summary>
    /// <param name="headerText">The header text.</param>
    /// <param name="headerStyle">The header style.</param>
    /// <param name="fieldStyleGenerator">The field style generator.</param>
    /// <returns>The self.</returns>
    public DataColumnCollection<T> Add(
        string headerText, CellStyle? headerStyle = null,
        Func<RecordContext<T>, CellStyle> fieldStyleGenerator = null
    ) {
        RecordDataColumn<T> column = new();
        ApplyDefaultStyles(column, headerText, headerStyle, fieldStyleGenerator);
        Add(column);
        return this;
    }

    /// <summary>
    /// Adds the data column to the end of the DataColumnCollection&lt;T&gt;.
    /// </summary>
    /// <param name="headerText">The header text.</param>
    /// <param name="providerSetter">The provider setter.</param>
    /// <param name="headerStyle">The header style. The dafault is <c>SpreadsheetManager.DefaultCellStyles.HeaderStyle</c>.</param>
    /// <param name="fieldStyleGenerator">The field style generator. The dafault is <c>(context) =&gt; SpreadsheetManager.DefaultCellStyles.FieldStyle</c>.</param>
    /// <returns>The self.</returns>
    public DataColumnCollection<T> Add(
        string headerText, Action<GeneratorProvider<T, RecordContext<T>>> providerSetter,
        CellStyle? headerStyle = null, Func<RecordContext<T>, CellStyle> fieldStyleGenerator = null
    ) {
        GeneratorProvider<T, RecordContext<T>> provider = new();
        providerSetter?.Invoke(provider);

        RecordDataColumn<T> dataColumn = new();
        ApplyDefaultStyles(dataColumn, headerText, headerStyle, fieldStyleGenerator);
        provider.SetGeneratorForColumn(dataColumn);
        Add(dataColumn);
        return this;
    }

    /// <summary>
    /// Adds the data column to the end of the DataColumnCollection&lt;T&gt;.
    /// </summary>
    /// <typeparam name="TField">The type of the field.</typeparam>
    /// <param name="headerText">The header text.</param>
    /// <param name="fieldKey">The field key.</param>
    /// <param name="headerStyle">The header style. The dafault is <c>SpreadsheetManager.DefaultCellStyles.HeaderStyle</c>.</param>
    /// <param name="fieldStyleGenerator">The field style generator. The dafault is <c>(context) =&gt; SpreadsheetManager.DefaultCellStyles.FieldStyle</c>.</param>
    /// <returns>The self.</returns>
    public DataColumnCollection<T> Add<TField>(
        string headerText, string fieldKey,
        CellStyle? headerStyle = null, Func<FieldContext<T, TField>, CellStyle> fieldStyleGenerator = null
    ) {
        DataColumn<T, TField> column = new(fieldKey);
        ApplyDefaultStyles(column, headerText, headerStyle, fieldStyleGenerator);
        Add(column);
        return this;
    }

    /// <summary>
    /// Adds the data column to the end of the DataColumnCollection&lt;T&gt;.
    /// </summary>
    /// <typeparam name="TField">The type of the field.</typeparam>
    /// <param name="headerText">The header text.</param>
    /// <param name="fieldKeyExpression">The field key expression.</param>
    /// <param name="headerStyle">The header style. The dafault is <c>SpreadsheetManager.DefaultCellStyles.HeaderStyle</c>.</param>
    /// <param name="fieldStyleGenerator">The field style generator. The dafault is <c>(context) =&gt; SpreadsheetManager.DefaultCellStyles.FieldStyle</c>.</param>
    /// <returns>The self.</returns>
    public DataColumnCollection<T> Add<TField>(
        string headerText, Expression<Func<T, TField>> fieldKeyExpression,
        CellStyle? headerStyle = null, Func<FieldContext<T, TField>, CellStyle> fieldStyleGenerator = null
    ) {
        DataColumn<T, TField> column = new(fieldKeyExpression);
        ApplyDefaultStyles(column, headerText, headerStyle, fieldStyleGenerator);
        Add(column);
        return this;
    }

    /// <summary>
    /// Adds the data column to the end of the DataColumnCollection&lt;T&gt;.
    /// </summary>
    /// <typeparam name="TField">The type of the field.</typeparam>
    /// <param name="headerText">The header text.</param>
    /// <param name="fieldKey">The field key.</param>
    /// <param name="providerSetter">The provider setter.</param>
    /// <param name="headerStyle">The header style. The dafault is <c>SpreadsheetManager.DefaultCellStyles.HeaderStyle</c>.</param>
    /// <param name="fieldStyleGenerator">The field style generator. The dafault is <c>(context) =&gt; SpreadsheetManager.DefaultCellStyles.FieldStyle</c>.</param>
    /// <returns>The self.</returns>
    public DataColumnCollection<T> Add<TField>(
        string headerText, string fieldKey, Action<GeneratorProvider<T, FieldContext<T, TField>>> providerSetter,
        CellStyle? headerStyle = null, Func<FieldContext<T, TField>, CellStyle> fieldStyleGenerator = null
    ) {
        GeneratorProvider<T, FieldContext<T, TField>> provider = new();
        providerSetter?.Invoke(provider);

        DataColumn<T, TField> dataColumn = new(fieldKey);
        ApplyDefaultStyles(dataColumn, headerText, headerStyle, fieldStyleGenerator);
        provider.SetGeneratorForColumn(dataColumn);
        Add(dataColumn);
        return this;
    }

    /// <summary>
    /// Adds the data column to the end of the DataColumnCollection&lt;T&gt;.
    /// </summary>
    /// <typeparam name="TField">The type of the field.</typeparam>
    /// <param name="headerText">The header text.</param>
    /// <param name="fieldKeyExpression">The field key expression.</param>
    /// <param name="providerSetter">The provider setter.</param>
    /// <param name="headerStyle">The header style. The dafault is <c>SpreadsheetManager.DefaultCellStyles.HeaderStyle</c>.</param>
    /// <param name="fieldStyleGenerator">The field style generator. The dafault is <c>(context) =&gt; SpreadsheetManager.DefaultCellStyles.FieldStyle</c>.</param>
    /// <returns>The self.</returns>
    public DataColumnCollection<T> Add<TField>(
        string headerText, Expression<Func<T, TField>> fieldKeyExpression,
        Action<GeneratorProvider<T, FieldContext<T, TField>>> providerSetter,
        CellStyle? headerStyle = null, Func<FieldContext<T, TField>, CellStyle> fieldStyleGenerator = null
    ) {
        GeneratorProvider<T, FieldContext<T, TField>> provider = new();
        providerSetter?.Invoke(provider);

        DataColumn<T, TField> dataColumn = new(fieldKeyExpression);
        ApplyDefaultStyles(dataColumn, headerText, headerStyle, fieldStyleGenerator);
        provider.SetGeneratorForColumn(dataColumn);
        Add(dataColumn);
        return this;
    }

    /// <summary>
    /// Adds the child data column at the end of the last data column.
    /// </summary>
    /// <param name="childColumn">The child data column.</param>
    /// <returns>The self.</returns>
    /// <exception cref="InvalidOperationException">No columns available to add child to.</exception>
    public DataColumnCollection<T> AddChildToLast(DataColumnBase<T> childColumn) {
        if (Count == 0) {
            throw new InvalidOperationException("No columns available to add child to.");
        }

        DataColumnBase<T> column = this.Last();
        column.ChildColumns.Add(childColumn);

        return this;
    }

    /// <summary>
    /// Adds the child data column at the end of the last data column.
    /// </summary>
    /// <param name="headerText">The header text.</param>
    /// <param name="headerStyle">The header style.</param>
    /// <param name="fieldStyleGenerator">The field style generator.</param>
    /// <returns>The self.</returns>
    public DataColumnCollection<T> AddChildToLast(
        string headerText, CellStyle? headerStyle = null,
        Func<RecordContext<T>, CellStyle> fieldStyleGenerator = null
    ) {
        RecordDataColumn<T> column = new();
        ApplyDefaultStyles(column, headerText, headerStyle, fieldStyleGenerator);
        return AddChildToLast(column);
    }

    /// <summary>
    /// Adds the child data column at the end of the last data column.
    /// </summary>
    /// <param name="headerText">The header text.</param>
    /// <param name="providerSetter">The provider setter.</param>
    /// <param name="headerStyle">The header style. The dafault is <c>SpreadsheetManager.DefaultCellStyles.HeaderStyle</c>.</param>
    /// <param name="fieldStyleGenerator">The field style generator. The dafault is <c>(context) =&gt; SpreadsheetManager.DefaultCellStyles.FieldStyle</c>.</param>
    /// <returns>The self.</returns>
    public DataColumnCollection<T> AddChildToLast(
        string headerText, Action<GeneratorProvider<T, RecordContext<T>>> providerSetter,
        CellStyle? headerStyle = null, Func<RecordContext<T>, CellStyle> fieldStyleGenerator = null
    ) {
        GeneratorProvider<T, RecordContext<T>> provider = new();
        providerSetter?.Invoke(provider);

        RecordDataColumn<T> dataColumn = new();
        ApplyDefaultStyles(dataColumn, headerText, headerStyle, fieldStyleGenerator);
        provider.SetGeneratorForColumn(dataColumn);
        return AddChildToLast(dataColumn);
    }

    /// <summary>
    /// Adds the child data column at the end of the last data column.
    /// </summary>
    /// <typeparam name="TField">The type of the field.</typeparam>
    /// <param name="headerText">The header text.</param>
    /// <param name="fieldKey">The field key.</param>
    /// <param name="headerStyle">The header style. The dafault is <c>SpreadsheetManager.DefaultCellStyles.HeaderStyle</c>.</param>
    /// <param name="fieldStyleGenerator">The field style generator. The dafault is <c>(context) =&gt; SpreadsheetManager.DefaultCellStyles.FieldStyle</c>.</param>
    /// <returns>The self.</returns>
    public DataColumnCollection<T> AddChildToLast<TField>(
        string headerText, string fieldKey,
        CellStyle? headerStyle = null, Func<FieldContext<T, TField>, CellStyle> fieldStyleGenerator = null
    ) {
        DataColumn<T, TField> column = new(fieldKey);
        ApplyDefaultStyles(column, headerText, headerStyle, fieldStyleGenerator);
        return AddChildToLast(column);
    }

    /// <summary>
    /// Adds the child data column at the end of the last data column.
    /// </summary>
    /// <typeparam name="TField">The type of the field.</typeparam>
    /// <param name="headerText">The header text.</param>
    /// <param name="fieldKeyExpression">The field key expression.</param>
    /// <param name="headerStyle">The header style. The dafault is <c>SpreadsheetManager.DefaultCellStyles.HeaderStyle</c>.</param>
    /// <param name="fieldStyleGenerator">The field style generator. The dafault is <c>(context) =&gt; SpreadsheetManager.DefaultCellStyles.FieldStyle</c>.</param>
    /// <returns>The self.</returns>
    public DataColumnCollection<T> AddChildToLast<TField>(
        string headerText, Expression<Func<T, TField>> fieldKeyExpression,
        CellStyle? headerStyle = null, Func<FieldContext<T, TField>, CellStyle> fieldStyleGenerator = null
    ) {
        DataColumn<T, TField> column = new(fieldKeyExpression);
        ApplyDefaultStyles(column, headerText, headerStyle, fieldStyleGenerator);
        return AddChildToLast(column);
    }

    /// <summary>
    /// Adds the child data column at the end of the last data column.
    /// </summary>
    /// <typeparam name="TField">The type of the field.</typeparam>
    /// <param name="headerText">The header text.</param>
    /// <param name="fieldKey">The field key.</param>
    /// <param name="providerSetter">The provider setter.</param>
    /// <param name="headerStyle">The header style. The dafault is <c>SpreadsheetManager.DefaultCellStyles.HeaderStyle</c>.</param>
    /// <param name="fieldStyleGenerator">The field style generator. The dafault is <c>(context) =&gt; SpreadsheetManager.DefaultCellStyles.FieldStyle</c>.</param>
    /// <returns>The self.</returns>
    public DataColumnCollection<T> AddChildToLast<TField>(
        string headerText, string fieldKey, Action<GeneratorProvider<T, FieldContext<T, TField>>> providerSetter,
        CellStyle? headerStyle = null, Func<FieldContext<T, TField>, CellStyle> fieldStyleGenerator = null
    ) {
        GeneratorProvider<T, FieldContext<T, TField>> provider = new();
        providerSetter?.Invoke(provider);

        DataColumn<T, TField> dataColumn = new(fieldKey);
        ApplyDefaultStyles(dataColumn, headerText, headerStyle, fieldStyleGenerator);
        provider.SetGeneratorForColumn(dataColumn);
        return AddChildToLast(dataColumn);
    }

    /// <summary>
    /// Adds the child data column at the end of the last data column.
    /// </summary>
    /// <typeparam name="TField">The type of the field.</typeparam>
    /// <param name="headerText">The header text.</param>
    /// <param name="fieldKeyExpression">The field key expression.</param>
    /// <param name="providerSetter">The provider setter.</param>
    /// <param name="headerStyle">The header style. The dafault is <c>SpreadsheetManager.DefaultCellStyles.HeaderStyle</c>.</param>
    /// <param name="fieldStyleGenerator">The field style generator. The dafault is <c>(context) =&gt; SpreadsheetManager.DefaultCellStyles.FieldStyle</c>.</param>
    /// <returns>The self.</returns>
    public DataColumnCollection<T> AddChildToLast<TField>(
        string headerText, Expression<Func<T, TField>> fieldKeyExpression,
        Action<GeneratorProvider<T, FieldContext<T, TField>>> providerSetter,
        CellStyle? headerStyle = null, Func<FieldContext<T, TField>, CellStyle> fieldStyleGenerator = null
    ) {
        GeneratorProvider<T, FieldContext<T, TField>> provider = new();
        providerSetter?.Invoke(provider);

        DataColumn<T, TField> dataColumn = new(fieldKeyExpression);
        ApplyDefaultStyles(dataColumn, headerText, headerStyle, fieldStyleGenerator);
        provider.SetGeneratorForColumn(dataColumn);
        return AddChildToLast(dataColumn);
    }

    private static void ApplyDefaultStyles(
        RecordDataColumn<T> column, string headerText,
        CellStyle? headerStyle, Func<RecordContext<T>, CellStyle> fieldStyleGenerator
    ) {
        column.HeaderText = headerText;
        column.HeaderStyle = headerStyle ?? SpreadsheetManager.DefaultCellStyles.HeaderStyle;
        column.FieldStyleGenerator = fieldStyleGenerator
            ?? ((context) => SpreadsheetManager.DefaultCellStyles.FieldStyle);
    }

    private static void ApplyDefaultStyles<TField>(
        DataColumn<T, TField> column, string headerText,
        CellStyle? headerStyle, Func<FieldContext<T, TField>, CellStyle> fieldStyleGenerator
    ) {
        column.HeaderText = headerText;
        column.HeaderStyle = headerStyle ?? SpreadsheetManager.DefaultCellStyles.HeaderStyle;
        column.FieldStyleGenerator = fieldStyleGenerator
            ?? ((context) => SpreadsheetManager.DefaultCellStyles.FieldStyle);
    }

    /// <inheritdoc/>
    protected override void InsertItem(int index, DataColumnBase<T> item) {
        if (item.ParentColumns is not null) {
            throw new ArgumentException(
                $"DataColumnBase<{typeof(T).Name}> with HeaderText '{item.HeaderText}' is already contained by another DataColumnCollection. " +
                $"A column can only belong to one collection at a time.",
                nameof(item)
            );
        }

        base.InsertItem(index, item);
        item.ParentColumns = this;
    }

    /// <inheritdoc/>
    protected override void RemoveItem(int index) {
        base.RemoveItem(index);
        Items[index].ParentColumns = null;
    }

    /// <inheritdoc/>
    protected override void SetItem(int index, DataColumnBase<T> item) {
        if (item.ParentColumns is not null) {
            throw new ArgumentException(
                $"DataColumnBase<{typeof(T).Name}> with HeaderText '{item.HeaderText}' is already contained by another DataColumnCollection. " +
                $"A column can only belong to one collection at a time.",
                nameof(item)
            );
        }

        base.SetItem(index, item);
        Items[index].ParentColumns = null;
        item.ParentColumns = this;
    }

    /// <inheritdoc/>
    protected override void ClearItems() {
        for (int i = 0; i < Items.Count; i++) {
            Items[i].ParentColumns = null;
        }
        base.ClearItems();
    }

    /// <summary>
    /// Get the column containing the properties of the data source.
    /// </summary>
    /// <value>
    /// The data columns.
    /// </value>
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

    /// <summary>
    /// The generator provider.
    /// </summary>
    /// <typeparam name="TRecord">The type of the record.</typeparam>
    /// <typeparam name="TContext">The type of the context.</typeparam>
    /// <seealso cref="Collection{DataColumnBase}">Collection{DataColumnBase}</seealso>
    public sealed class GeneratorProvider<TRecord, TContext> where TContext : RecordContext<TRecord> {
        private ProviderTypes types = ProviderTypes.None;
        private Func<TContext, object> valueGenerator;
        private Func<TContext, string> formulaGenerator;
        private Func<TContext, DataValidation> dataValidationGenerator;

        /// <summary>
        /// Uses the value.
        /// </summary>
        /// <param name="generator">The generator.</param>
        public void UseValue(Func<TContext, object> generator) {
            types |= ProviderTypes.Value;
            valueGenerator = generator;
        }

        /// <summary>
        /// Uses the formula.
        /// </summary>
        /// <param name="generator">The generator.</param>
        public void UseFormula(Func<TContext, string> generator) {
            types |= ProviderTypes.Formula;
            formulaGenerator = generator;
        }

        /// <summary>
        /// Uses the data validation.
        /// </summary>
        /// <param name="dataValidationGenerator">The data validation generator.</param>
        public void UseDataValidation(Func<TContext, DataValidation> dataValidationGenerator) {
            types |= ProviderTypes.DataValidation;
            this.dataValidationGenerator = dataValidationGenerator;
        }

        internal void SetGeneratorForColumn(RecordDataColumn<TRecord> dataColumn) {
            ValidateProviderTypes();

            if (types.HasFlag(ProviderTypes.Value)) {
                dataColumn.FieldValueGenerator = ConvertGenerator<Func<RecordContext<TRecord>, object>>(valueGenerator);
            }

            if (types.HasFlag(ProviderTypes.Formula)) {
                dataColumn.FieldFormulaGenerator = ConvertGenerator<Func<RecordContext<TRecord>, string>>(formulaGenerator);
            }

            if (types.HasFlag(ProviderTypes.DataValidation)) {
                dataColumn.FieldDataValidationGenerator = ConvertGenerator<Func<RecordContext<TRecord>, DataValidation>>(dataValidationGenerator);
            }
        }

        internal void SetGeneratorForColumn<TField>(DataColumn<TRecord, TField> dataColumn) {
            ValidateProviderTypes();

            if (types.HasFlag(ProviderTypes.Value)) {
                dataColumn.FieldValueGenerator = ConvertGenerator<Func<FieldContext<TRecord, TField>, object>>(valueGenerator);
            }

            if (types.HasFlag(ProviderTypes.Formula)) {
                dataColumn.FieldFormulaGenerator = ConvertGenerator<Func<FieldContext<TRecord, TField>, string>>(formulaGenerator);
            }

            if (types.HasFlag(ProviderTypes.DataValidation)) {
                dataColumn.FieldDataValidationGenerator = ConvertGenerator<Func<FieldContext<TRecord, TField>, DataValidation>>(dataValidationGenerator);
            }
        }

        private void ValidateProviderTypes() {
            if (types == ProviderTypes.None) {
                throw new ArgumentException("No provider type specified. Use UseValue, UseFormula, or UseDataValidation.");
            }

            if (types.HasFlag(ProviderTypes.Value) && types.HasFlag(ProviderTypes.Formula)) {
                throw new InvalidOperationException(
                    "Cannot use both UseValue and UseFormula on the same column. "
                    + "A cell can only have a value or a formula, not both."
                );
            }
        }

        private static TField ConvertGenerator<TField>(object generator) where TField : class =>
            generator as TField ?? throw new InvalidCastException("Generator type mismatch.");

        [Flags]
        private enum ProviderTypes {
            None = 0,
            Value = 1,
            Formula = 2,
            DataValidation = 4
        }
    }
}
