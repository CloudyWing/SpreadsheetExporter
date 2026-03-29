using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Drawing;
using System.Linq;
using System.Linq.Expressions;

namespace CloudyWing.SpreadsheetExporter.Templates.RecordSet;

/// <summary>
/// The record set column collection.
/// </summary>
/// <typeparam name="T">The type of the record.</typeparam>
public class RecordSetColumnCollection<T> : Collection<RecordSetColumnBase<T>> {
    private readonly RecordSetColumnBase<T>? parentItem;

    internal RecordSetColumnCollection(RecordSetColumnBase<T>? cell) {
        parentItem = cell;
    }

    /// <summary>
    /// Gets the column span.
    /// </summary>
    /// <value>The column span.</value>
    public int ColumnSpan => Count == 0 ? 0 : this.Sum(x => x.ColumnSpan);

    /// <summary>
    /// Gets the row span.
    /// </summary>
    /// <value>The row span.</value>
    public int RowSpan => Count == 0 ? 0 : this.Max(x => x.HeaderDepth);

    /// <summary>
    /// Gets the last column in this collection.
    /// </summary>
    /// <returns>The last <see cref="RecordSetColumnBase{T}"/> in the collection.</returns>
    /// <exception cref="InvalidOperationException">The collection is empty.</exception>
    public RecordSetColumnBase<T> GetLastColumn() {
        if (Count == 0) {
            throw new InvalidOperationException("No columns available.");
        }
        return Items[Count - 1];
    }

    /// <summary>
    /// Gets the parent collection that contains the column owning this child collection.
    /// </summary>
    /// <value>The parent <see cref="RecordSetColumnCollection{T}"/>.</value>
    /// <exception cref="InvalidOperationException">This is already the root collection.</exception>
    public RecordSetColumnCollection<T> Parent {
        get {
            if (parentItem?.ParentCollection is null) {
                throw new InvalidOperationException("Cannot navigate to parent: this is already the root collection.");
            }
            return parentItem.ParentCollection;
        }
    }

    /// <summary>
    /// Gets the root collection at the top of the hierarchy.
    /// </summary>
    /// <value>The root column collection.</value>
    public RecordSetColumnCollection<T> RootColumns {
        get {
            RecordSetColumnCollection<T> items = this;
            while (items.parentItem?.ParentCollection is not null) {
                items = items.parentItem.ParentCollection;
            }
            return items;
        }
    }

    /// <summary>
    /// Calculates the points for all columns in the collection.
    /// </summary>
    internal void CalculatePoints(Point point) {
        Size offset = new();

        foreach (RecordSetColumnBase<T> item in this) {
            item.ChildColumns?
                .CalculatePoints(point + new Size(offset.Width, item.RowSpan));
            item.Point = point + offset;
            offset.Width += item.ColumnSpan;
        }
    }

    /// <summary>
    /// Adds the specified column to the collection.
    /// </summary>
    /// <param name="dataColumn">The column to add.</param>
    /// <returns>The collection instance for chaining.</returns>
    public new RecordSetColumnCollection<T> Add(RecordSetColumnBase<T> dataColumn) {
        base.Add(dataColumn);
        return this;
    }

    /// <summary>
    /// Adds a generator column with the specified header text and optional styles.
    /// </summary>
    /// <param name="headerText">The header text.</param>
    /// <param name="headerStyle">The header style.</param>
    /// <param name="fieldStyleGenerator">The field style generator.</param>
    /// <returns>The collection instance for chaining.</returns>
    public RecordSetColumnCollection<T> Add(
        string headerText, CellStyle? headerStyle = null,
        Func<RecordContext<T>, CellStyle>? fieldStyleGenerator = null
    ) {
        GeneratorColumn<T> column = new();
        ApplyDefaultStyles(column, headerText, headerStyle, fieldStyleGenerator);
        Add(column);
        return this;
    }

    /// <summary>
    /// Adds a generator column with custom generator configuration.
    /// </summary>
    /// <param name="headerText">The header text.</param>
    /// <param name="configureGenerators">The action to configure column generators.</param>
    /// <param name="headerStyle">The header style.</param>
    /// <param name="fieldStyleGenerator">The field style generator.</param>
    /// <returns>The collection instance for chaining.</returns>
    public RecordSetColumnCollection<T> Add(
        string headerText, Action<ColumnGeneratorConfigurator<T, RecordContext<T>>> configureGenerators,
        CellStyle? headerStyle = null, Func<RecordContext<T>, CellStyle>? fieldStyleGenerator = null
    ) {
        ColumnGeneratorConfigurator<T, RecordContext<T>> provider = new();
        configureGenerators?.Invoke(provider);

        GeneratorColumn<T> dataColumn = new();
        ApplyDefaultStyles(dataColumn, headerText, headerStyle, fieldStyleGenerator);
        provider.SetGeneratorForColumn(dataColumn);
        Add(dataColumn);
        return this;
    }

    /// <summary>
    /// Adds a field column with the specified field key.
    /// </summary>
    /// <typeparam name="TField">The type of the field.</typeparam>
    /// <param name="headerText">The header text.</param>
    /// <param name="fieldKey">The field key.</param>
    /// <param name="headerStyle">The header style.</param>
    /// <param name="fieldStyleGenerator">The field style generator.</param>
    /// <returns>The collection instance for chaining.</returns>
    public RecordSetColumnCollection<T> Add<TField>(
        string headerText, string fieldKey,
        CellStyle? headerStyle = null, Func<FieldContext<T, TField>, CellStyle>? fieldStyleGenerator = null
    ) {
        FieldColumn<T, TField> column = new(fieldKey);
        ApplyDefaultStyles(column, headerText, headerStyle, fieldStyleGenerator);
        Add(column);
        return this;
    }

    /// <summary>
    /// Adds a field column using a property expression.
    /// </summary>
    /// <typeparam name="TField">The type of the field.</typeparam>
    /// <param name="headerText">The header text.</param>
    /// <param name="fieldKeyExpression">The field key expression.</param>
    /// <param name="headerStyle">The header style.</param>
    /// <param name="fieldStyleGenerator">The field style generator.</param>
    /// <returns>The collection instance for chaining.</returns>
    public RecordSetColumnCollection<T> Add<TField>(
        string headerText, Expression<Func<T, TField>> fieldKeyExpression,
        CellStyle? headerStyle = null, Func<FieldContext<T, TField>, CellStyle>? fieldStyleGenerator = null
    ) {
        FieldColumn<T, TField> column = new(fieldKeyExpression);
        ApplyDefaultStyles(column, headerText, headerStyle, fieldStyleGenerator);
        Add(column);
        return this;
    }

    /// <summary>
    /// Adds a field column with the specified field key and generator configuration.
    /// </summary>
    /// <typeparam name="TField">The type of the field.</typeparam>
    /// <param name="headerText">The header text.</param>
    /// <param name="fieldKey">The field key.</param>
    /// <param name="configureGenerators">The action to configure column generators.</param>
    /// <param name="headerStyle">The header style.</param>
    /// <param name="fieldStyleGenerator">The field style generator.</param>
    /// <returns>The collection instance for chaining.</returns>
    public RecordSetColumnCollection<T> Add<TField>(
        string headerText, string fieldKey, Action<ColumnGeneratorConfigurator<T, FieldContext<T, TField>>> configureGenerators,
        CellStyle? headerStyle = null, Func<FieldContext<T, TField>, CellStyle>? fieldStyleGenerator = null
    ) {
        ColumnGeneratorConfigurator<T, FieldContext<T, TField>> provider = new();
        configureGenerators?.Invoke(provider);

        FieldColumn<T, TField> dataColumn = new(fieldKey);
        ApplyDefaultStyles(dataColumn, headerText, headerStyle, fieldStyleGenerator);
        provider.SetGeneratorForColumn(dataColumn);
        Add(dataColumn);
        return this;
    }

    /// <summary>
    /// Adds a field column using a property expression and generator configuration.
    /// </summary>
    /// <typeparam name="TField">The type of the field.</typeparam>
    /// <param name="headerText">The header text.</param>
    /// <param name="fieldKeyExpression">The field key expression.</param>
    /// <param name="configureGenerators">The action to configure column generators.</param>
    /// <param name="headerStyle">The header style.</param>
    /// <param name="fieldStyleGenerator">The field style generator.</param>
    /// <returns>The collection instance for chaining.</returns>
    public RecordSetColumnCollection<T> Add<TField>(
        string headerText, Expression<Func<T, TField>> fieldKeyExpression,
        Action<ColumnGeneratorConfigurator<T, FieldContext<T, TField>>> configureGenerators,
        CellStyle? headerStyle = null, Func<FieldContext<T, TField>, CellStyle>? fieldStyleGenerator = null
    ) {
        ColumnGeneratorConfigurator<T, FieldContext<T, TField>> provider = new();
        configureGenerators?.Invoke(provider);

        FieldColumn<T, TField> dataColumn = new(fieldKeyExpression);
        ApplyDefaultStyles(dataColumn, headerText, headerStyle, fieldStyleGenerator);
        provider.SetGeneratorForColumn(dataColumn);
        Add(dataColumn);
        return this;
    }

    private static void ApplyDefaultStyles(
        GeneratorColumn<T> column, string headerText,
        CellStyle? headerStyle, Func<RecordContext<T>, CellStyle>? fieldStyleGenerator
    ) {
        column.HeaderText = headerText;
        column.HeaderStyle = headerStyle ?? SpreadsheetManager.DefaultCellStyles.HeaderStyle;
        column.FieldStyleGenerator = fieldStyleGenerator
            ?? ((context) => SpreadsheetManager.DefaultCellStyles.FieldStyle);
    }

    private static void ApplyDefaultStyles<TField>(
        FieldColumn<T, TField> column, string headerText,
        CellStyle? headerStyle, Func<FieldContext<T, TField>, CellStyle>? fieldStyleGenerator
    ) {
        column.HeaderText = headerText;
        column.HeaderStyle = headerStyle ?? SpreadsheetManager.DefaultCellStyles.HeaderStyle;
        column.FieldStyleGenerator = fieldStyleGenerator
            ?? ((context) => SpreadsheetManager.DefaultCellStyles.FieldStyle);
    }

    /// <inheritdoc/>
    protected override void InsertItem(int index, RecordSetColumnBase<T> item) {
        if (item.ParentCollection is not null) {
            throw new ArgumentException(
                $"RecordSetColumnBase<{typeof(T).Name}> with HeaderText '{item.HeaderText}' is already contained by another RecordSetColumnCollection. " +
                $"A column can only belong to one collection at a time.",
                nameof(item)
            );
        }

        base.InsertItem(index, item);
        item.ParentCollection = this;
    }

    /// <inheritdoc/>
    protected override void RemoveItem(int index) {
        base.RemoveItem(index);
        Items[index].ParentCollection = null;
    }

    /// <inheritdoc/>
    protected override void SetItem(int index, RecordSetColumnBase<T> item) {
        if (item.ParentCollection is not null) {
            throw new ArgumentException(
                $"RecordSetColumnBase<{typeof(T).Name}> with HeaderText '{item.HeaderText}' is already contained by another RecordSetColumnCollection. " +
                $"A column can only belong to one collection at a time.",
                nameof(item)
            );
        }

        base.SetItem(index, item);
        Items[index].ParentCollection = null;
        item.ParentCollection = this;
    }

    /// <inheritdoc/>
    protected override void ClearItems() {
        for (int i = 0; i < Items.Count; i++) {
            Items[i].ParentCollection = null;
        }
        base.ClearItems();
    }

    /// <summary>
    /// Gets the leaf columns that directly bind to data source properties.
    /// </summary>
    /// <value>The leaf columns.</value>
    public IEnumerable<RecordSetColumnBase<T>> LeafColumns {
        get {
            foreach (RecordSetColumnBase<T> item in this) {
                if (item.ChildColumns.Count == 0) {
                    yield return item;
                } else {
                    foreach (RecordSetColumnBase<T> leaf in item.ChildColumns.LeafColumns) {
                        yield return leaf;
                    }
                }
            }
        }
    }

    /// <summary>
    /// Configures value, formula, and data validation generators for a column.
    /// </summary>
    /// <typeparam name="TRecord">The type of the record.</typeparam>
    /// <typeparam name="TContext">The type of the context.</typeparam>
    public sealed class ColumnGeneratorConfigurator<TRecord, TContext> where TContext : RecordContext<TRecord> {
        private ProviderTypes types = ProviderTypes.None;
        private Func<TContext, object?>? valueGenerator;
        private Func<TContext, string?>? formulaGenerator;
        private Func<TContext, DataValidation?>? dataValidationGenerator;

        /// <summary>
        /// Configures the column to use a value generator.
        /// </summary>
        /// <param name="generator">The value generator.</param>
        public void UseValue(Func<TContext, object?> generator) {
            types |= ProviderTypes.Value;
            valueGenerator = generator;
        }

        /// <summary>
        /// Configures the column to use a formula generator.
        /// </summary>
        /// <param name="generator">The formula generator.</param>
        public void UseFormula(Func<TContext, string?> generator) {
            types |= ProviderTypes.Formula;
            formulaGenerator = generator;
        }

        /// <summary>
        /// Configures the column to use a data validation generator.
        /// </summary>
        /// <param name="dataValidationGenerator">The data validation generator.</param>
        public void UseDataValidation(Func<TContext, DataValidation?> dataValidationGenerator) {
            types |= ProviderTypes.DataValidation;
            this.dataValidationGenerator = dataValidationGenerator;
        }

        internal void SetGeneratorForColumn(GeneratorColumn<TRecord> dataColumn) {
            ValidateProviderTypes();

            if (types.HasFlag(ProviderTypes.Value)) {
                dataColumn.FieldValueGenerator = ConvertGenerator<Func<RecordContext<TRecord>, object?>>(valueGenerator!);
            }

            if (types.HasFlag(ProviderTypes.Formula)) {
                dataColumn.FieldFormulaGenerator = ConvertGenerator<Func<RecordContext<TRecord>, string?>>(formulaGenerator!);
            }

            if (types.HasFlag(ProviderTypes.DataValidation)) {
                dataColumn.FieldDataValidationGenerator = ConvertGenerator<Func<RecordContext<TRecord>, DataValidation?>>(dataValidationGenerator!);
            }
        }

        internal void SetGeneratorForColumn<TField>(FieldColumn<TRecord, TField> dataColumn) {
            ValidateProviderTypes();

            if (types.HasFlag(ProviderTypes.Value)) {
                dataColumn.FieldValueGenerator = ConvertGenerator<Func<FieldContext<TRecord, TField>, object?>>(valueGenerator!);
            }

            if (types.HasFlag(ProviderTypes.Formula)) {
                dataColumn.FieldFormulaGenerator = ConvertGenerator<Func<FieldContext<TRecord, TField>, string?>>(formulaGenerator!);
            }

            if (types.HasFlag(ProviderTypes.DataValidation)) {
                dataColumn.FieldDataValidationGenerator = ConvertGenerator<Func<FieldContext<TRecord, TField>, DataValidation?>>(dataValidationGenerator!);
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
