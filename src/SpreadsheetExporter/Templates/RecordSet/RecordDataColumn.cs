using System;

namespace CloudyWing.SpreadsheetExporter.Templates.RecordSet {
    /// <summary>
    /// The simple data column.
    /// </summary>
    /// <typeparam name="T">The type of the record.</typeparam>
    internal class RecordDataColumn<T> : DataColumnBase<T> {
        /// <summary>
        /// Gets the field value generator.
        /// </summary>
        /// <value>
        /// The field value generator.
        /// </value>
        public Func<RecordContext<T>, object> FieldValueGenerator { get; set; }

        /// <summary>
        /// Gets the field formula generator.
        /// </summary>
        /// <value>
        /// The field formula generator.
        /// </value>
        public Func<RecordContext<T>, string> FieldFormulaGenerator { get; set; }

        /// <summary>
        /// Gets or sets the field style generator.
        /// </summary>
        /// <value>
        /// The field style generator.
        /// </value>
        public Func<RecordContext<T>, CellStyle> FieldStyleGenerator { get; set; }

        /// <inheritdoc/>
        public override object GetFieldValue(RecordContext<T> context) {
            return GetInternal(FieldValueGenerator, null, context);
        }

        /// <inheritdoc/>
        public override string GetFieldFormula(RecordContext<T> context) {
            return GetInternal(FieldFormulaGenerator, null, context);
        }

        /// <inheritdoc/>
        public override CellStyle GetFieldStyle(RecordContext<T> context) {
            return GetInternal(FieldStyleGenerator, SpreadsheetManager.DefaultCellStyles.FieldStyle, context);
        }

        private TResult GetInternal<TResult>(
            Func<RecordContext<T>, TResult> functor, TResult defaultValue, RecordContext<T> context
        ) {
            return functor is null ? defaultValue : functor(context);
        }
    }
}
