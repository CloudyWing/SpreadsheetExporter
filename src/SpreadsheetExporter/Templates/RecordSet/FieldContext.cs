using System;

namespace CloudyWing.SpreadsheetExporter.Templates.RecordSet {
    /// <summary>
    /// The field context.
    /// </summary>
    /// <typeparam name="TRecord">The type of the record.</typeparam>
    /// <typeparam name="TField">The type of the field.</typeparam>
    /// <param name="recordContext">The record context.</param>
    /// <param name="key">The key.</param>
    /// <param name="value">The value.</param>
    /// <exception cref="ArgumentNullException">key</exception>
    public class FieldContext<TRecord, TField>(RecordContext<TRecord> recordContext, string key, TField value)
        : RecordContext<TRecord>(recordContext.CellIndex, recordContext.RowIndex, recordContext.Record) {
        /// <summary>
        /// Gets the key.
        /// </summary>
        /// <value>
        /// The key.
        /// </value>
        public string Key { get; } = key ?? throw new ArgumentNullException(nameof(key));

        /// <summary>
        /// Gets the value.
        /// </summary>
        /// <value>
        /// The value.
        /// </value>
        public TField Value { get; } = value;
    }
}
