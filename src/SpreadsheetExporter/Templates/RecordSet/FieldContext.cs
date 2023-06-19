using System;

namespace CloudyWing.SpreadsheetExporter.Templates.RecordSet {
    /// <summary>The field context.</summary>
    /// <typeparam name="TRecoed">The type of the recoed.</typeparam>
    /// <typeparam name="TField">The type of the field.</typeparam>
    public class FieldContext<TRecoed, TField> : RecordContext<TRecoed> {
        /// <summary>Initializes a new instance of the <see cref="FieldContext{TRecoed, TField}" /> class.</summary>
        /// <param name="recordContext">The record context.</param>
        /// <param name="key">The key.</param>
        /// <param name="value">The value.</param>
        /// <exception cref="ArgumentNullException">key</exception>
        public FieldContext(RecordContext<TRecoed> recordContext, string key, TField value)
            : base(recordContext.CellIndex, recordContext.RowIndex, recordContext.Record) {
            Key = key ?? throw new ArgumentNullException(nameof(key));
            Value = value;
        }

        /// <summary>Gets the key.</summary>
        /// <value>The key.</value>
        public string Key { get; }

        /// <summary>Gets the value.</summary>
        /// <value>The value.</value>
        public TField Value { get; }
    }
}
