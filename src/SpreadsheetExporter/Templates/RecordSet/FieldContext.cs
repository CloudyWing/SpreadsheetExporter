using System;

namespace CloudyWing.SpreadsheetExporter.Templates.RecordSet {
    /// <summary>The field context.</summary>
    /// <typeparam name="TField">The type of the field.</typeparam>
    /// <typeparam name="TRecoed">The type of the recoed.</typeparam>
    public class FieldContext<TField, TRecoed> : RecordContext<TRecoed> {
        /// <summary>Initializes a new instance of the <see cref="FieldContext{TField, TRecoed}" /> class.</summary>
        /// <param name="recordContext">The record context.</param>
        /// <param name="key">The key.</param>
        /// <param name="value">The value.</param>
        /// <exception cref="ArgumentNullException">key</exception>
        public FieldContext(RecordContext<TRecoed> recordContext, string key, TField value)
            : base(recordContext.HeaderRowSpan, recordContext.RecordIndex, recordContext.Record) {
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
