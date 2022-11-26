namespace CloudyWing.SpreadsheetExporter.Templates.RecordSet {
    /// <summary>The record context.</summary>
    /// <typeparam name="T">The type of the record.</typeparam>
    public class RecordContext<T> {
        /// <summary>Initializes a new instance of the <see cref="RecordContext{T}" /> class.</summary>
        /// <param name="headerRowSpan">The header row span.</param>
        /// <param name="recordIndex">Index of the record.</param>
        /// <param name="record">The record.</param>
        public RecordContext(int headerRowSpan, int recordIndex, T record) {
            HeaderRowSpan = headerRowSpan;
            RecordIndex = recordIndex;
            Record = record;
        }

        /// <summary>Gets the header row span.</summary>
        /// <value>The header row span.</value>
        public int HeaderRowSpan { get; }

        /// <summary>Gets the index of the record.</summary>
        /// <value>The index of the record.</value>
        public int RecordIndex { get; }

        /// <summary>Gets the record.</summary>
        /// <value>The record.</value>
        public T Record { get; }
    }
}
