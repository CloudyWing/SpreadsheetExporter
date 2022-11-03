namespace CloudyWing.SpreadsheetExporter.Templates.RecordSet {
    /// <summary>The record context.</summary>
    /// <typeparam name="T">The type of the record.</typeparam>
    public class RecordContext<T> {
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
