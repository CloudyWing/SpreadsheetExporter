using System;
using System.Collections.Generic;

namespace CloudyWing.SpreadsheetExporter {
    /// <summary>
    /// Event arguments before spreadsheet export.
    /// </summary>
    public class SpreadsheetExportingEventArgs {
        /// <summary>
        /// Initializes a new instance of the <see cref="SpreadsheetExportingEventArgs" /> class.
        /// </summary>
        /// <param name="sheeterContexts">The sheeter contexts.</param>
        /// <exception cref="ArgumentNullException">sheeterContexts</exception>
        public SpreadsheetExportingEventArgs(IEnumerable<SheeterContext> sheeterContexts) {
            SheeterContexts = sheeterContexts ?? throw new ArgumentNullException(nameof(sheeterContexts));
        }

        /// <summary>
        /// Gets the sheeter contexts.
        /// </summary>
        /// <value>
        /// The sheeter contexts.
        /// </value>
        public IEnumerable<SheeterContext> SheeterContexts { get; }
    }
}
