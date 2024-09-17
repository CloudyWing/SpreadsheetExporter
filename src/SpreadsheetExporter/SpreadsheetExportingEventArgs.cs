using System;
using System.Collections.Generic;

namespace CloudyWing.SpreadsheetExporter {
    /// <summary>
    /// Event arguments before spreadsheet export.
    /// </summary>
    /// <param name="sheeterContexts">The sheeter contexts.</param>
    /// <exception cref="ArgumentNullException">sheeterContexts</exception>
    public class SpreadsheetExportingEventArgs(IEnumerable<SheeterContext> sheeterContexts) {

        /// <summary>
        /// Gets the sheeter contexts.
        /// </summary>
        /// <value>
        /// The sheeter contexts.
        /// </value>
        public IEnumerable<SheeterContext> SheeterContexts { get; } =
            sheeterContexts ?? throw new ArgumentNullException(nameof(sheeterContexts));
    }
}
