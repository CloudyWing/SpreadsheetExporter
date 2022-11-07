using System;
using System.Collections.Generic;

namespace CloudyWing.SpreadsheetExporter {
    /// <summary>Event arguments after spreadsheet export.</summary>
    public class SpreadsheetExportedEventArgs {
        /// <summary>Initializes a new instance of the <see cref="SpreadsheetExportedEventArgs" /> class.</summary>
        /// <param name="sheeterContexts">The sheeter contexts.</param>
        /// <exception cref="ArgumentNullException">sheeterContexts</exception>
        public SpreadsheetExportedEventArgs(IEnumerable<SheeterContext> sheeterContexts) {
            SheeterContexts = sheeterContexts ?? throw new ArgumentNullException(nameof(sheeterContexts));
        }

        /// <summary>Gets the sheeter contexts.</summary>
        /// <value>The sheeter contexts.</value>
        public IEnumerable<SheeterContext> SheeterContexts { get; }
    }
}
