using System;

namespace CloudyWing.SpreadsheetExporter {
    /// <summary>Event arguments after sheet create.</summary>
    public class SheetCreatedEventArgs {
        /// <summary>Initializes a new instance of the <see cref="SpreadsheetExportedEventArgs" /> class.</summary>
        /// <param name="sheetObject">The sheet object that represents the Excel sheet from the underlying Excel library.</param>
        /// <param name="sheeterContext">The sheeter contexts.</param>
        /// <exception cref="ArgumentNullException">sheetObject
        /// or
        /// sheeterContext</exception>
        public SheetCreatedEventArgs(object sheetObject, SheeterContext sheeterContext) {
            SheetObject = sheetObject ?? throw new ArgumentNullException(nameof(sheetObject));
            SheeterContext = sheeterContext ?? throw new ArgumentNullException(nameof(sheeterContext));
        }

        /// <summary>Gets or sets the sheet object that represents the Excel sheet from the underlying Excel library.</summary>
        /// <value>The sheet object.</value>
        public object SheetObject { get; }

        /// <summary>Gets the sheeter contexts.</summary>
        /// <value>The sheeter contexts.</value>
        public SheeterContext SheeterContext { get; }
    }
}
