using System;

namespace CloudyWing.SpreadsheetExporter {
    /// <summary>
    /// Event arguments after sheet create.
    /// </summary>
    /// <param name="sheetObject">The sheet object that represents the Excel sheet from the underlying Excel library.</param>
    /// <param name="sheeterContext">The sheeter contexts.</param>
    /// <exception cref="ArgumentNullException">sheetObject
    /// or
    /// sheeterContext</exception>
    public class SheetCreatedEventArgs(object sheetObject, SheeterContext sheeterContext) {

        /// <summary>
        /// Gets or sets the sheet object that represents the Excel sheet from the underlying Excel library.
        /// </summary>
        /// <value>
        /// The sheet object.
        /// </value>
        public object SheetObject { get; } =
            sheetObject ?? throw new ArgumentNullException(nameof(sheetObject));

        /// <summary>
        /// Gets the sheeter contexts.
        /// </summary>
        /// <value>
        /// The sheeter contexts.
        /// </value>
        public SheeterContext SheeterContext { get; } =
            sheeterContext ?? throw new ArgumentNullException(nameof(sheeterContext));
    }
}
