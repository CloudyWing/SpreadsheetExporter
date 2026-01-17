using System;

namespace CloudyWing.SpreadsheetExporter {
    /// <summary>
    /// Interface for spreadsheet exporters.
    /// </summary>
    public interface ISpreadsheetExporter {
        /// <summary>
        /// Occurs when [spreadsheet exporting event].
        /// </summary>
        event EventHandler<SpreadsheetExportingEventArgs> SpreadsheetExportingEvent;

        /// <summary>
        /// Occurs when [spreadsheet exported event].
        /// </summary>
        event EventHandler<SpreadsheetExportedEventArgs> SpreadsheetExportedEvent;

        /// <summary>
        /// Occurs when [sheet created event].
        /// </summary>
        event EventHandler<SheetCreatedEventArgs> SheetCreatedEvent;

        /// <summary>
        /// Gets the content-type.
        /// </summary>
        /// <value>The content-type.</value>
        string ContentType { get; }

        /// <summary>
        /// Gets the file name extension.
        /// </summary>
        /// <value>The file name extension.</value>
        string FileNameExtension { get; }

        /// <summary>
        /// Gets or sets the workbook protection password.
        /// </summary>
        /// <value>The password.</value>
        string Password { get; set; }

        /// <summary>
        /// Gets a value indicating whether this instance has password.
        /// </summary>
        /// <value><c>true</c> if this instance has password; otherwise, <c>false</c>.</value>
        bool HasPassword { get; }

        /// <summary>
        /// Gets or sets the default font.
        /// </summary>
        /// <value>The default font.</value>
        CellFont? DefaultFont { get; set; }

        /// <summary>
        /// Gets or sets the default basic name of the sheet.
        /// </summary>
        /// <value>The default basic name of the sheet.</value>
        string DefaultBasicSheetName { get; set; }

        /// <summary>
        /// Gets the last sheeter. When there is no sheeter, create a sheeter.
        /// </summary>
        /// <value>The last sheeter.</value>
        Sheeter LastSheeter { get; }

        /// <summary>
        /// Creates the sheeter.
        /// </summary>
        /// <param name="sheetName">Name of the sheet.</param>
        /// <param name="defaultRowHeight">Default height of the row.</param>
        /// <returns>The sheeter.</returns>
        Sheeter CreateSheeter(string sheetName = null, double? defaultRowHeight = null);

        /// <summary>
        /// Exports bytes of spreadsheet.
        /// </summary>
        /// <returns>The bytes of spreadsheet.</returns>
        /// <exception cref="Exceptions.SheeterNotFoundException">No sheeters have been created.</exception>
        byte[] Export();

        /// <summary>
        /// Exports the spreadsheet file.
        /// </summary>
        /// <param name="path">The path.</param>
        /// <param name="fileMode">The file mode.</param>
        void ExportFile(string path, SpreadsheetFileMode fileMode = SpreadsheetFileMode.Create);

        /// <summary>
        /// Gets the sheeter.
        /// </summary>
        /// <param name="index">The index.</param>
        /// <returns>The sheeter.</returns>
        /// <exception cref="ArgumentOutOfRangeException">
        /// <paramref name="index"/> is less than 0 or greater than or equal to the number of sheeters.
        /// </exception>
        Sheeter GetSheeter(int index);
    }
}
