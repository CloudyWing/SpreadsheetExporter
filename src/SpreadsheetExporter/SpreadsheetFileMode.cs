namespace CloudyWing.SpreadsheetExporter {
    /// <summary>
    /// The file mode of spreadsheet
    /// </summary>
    public enum SpreadsheetFileMode {
        /// <summary>
        /// Create a file. If the file already exists, it will be overwritten.
        /// </summary>
        Create,

        /// <summary>
        /// Create a new file. If the file already exists, an IOException exception is thrown.
        /// </summary>
        CreateNew
    }
}
