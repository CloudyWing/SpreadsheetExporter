using System;

namespace CloudyWing.SpreadsheetExporter.Exceptions;

/// <summary>
/// The exception thrown when no sheet definitions are found in the document.
/// </summary>
public class SheetDefinitionNotFoundException : Exception {
    /// <summary>
    /// Initializes a new instance of the <see cref="SheetDefinitionNotFoundException"/> class.
    /// </summary>
    public SheetDefinitionNotFoundException()
        : this("Could not find SheetDefinition.") { }

    /// <summary>
    /// Initializes a new instance of the <see cref="SheetDefinitionNotFoundException"/> class.
    /// </summary>
    /// <param name="message">The message that describes the error.</param>
    public SheetDefinitionNotFoundException(string message) : base(message) { }

    /// <summary>
    /// Initializes a new instance of the <see cref="SheetDefinitionNotFoundException"/> class.
    /// </summary>
    /// <param name="message">The error message.</param>
    /// <param name="innerException">The inner exception.</param>
    public SheetDefinitionNotFoundException(string message, Exception innerException)
        : base(message, innerException) { }
}
