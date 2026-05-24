using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using CloudyWing.SpreadsheetExporter;

namespace CloudyWing.SpreadsheetExporter.Exceptions;

/// <summary>
/// Represents errors found while validating spreadsheet layout.
/// </summary>
public sealed class SpreadsheetLayoutException : Exception {
    /// <summary>
    /// Initializes a new instance of the <see cref="SpreadsheetLayoutException" /> class.
    /// </summary>
    /// <param name="diagnostics">The layout diagnostics.</param>
    public SpreadsheetLayoutException(IEnumerable<LayoutDiagnostic> diagnostics)
        : this("Spreadsheet layout validation failed.", diagnostics) {
    }

    /// <summary>
    /// Initializes a new instance of the <see cref="SpreadsheetLayoutException" /> class.
    /// </summary>
    /// <param name="message">The exception message.</param>
    /// <param name="diagnostics">The layout diagnostics.</param>
    public SpreadsheetLayoutException(string message, IEnumerable<LayoutDiagnostic> diagnostics)
        : base(message) {
        ArgumentNullException.ThrowIfNull(diagnostics);
        Diagnostics = new ReadOnlyCollection<LayoutDiagnostic>(diagnostics.ToList());
    }

    /// <summary>
    /// Gets the layout diagnostics that caused this exception.
    /// </summary>
    public IReadOnlyList<LayoutDiagnostic> Diagnostics { get; }
}
