namespace CloudyWing.SpreadsheetExporter;

/// <summary>
/// Defines the types of data validation that can be applied to cells.
/// </summary>
public enum DataValidationType {
    /// <summary>
    /// List validation - restricts input to values from a predefined list.
    /// </summary>
    List,

    /// <summary>
    /// Integer validation - restricts input to whole numbers.
    /// </summary>
    Integer,

    /// <summary>
    /// Decimal validation - restricts input to decimal numbers.
    /// </summary>
    Decimal,

    /// <summary>
    /// Date validation - restricts input to dates.
    /// </summary>
    Date,

    /// <summary>
    /// Time validation - restricts input to time values.
    /// </summary>
    Time,

    /// <summary>
    /// Text length validation - restricts the length of text input.
    /// </summary>
    TextLength,

    /// <summary>
    /// Custom validation - allows custom validation formulas.
    /// </summary>
    Custom
}
