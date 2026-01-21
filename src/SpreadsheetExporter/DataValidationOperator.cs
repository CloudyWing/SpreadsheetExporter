namespace CloudyWing.SpreadsheetExporter;

/// <summary>
/// Defines the comparison operators used in data validation.
/// </summary>
public enum DataValidationOperator {
    /// <summary>
    /// Value must be between two values (inclusive).
    /// </summary>
    Between,

    /// <summary>
    /// Value must not be between two values.
    /// </summary>
    NotBetween,

    /// <summary>
    /// Value must be equal to a specified value.
    /// </summary>
    Equal,

    /// <summary>
    /// Value must not be equal to a specified value.
    /// </summary>
    NotEqual,

    /// <summary>
    /// Value must be greater than a specified value.
    /// </summary>
    GreaterThan,

    /// <summary>
    /// Value must be less than a specified value.
    /// </summary>
    LessThan,

    /// <summary>
    /// Value must be greater than or equal to a specified value.
    /// </summary>
    GreaterThanOrEqual,

    /// <summary>
    /// Value must be less than or equal to a specified value.
    /// </summary>
    LessThanOrEqual
}
