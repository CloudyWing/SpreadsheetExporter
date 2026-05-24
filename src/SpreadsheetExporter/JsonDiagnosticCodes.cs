namespace CloudyWing.SpreadsheetExporter;

/// <summary>
/// Defines diagnostic codes produced while parsing spreadsheet JSON.
/// </summary>
public static class JsonDiagnosticCodes {
    /// <summary>
    /// Indicates that a JSON value has an invalid type.
    /// </summary>
    public const string InvalidType = "SE-JSON-001";

    /// <summary>
    /// Indicates that a required JSON property is missing.
    /// </summary>
    public const string MissingRequiredProperty = "SE-JSON-002";

    /// <summary>
    /// Indicates that a JSON value is invalid for the target model.
    /// </summary>
    public const string InvalidValue = "SE-JSON-003";
}
