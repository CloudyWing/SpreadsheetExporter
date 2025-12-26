using System.Collections.Generic;

namespace CloudyWing.SpreadsheetExporter;

/// <summary>
/// Represents data validation rules that can be applied to spreadsheet cells.
/// Data validation restricts the type of data or values that users can enter into a cell.
/// </summary>
public class DataValidation {
    /// <summary>
    /// Gets or sets the type of validation to apply.
    /// </summary>
    /// <value>
    /// The validation type.
    /// </value>
    public DataValidationType ValidationType { get; set; }

    /// <summary>
    /// Gets or sets the comparison operator for the validation.
    /// Not required for List or Custom validation types.
    /// </summary>
    /// <value>
    /// The operator, or <c>null</c> if not applicable.
    /// </value>
    public DataValidationOperator? Operator { get; set; }

    /// <summary>
    /// Gets or sets the first value for comparison.
    /// For Between/NotBetween operators, this is the minimum value.
    /// For other operators, this is the value to compare against.
    /// </summary>
    /// <value>
    /// The first comparison value.
    /// </value>
    public object Value1 { get; set; }

    /// <summary>
    /// Gets or sets the second value for comparison.
    /// Only used with Between and NotBetween operators (represents the maximum value).
    /// </summary>
    /// <value>
    /// The second comparison value, or <c>null</c> if not applicable.
    /// </value>
    public object Value2 { get; set; }

    /// <summary>
    /// Gets or sets the list of valid items for List validation type.
    /// Only used when ValidationType is List.
    /// </summary>
    /// <value>
    /// The list items.
    /// </value>
    public IEnumerable<string> ListItems { get; set; }

    /// <summary>
    /// Gets or sets the formula for validation.
    /// Used with Custom validation type (required), and can also be used with Date, Time, Integer, Decimal, and TextLength validation types as an alternative to Value1/Value2.
    /// When using formulas with Date/Time/Numeric validations, the formula will be automatically prefixed with '=' if not already present.
    /// </summary>
    /// <value>
    /// The validation formula (e.g., "TODAY()", "=A1+10").
    /// </value>
    public string Formula { get; set; }

    /// <summary>
    /// Gets or sets a value indicating whether the dropdown button is shown for list validation.
    /// Only applicable when ValidationType is List.
    /// </summary>
    /// <value>
    /// <c>true</c> if dropdown is shown; otherwise, <c>false</c>. Default is <c>true</c>.
    /// </value>
    public bool IsDropdownShown { get; set; } = true;

    /// <summary>
    /// Gets or sets a value indicating whether blank/empty values are allowed.
    /// </summary>
    /// <value>
    /// <c>true</c> if blank values are allowed; otherwise, <c>false</c>. Default is <c>true</c>.
    /// </value>
    public bool IsBlankAllowed { get; set; } = true;

    /// <summary>
    /// Gets or sets the title of the error message shown when invalid data is entered.
    /// </summary>
    /// <value>
    /// The error title.
    /// </value>
    public string ErrorTitle { get; set; }

    /// <summary>
    /// Gets or sets the error message shown when invalid data is entered.
    /// </summary>
    /// <value>
    /// The error message.
    /// </value>
    public string ErrorMessage { get; set; }

    /// <summary>
    /// Gets or sets a value indicating whether the error alert is shown when invalid data is entered.
    /// </summary>
    /// <value>
    /// <c>true</c> if error alert is shown; otherwise, <c>false</c>. Default is <c>true</c>.
    /// </value>
    public bool IsErrorAlertShown { get; set; } = true;

    /// <summary>
    /// Gets or sets the title of the input prompt message.
    /// </summary>
    /// <value>
    /// The prompt title.
    /// </value>
    public string PromptTitle { get; set; }

    /// <summary>
    /// Gets or sets the input prompt message shown when the cell is selected.
    /// </summary>
    /// <value>
    /// The prompt message.
    /// </value>
    public string PromptMessage { get; set; }

    /// <summary>
    /// Gets or sets a value indicating whether the input prompt is shown when the cell is selected.
    /// </summary>
    /// <value>
    /// <c>true</c> if input prompt is shown; otherwise, <c>false</c>. Default is <c>false</c>.
    /// </value>
    public bool IsInputPromptShown { get; set; }
}
