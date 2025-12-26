using System;

namespace CloudyWing.SpreadsheetExporter.Templates.DataTable;

/// <summary>
/// The DataTable column.
/// </summary>
public class DataTableColumn {
    /// <summary>
    /// Gets or sets the name of the column in the DataTable.
    /// </summary>
    /// <value>
    /// The name of the column.
    /// </value>
    public string ColumnName { get; set; }

    /// <summary>
    /// Gets or sets the header text.
    /// </summary>
    /// <value>
    /// The header text.
    /// </value>
    public string HeaderText { get; set; }

    /// <summary>
    /// Gets or sets the header style.
    /// </summary>
    /// <value>
    /// The header style.
    /// </value>
    public CellStyle HeaderStyle { get; set; }

    /// <summary>
    /// Gets or sets the field value generator.
    /// </summary>
    /// <value>
    /// The field value generator.
    /// </value>
    public Func<object, object> FieldValueGenerator { get; set; }

    /// <summary>
    /// Gets or sets the field formula generator.
    /// </summary>
    /// <value>
    /// The field formula generator.
    /// </value>
    public Func<object, string> FieldFormulaGenerator { get; set; }

    /// <summary>
    /// Gets or sets the field style generator.
    /// </summary>
    /// <value>
    /// The field style generator.
    /// </value>
    public Func<object, CellStyle> FieldStyleGenerator { get; set; }
}
