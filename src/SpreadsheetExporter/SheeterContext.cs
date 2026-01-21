using System.Collections.Generic;
using CloudyWing.SpreadsheetExporter.Extensions;
using CloudyWing.SpreadsheetExporter.Templates;

namespace CloudyWing.SpreadsheetExporter;

/// <summary>
/// The sheeter context.
/// </summary>
public class SheeterContext {
    /// <summary>
    /// Initializes a new instance of the <see cref="SheeterContext" /> class.
    /// </summary>
    /// <param name="sheeter">The sheeter.</param>
    public SheeterContext(Sheeter sheeter) {
        SheetName = sheeter.SheetName;
        DefaultRowHeight = sheeter.DefaultRowHeight;
        TemplateContext templateContext = TemplateContext.Create(sheeter.Templates);
        Cells = templateContext.Cells;
        RowHeights = templateContext.RowHeights.AsReadOnly();
        ColumnWidths = sheeter.ColumnWidths.AsReadOnly();
        Password = sheeter.Password;
        PageSettings = sheeter.PageSettings;
        FreezePanes = templateContext.FreezePanes;
        IsAutoFilterEnabled = templateContext.IsAutoFilterEnabled;
    }

    /// <summary>
    /// Gets the name of the sheet.
    /// </summary>
    /// <value>
    /// The name of the sheet.
    /// </value>
    public string SheetName { get; }

    /// <summary>
    /// Gets or sets the default height of the row.
    /// </summary>
    /// <value>
    /// The default height of the row.
    /// </value>
    public double? DefaultRowHeight { get; }

    /// <summary>
    /// Gets the cells.
    /// </summary>
    /// <value>
    /// The cells.
    /// </value>
    public IReadOnlyList<Cell> Cells { get; }

    /// <summary>
    /// Gets the width of columns.
    /// </summary>
    /// <value>
    /// The width of columns.
    /// </value>
    public IReadOnlyDictionary<int, double> ColumnWidths { get; }

    /// <summary>
    /// Gets the height of rows.
    /// </summary>
    /// <value>
    /// The height of rows.
    /// </value>
    public IReadOnlyDictionary<int, double?> RowHeights { get; }

    /// <summary>
    /// Gets the password.
    /// </summary>
    /// <value>
    /// The password.
    /// </value>
    public string Password { get; }

    /// <summary>
    /// Gets a value indicating whether this instance is protected.
    /// </summary>
    /// <value>
    ///   <c>true</c> if this instance is protected; otherwise, <c>false</c>.</value>
    public bool IsProtected => !string.IsNullOrEmpty(Password);

    /// <summary>
    /// Gets the page settings.
    /// </summary>
    /// <value>
    /// The page settings.
    /// </value>
    public PageSettings PageSettings { get; }

    /// <summary>
    /// Gets the freeze panes position. The column and row at this position will be the first unfrozen cell.
    /// For example, (1, 1) freezes the first column and first row (A1 becomes the frozen pane).
    /// </summary>
    /// <value>
    /// The freeze panes position. Null means no freeze panes.
    /// </value>
    public System.Drawing.Point? FreezePanes { get; }

    /// <summary>
    /// Gets a value indicating whether auto filter is enabled for the data range.
    /// </summary>
    /// <value>
    /// <c>true</c> if auto filter is enabled; otherwise, <c>false</c>.
    /// </value>
    public bool IsAutoFilterEnabled { get; }
}
