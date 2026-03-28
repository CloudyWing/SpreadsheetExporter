using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Drawing;
using CloudyWing.SpreadsheetExporter.Templates;

namespace CloudyWing.SpreadsheetExporter;

/// <summary>
/// Represents a single sheet definition within a spreadsheet document,
/// containing templates, column widths, and page settings.
/// </summary>
public class SheetDefinition {
    private readonly Dictionary<int, double> columnWidths = [];
    private readonly List<ISheetTemplate> templates = [];
    private ReadOnlyDictionary<int, double>? columnWidthsReadOnly;
    private ReadOnlyCollection<ISheetTemplate>? templatesReadOnly;

    internal SheetDefinition(string sheetName) {
        SheetName = sheetName;
    }

    /// <summary>
    /// Gets or sets the name of the sheet.
    /// </summary>
    /// <value>The sheet name.</value>
    public string SheetName { get; set; }

    /// <summary>
    /// Gets or sets the default row height. <see langword="null"/> uses the renderer default.
    /// </summary>
    /// <value>The default row height in points, or <see langword="null"/>.</value>
    public double? DefaultRowHeight { get; set; }

    /// <summary>
    /// Gets or sets the worksheet protection password.
    /// </summary>
    /// <value>The password, or <see langword="null"/> for no protection.</value>
    public string? Password { get; set; }

    /// <summary>
    /// Gets the page settings for this sheet.
    /// </summary>
    /// <value>The page settings.</value>
    public PageSettings PageSettings { get; } = new PageSettings();

    /// <summary>
    /// Gets the column widths indexed by zero-based column index.
    /// </summary>
    /// <value>A read-only dictionary of column widths.</value>
    public IReadOnlyDictionary<int, double> ColumnWidths
        => columnWidthsReadOnly ??= new ReadOnlyDictionary<int, double>(columnWidths);

    /// <summary>
    /// Gets the templates added to this sheet.
    /// </summary>
    /// <value>A read-only list of sheet templates.</value>
    public IReadOnlyList<ISheetTemplate> Templates
        => templatesReadOnly ??= new ReadOnlyCollection<ISheetTemplate>(templates);

    /// <summary>
    /// Gets or sets the freeze panes position (zero-based column and row). The cell at this position becomes the first unfrozen cell.
    /// </summary>
    /// <value>The freeze panes position, or <see langword="null"/> for no freeze.</value>
    public Point? FreezePanes { get; set; }

    /// <summary>
    /// Gets or sets a value indicating whether auto filter is enabled for the data range.
    /// </summary>
    /// <value><see langword="true"/> to enable auto filter; otherwise, <see langword="false"/>. Default is <see langword="false"/>.</value>
    public bool IsAutoFilterEnabled { get; set; }

    /// <summary>
    /// Gets a dictionary of arbitrary metadata for renderer-specific extensions.
    /// </summary>
    /// <value>The metadata dictionary.</value>
    public IDictionary<string, object?> Metadata { get; } = new Dictionary<string, object?>();

    /// <summary>
    /// Sets the name of the sheet.
    /// </summary>
    /// <param name="sheetName">The sheet name.</param>
    /// <returns>The current <see cref="SheetDefinition"/> instance.</returns>
    public SheetDefinition SetSheetName(string sheetName) {
        SheetName = sheetName;
        return this;
    }

    /// <summary>
    /// Sets the default row height for this sheet.
    /// </summary>
    /// <param name="defaultRowHeight">The default row height in points, or <see langword="null"/> to use the renderer default.</param>
    /// <returns>The current <see cref="SheetDefinition"/> instance.</returns>
    public SheetDefinition SetDefaultRowHeight(double? defaultRowHeight) {
        DefaultRowHeight = defaultRowHeight;
        return this;
    }

    /// <summary>
    /// Sets the worksheet protection password.
    /// </summary>
    /// <param name="password">The password, or <see langword="null"/> to remove protection.</param>
    /// <returns>The current <see cref="SheetDefinition"/> instance.</returns>
    public SheetDefinition SetPassword(string? password) {
        Password = password;
        return this;
    }

    /// <summary>
    /// Sets the freeze panes position for this sheet.
    /// </summary>
    /// <param name="column">The zero-based column index of the first unfrozen column.</param>
    /// <param name="row">The zero-based row index of the first unfrozen row.</param>
    /// <returns>The current <see cref="SheetDefinition"/> instance.</returns>
    public SheetDefinition SetFreezePanes(int column, int row) {
        FreezePanes = new Point(column, row);
        return this;
    }

    /// <summary>
    /// Sets a value indicating whether auto filter is enabled for this sheet.
    /// </summary>
    /// <param name="enabled">Whether to enable auto filter.</param>
    /// <returns>The current <see cref="SheetDefinition"/> instance.</returns>
    public SheetDefinition SetAutoFilterEnabled(bool enabled = true) {
        IsAutoFilterEnabled = enabled;
        return this;
    }

    /// <summary>
    /// Configures the page settings for this sheet.
    /// </summary>
    /// <param name="configure">An action that configures the <see cref="PageSettings"/>.</param>
    /// <returns>The current <see cref="SheetDefinition"/> instance.</returns>
    public SheetDefinition ConfigurePageSettings(Action<PageSettings> configure) {
        ArgumentNullException.ThrowIfNull(configure);
        configure(PageSettings);
        return this;
    }

    /// <summary>
    /// Sets the width of a column.
    /// </summary>
    /// <param name="index">The zero-based column index.</param>
    /// <param name="width">The width in character units. Use <see cref="Constants.HiddenColumn"/> to hide,
    /// or <see cref="Constants.AutoFitColumnWidth"/> to auto-fit.</param>
    /// <returns>The current <see cref="SheetDefinition"/> instance.</returns>
    public SheetDefinition SetColumnWidth(int index, double width) {
        columnWidths[index] = width;
        return this;
    }

    /// <summary>
    /// Adds a template to this sheet.
    /// </summary>
    /// <param name="template">The template to add.</param>
    /// <returns>The current <see cref="SheetDefinition"/> instance.</returns>
    public SheetDefinition AddTemplate(ISheetTemplate template) {
        templates.Add(template);
        return this;
    }

    /// <summary>
    /// Adds multiple templates to this sheet.
    /// </summary>
    /// <param name="templates">The templates to add.</param>
    /// <returns>The current <see cref="SheetDefinition"/> instance.</returns>
    public SheetDefinition AddTemplates(IEnumerable<ISheetTemplate> templates) {
        foreach (ISheetTemplate template in templates) {
            AddTemplate(template);
        }
        return this;
    }

    /// <summary>
    /// Adds multiple templates to this sheet.
    /// </summary>
    /// <param name="templates">The templates to add.</param>
    /// <returns>The current <see cref="SheetDefinition"/> instance.</returns>
    public SheetDefinition AddTemplates(params ISheetTemplate[] templates) {
        return AddTemplates(templates as IEnumerable<ISheetTemplate>);
    }
}
