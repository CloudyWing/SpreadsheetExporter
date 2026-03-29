using System;
using System.Collections.Generic;
using System.Drawing;
using CloudyWing.SpreadsheetExporter.Util;
using CloudyWing.SpreadsheetExporter.Templates;

namespace CloudyWing.SpreadsheetExporter;

/// <summary>
/// Represents the resolved rendering layout for a single sheet,
/// containing pre-calculated cells, dimensions, and settings.
/// </summary>
public class SheetLayout {
    /// <summary>
    /// Initializes a new instance of the <see cref="SheetLayout"/> class.
    /// </summary>
    /// <param name="sheet">The sheet definition to resolve.</param>
    /// <param name="defaultFont">The default font for the workbook.</param>
    public SheetLayout(SheetDefinition sheet, CellFont? defaultFont = null) {
        ArgumentNullException.ThrowIfNull(sheet);

        SheetName = sheet.SheetName;
        DefaultRowHeight = sheet.DefaultRowHeight;
        TemplateLayout templateLayout = TemplateLayout.Create(sheet.Templates);
        Cells = templateLayout.Cells;
        RowHeights = templateLayout.RowHeights.AsReadOnly();
        ColumnWidths = sheet.ColumnWidths.AsReadOnly();
        Password = sheet.Password;
        PageSettings = sheet.PageSettings;
        FreezePanes = sheet.FreezePanes;
        IsAutoFilterEnabled = sheet.IsAutoFilterEnabled;
        Metadata = sheet.Metadata;
        DefaultFont = defaultFont;
    }

    /// <summary>
    /// Gets the name of the sheet.
    /// </summary>
    public string SheetName { get; }

    /// <summary>
    /// Gets the default row height, or <see langword="null"/> to use the renderer default.
    /// </summary>
    public double? DefaultRowHeight { get; }

    /// <summary>
    /// Gets the pre-calculated cells for this sheet.
    /// </summary>
    public IReadOnlyList<Cell> Cells { get; }

    /// <summary>
    /// Gets the column widths indexed by zero-based column index.
    /// </summary>
    public IReadOnlyDictionary<int, double> ColumnWidths { get; }

    /// <summary>
    /// Gets the row heights indexed by zero-based row index.
    /// </summary>
    public IReadOnlyDictionary<int, double?> RowHeights { get; }

    /// <summary>
    /// Gets the worksheet protection password.
    /// </summary>
    public string? Password { get; }

    /// <summary>
    /// Gets a value indicating whether this sheet is protected.
    /// </summary>
    public bool IsProtected => !string.IsNullOrEmpty(Password);

    /// <summary>
    /// Gets the page settings for this sheet.
    /// </summary>
    public PageSettings PageSettings { get; }

    /// <summary>
    /// Gets the freeze panes position (zero-based). <see langword="null"/> means no freeze.
    /// </summary>
    public Point? FreezePanes { get; }

    /// <summary>
    /// Gets a value indicating whether auto filter is enabled for the data range.
    /// </summary>
    public bool IsAutoFilterEnabled { get; }

    /// <summary>
    /// Gets the metadata dictionary from the sheet definition.
    /// </summary>
    public IDictionary<string, object?> Metadata { get; }

    /// <summary>
    /// Gets the default font for the workbook, or <see langword="null"/> to use the renderer default.
    /// </summary>
    public CellFont? DefaultFont { get; }

}
