using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Drawing;
using System.Linq;
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
    public SheetLayout(SheetDefinition sheet, CellFont? defaultFont = null)
        : this(sheet, defaultFont, GetTemplateLayouts(sheet)) {
    }

    internal SheetLayout(SheetDefinition sheet, CellFont? defaultFont, IReadOnlyList<TemplateLayout> templateLayouts) {
        ArgumentNullException.ThrowIfNull(sheet);

        SheetName = sheet.SheetName;
        DefaultRowHeight = sheet.DefaultRowHeight;
        TemplateLayout templateLayout = TemplateLayout.Combine(templateLayouts);
        Cells = templateLayout.Cells;
        RowHeights = templateLayout.RowHeights.AsReadOnly();
        ColumnWidths = GetColumnWidths(sheet);
        Password = sheet.Password;
        PageSettings = sheet.PageSettings;
        FreezePanes = sheet.FreezePanes;
        IsAutoFilterEnabled = sheet.IsAutoFilterEnabled;
        AutoFilterRange = sheet.AutoFilterRange ?? GetDefaultAutoFilterRange(sheet.Templates, templateLayouts);
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
    /// Gets the zero-based range to apply the auto filter to.
    /// </summary>
    /// <value>
    /// <see cref="SheetDefinition.AutoFilterRange"/> when it is set. Otherwise, when the sheet contains more than one template,
    /// the full range of the first record set or data table template, so that rows such as a title above the table
    /// are excluded. <see langword="null"/> means the renderer uses the used range of the sheet.
    /// </value>
    public CellRange? AutoFilterRange { get; }

    /// <summary>
    /// Gets the metadata dictionary from the sheet definition.
    /// </summary>
    public IDictionary<string, object?> Metadata { get; }

    /// <summary>
    /// Gets the default font for the workbook, or <see langword="null"/> to use the renderer default.
    /// </summary>
    public CellFont? DefaultFont { get; }

    internal static IReadOnlyList<TemplateLayout> GetTemplateLayouts(SheetDefinition sheet) {
        ArgumentNullException.ThrowIfNull(sheet);
        return sheet.Templates
            .Select(template => template.GetLayout())
            .ToList();
    }

    private static IReadOnlyDictionary<int, double> GetColumnWidths(SheetDefinition sheet) {
        Dictionary<int, double>? templateWidths = null;
        foreach (ISheetTemplate template in sheet.Templates) {
            if (template is not IColumnWidthProvider provider) {
                continue;
            }

            foreach (KeyValuePair<int, double> pair in provider.GetColumnWidths()) {
                templateWidths ??= [];
                templateWidths.TryAdd(pair.Key, pair.Value);
            }
        }

        if (templateWidths is null) {
            return sheet.ColumnWidths.AsReadOnly();
        }

        // 工作表明確設定的欄寬優先於 template 欄寬
        foreach (KeyValuePair<int, double> pair in sheet.ColumnWidths) {
            templateWidths[pair.Key] = pair.Value;
        }

        return new ReadOnlyDictionary<int, double>(templateWidths);
    }

    private static CellRange? GetDefaultAutoFilterRange(
        IReadOnlyList<ISheetTemplate> templates, IReadOnlyList<TemplateLayout> templateLayouts
    ) {
        // 只有單一模板時沿用 renderer 的已使用範圍，維持既有輸出
        if (templates.Count <= 1) {
            return null;
        }

        int rowOffset = 0;
        for (int i = 0; i < templates.Count; i++) {
            TemplateLayout layout = templateLayouts[i];
            if (templates[i] is ITableTemplate table) {
                CellRange range = new(0, rowOffset, table.ColumnSpan, layout.RowSpan);
                return range.IsValid ? range : null;
            }

            rowOffset += layout.RowSpan;
        }

        return null;
    }
}
