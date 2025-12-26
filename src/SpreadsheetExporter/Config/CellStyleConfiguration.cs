using System;

namespace CloudyWing.SpreadsheetExporter.Config;

/// <summary>
/// The spreadsheet default style.
/// </summary>
public class CellStyleConfiguration {
    /// <summary>
    /// Default font name used for cell styles.
    /// </summary>
    private const string DefaultFontName = "新細明體";

    /// <summary>
    /// Default font size used for cell styles.
    /// </summary>
    private const short DefaultFontSize = 10;
    /// <summary>
    /// Initializes a new instance of the <see cref="CellStyleConfiguration" /> class.
    /// </summary>
    /// <param name="loader">The loader.</param>
    public CellStyleConfiguration(Action<CellStyleSetuper> loader) {
        CellStyleSetuper setuper = new(this);
        loader(setuper);
    }

    /// <summary>
    /// Initializes a new instance of the <see cref="CellStyleConfiguration"/> class.
    /// <code>
    /// CellStyle: {
    ///     HorizontalAlignment: None,
    ///     VerticalAlignment: Middle,
    ///     HasBorder: false,
    ///     WrapText: false,
    ///     Font: {
    ///         Name: "新細明體",
    ///         Size: 10,
    ///         IsBold: false,
    ///         IsItalic: false,
    ///         HasUnderline: false,
    ///         IsStrikeout: false
    ///     }
    /// }
    ///
    /// GridCellStyle: {
    ///     StyleBase: CellStyle
    /// }
    ///
    /// HeaderStyle: {
    ///     StyleBase: CellStyle
    ///     HorizontalAlignment: Center,
    ///     HasBorder: true,
    ///     Font: {
    ///         IsBold: true,
    ///     }
    /// }
    ///
    /// FieldStyle: {
    ///     StyleBase: CellStyle
    ///     HasBorder: true
    /// }
    /// </code>
    /// </summary>
    public CellStyleConfiguration() {
        CellStyle cellStyle = new(
            HorizontalAlignment.Center,
            VerticalAlignment.Middle,
            false, false,
            default,
            new CellFont(DefaultFontName, DefaultFontSize, default, FontStyles.None),
            null,
            false
        );

        CellFont headerFont = cellStyle.Font with {
            Style = cellStyle.Font.Style | FontStyles.IsBold
        };

        CellStyle = cellStyle;
        GridCellStyle = cellStyle;
        HeaderStyle = cellStyle with {
            HorizontalAlignment = HorizontalAlignment.Center,
            HasBorder = true,
            Font = headerFont
        };
        FieldStyle = cellStyle with {
            HorizontalAlignment = HorizontalAlignment.Center,
            HasBorder = true
        };
    }

    /// <summary>
    /// Gets the cell style.
    /// </summary>
    /// <value>
    /// The cell style.
    /// </value>
    public virtual CellStyle CellStyle { get; internal set; }

    /// <summary>
    /// Gets the grid cell style.
    /// </summary>
    /// <value>
    /// The grid cell style.
    /// </value>
    public virtual CellStyle GridCellStyle { get; internal set; }

    /// <summary>
    /// Gets the header style.
    /// </summary>
    /// <value>
    /// The header style.
    /// </value>
    public virtual CellStyle HeaderStyle { get; internal set; }

    /// <summary>
    /// Gets the field style.
    /// </summary>
    /// <value>
    /// The field style.
    /// </value>
    public virtual CellStyle FieldStyle { get; internal set; }
}
