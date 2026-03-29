namespace CloudyWing.SpreadsheetExporter.Config;

/// <summary>
/// Represents the default cell style configuration used across the spreadsheet.
/// </summary>
public sealed class CellStyleConfiguration {
    private const string DefaultFontName = "新細明體";
    private const short DefaultFontSize = 10;

    /// <summary>
    /// Initializes a new instance of the <see cref="CellStyleConfiguration"/> class with default styles.
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
            Style = cellStyle.Font.Style | FontStyles.Bold
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
    /// Gets the base cell style applied to general cells.
    /// </summary>
    public CellStyle CellStyle { get; init; }

    /// <summary>
    /// Gets the cell style applied to grid template cells.
    /// </summary>
    public CellStyle GridCellStyle { get; init; }

    /// <summary>
    /// Gets the cell style applied to column headers.
    /// </summary>
    public CellStyle HeaderStyle { get; init; }

    /// <summary>
    /// Gets the cell style applied to data fields.
    /// </summary>
    public CellStyle FieldStyle { get; init; }
}
