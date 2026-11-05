using System.Drawing;

namespace CloudyWing.SpreadsheetExporter;

/// <summary>
/// The cell style.
/// </summary>
/// <param name="HorizontalAlignment"> Gets the horizontal alignment. </param>
/// <param name="VerticalAlignment"> Gets the vertical alignment. </param>
/// <param name="HasBorder"> Gets a value indicating whether this instance has border. </param>
/// <param name="WrapText"> Gets a value indicating whether [wrap text]. </param>
/// <param name="BackgroundColor"> Gets the color of the background. </param>
/// <param name="Font"> Gets the font. </param>
/// <param name="DataFormat"> Gets the data format. </param>
/// <param name="IsLocked"> Gets a value indicating whether this instance is locked. </param>
public record struct CellStyle(
    HorizontalAlignment HorizontalAlignment = HorizontalAlignment.General,
    VerticalAlignment VerticalAlignment = VerticalAlignment.Top,
    bool HasBorder = false, bool WrapText = false,
    Color BackgroundColor = default, CellFont Font = default,
    string? DataFormat = null,
    bool IsLocked = false
) {
    /// <summary>
    /// Gets the line style of the outside border.
    /// </summary>
    /// <value>
    /// The border line style. When it is not <see cref="BorderLineStyle.None"/>, an outside border with this style is drawn
    /// regardless of <see cref="HasBorder"/>. Default is <see cref="BorderLineStyle.None"/>.
    /// </value>
    public BorderLineStyle BorderStyle { get; init; }

    /// <summary>
    /// Gets an empty <see cref="CellStyle"/> with all default values.
    /// </summary>
    public static CellStyle Empty { get; } = new() {
        HorizontalAlignment = HorizontalAlignment.General,
        VerticalAlignment = VerticalAlignment.Top,
        HasBorder = false,
        WrapText = false,
        BackgroundColor = default,
        Font = default,
        DataFormat = null,
        IsLocked = false
    };
}
