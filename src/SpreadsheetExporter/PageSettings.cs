namespace CloudyWing.SpreadsheetExporter;

/// <summary>
/// The page settings.
/// </summary>
public class PageSettings {
    /// <summary>
    /// Gets or sets the page orientation. The default is <c>Portrait</c>.
    /// </summary>
    /// <value>
    /// The page orientation.
    /// </value>
    public PageOrientation PageOrientation { get; set; } = PageOrientation.Portrait;

    /// <summary>
    /// Gets or sets the paper size.
    /// </summary>
    /// <value>
    /// The paper size. The default is <c>A4</c>.
    /// </value>
    public PaperSize PaperSize { get; set; } = PaperSize.A4;
}
