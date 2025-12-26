namespace CloudyWing.SpreadsheetExporter;

/// <summary>
/// The page settings.
/// </summary>
public class PageSettings {
    /// <summary>
    /// Gets or sets the page orientation. The defaults is <c>Portrait</c>.
    /// </summary>
    /// <value>
    /// The page orientation.
    /// </value>
    public PageOrientation PageOrientation { get; set; } = PageOrientation.Portrait;

    /// <summary>
    /// Gets or sets the size of the pager.
    /// </summary>
    /// <value>
    /// The size of the pager. The defaults is <c>A4</c>.
    /// </value>
    public PaperSize PaperSize { get; set; } = PaperSize.A4;
}
