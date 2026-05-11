namespace CloudyWing.SpreadsheetExporter.Templates.Json;

internal sealed class JsonStyleResolver {
    private readonly SpreadsheetStyleRegistry? documentStyles;
    private readonly SpreadsheetStyleRegistry? sheetStyles;

    internal JsonStyleResolver(SpreadsheetStyleRegistry? documentStyles, SpreadsheetStyleRegistry? sheetStyles) {
        this.documentStyles = documentStyles;
        this.sheetStyles = sheetStyles;
    }

    internal static JsonStyleResolver Empty { get; } = new(null, null);

    internal bool TryGet(string name, out CellStyle style) {
        if (sheetStyles?.TryGet(name, out style) == true) {
            return true;
        }

        if (documentStyles?.TryGet(name, out style) == true) {
            return true;
        }

        style = default;
        return false;
    }
}
