using System.Text.Json;

namespace CloudyWing.SpreadsheetExporter.Templates.Json;

/// <summary>
/// Defines a parser that can use JSON path context while creating a sheet template.
/// </summary>
public interface ITemplateJsonParserWithContext : ITemplateJsonParser {
    /// <summary>
    /// Parses the specified JSON element and returns a configured <see cref="ISheetTemplate"/>.
    /// </summary>
    /// <param name="element">The JSON element representing the template definition.</param>
    /// <param name="context">The JSON path context for diagnostics.</param>
    /// <returns>A configured <see cref="ISheetTemplate"/>.</returns>
    ISheetTemplate Parse(JsonElement element, JsonParseContext context);
}
