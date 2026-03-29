using System.Text.Json;

namespace CloudyWing.SpreadsheetExporter.Templates.Json;

/// <summary>
/// Defines a parser that creates a <see cref="ISheetTemplate"/> from a JSON element.
/// </summary>
public interface ITemplateJsonParser {
    /// <summary>
    /// Gets the type name used to match the JSON "Type" property.
    /// </summary>
    string TypeName { get; }

    /// <summary>
    /// Parses the specified JSON element and returns a configured <see cref="ISheetTemplate"/>.
    /// </summary>
    /// <param name="element">The JSON element representing the template definition.</param>
    /// <returns>A configured <see cref="ISheetTemplate"/>.</returns>
    ISheetTemplate Parse(JsonElement element);
}
