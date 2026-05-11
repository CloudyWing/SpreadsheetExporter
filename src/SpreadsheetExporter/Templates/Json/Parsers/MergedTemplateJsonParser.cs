using System;
using System.Collections.Generic;
using System.Text.Json;

namespace CloudyWing.SpreadsheetExporter.Templates.Json.Parsers;

/// <summary>
/// Parses a JSON element into a <see cref="MergedTemplate"/>.
/// </summary>
public class MergedTemplateJsonParser : ITemplateJsonParser {
    /// <inheritdoc/>
    public string TypeName => "Merged";

    /// <inheritdoc/>
    public ISheetTemplate Parse(JsonElement element) {
        if (element.ValueKind != JsonValueKind.Object) {
            throw new FormatException("Merged template JSON must be an object.");
        }

        List<ISheetTemplate> templates = [];

        if (!element.TryGetPropertyIgnoreCase("Templates", out JsonElement templatesElement)) {
            return new MergedTemplate(templates);
        }

        if (templatesElement.ValueKind != JsonValueKind.Array) {
            throw new FormatException("Merged template property 'Templates' must be an array.");
        }

        foreach (JsonElement templateElement in templatesElement.EnumerateArray()) {
            templates.Add(CreateTemplate(templateElement));
        }

        return new MergedTemplate(templates);
    }

    private static ISheetTemplate CreateTemplate(JsonElement templateElement) {
        if (templateElement.ValueKind != JsonValueKind.Object) {
            throw new FormatException("Each merged template item must be a JSON object.");
        }

        if (!templateElement.TryGetPropertyIgnoreCase("Type", out JsonElement typeElement)) {
            throw new FormatException("Merged template item requires a 'Type' property.");
        }

        string typeName = typeElement.GetStringValue("Type");
        return JsonTemplateRegistry.Create(typeName, templateElement);
    }
}
