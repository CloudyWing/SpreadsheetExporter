using System;
using System.Collections.Generic;
using System.Text.Json;

namespace CloudyWing.SpreadsheetExporter.Templates.Json.Parsers;

/// <summary>
/// Parses a JSON element into a <see cref="MergedTemplate"/>.
/// </summary>
public class MergedTemplateJsonParser : ITemplateJsonParserWithContext {
    /// <inheritdoc/>
    public string TypeName => "Merged";

    /// <inheritdoc/>
    public ISheetTemplate Parse(JsonElement element) {
        return Parse(element, JsonParseContext.Root);
    }

    /// <inheritdoc/>
    public ISheetTemplate Parse(JsonElement element, JsonParseContext context) {
        if (element.ValueKind != JsonValueKind.Object) {
            throw JsonParseExceptionFactory.InvalidType(context, "a JSON object");
        }

        List<ISheetTemplate> templates = [];

        if (!element.TryGetPropertyIgnoreCase("Templates", out JsonElement templatesElement)) {
            return new MergedTemplate(templates);
        }

        JsonParseContext templatesContext = context.Property("Templates");
        if (templatesElement.ValueKind != JsonValueKind.Array) {
            throw JsonParseExceptionFactory.InvalidType(templatesContext, "an array");
        }

        int templateIndex = 0;
        foreach (JsonElement templateElement in templatesElement.EnumerateArray()) {
            templates.Add(CreateTemplate(templateElement, templatesContext.Index(templateIndex++)));
        }

        return new MergedTemplate(templates);
    }

    private static ISheetTemplate CreateTemplate(JsonElement templateElement, JsonParseContext templateContext) {
        if (templateElement.ValueKind != JsonValueKind.Object) {
            throw JsonParseExceptionFactory.InvalidType(templateContext, "a JSON object");
        }

        if (!templateElement.TryGetPropertyIgnoreCase("Type", out JsonElement typeElement)) {
            throw JsonParseExceptionFactory.MissingRequiredProperty(templateContext.Property("Type"));
        }

        string typeName = typeElement.GetStringValue(templateContext.Property("Type"));
        return JsonTemplateRegistry.Create(typeName, templateElement, templateContext);
    }
}
