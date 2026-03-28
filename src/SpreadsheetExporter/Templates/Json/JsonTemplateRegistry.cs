using System;
using System.Collections.Generic;
using System.Text.Json;
using CloudyWing.SpreadsheetExporter.Templates.Json.Parsers;

namespace CloudyWing.SpreadsheetExporter.Templates.Json;

/// <summary>
/// Manages the registration of <see cref="ITemplateJsonParser"/> instances
/// used by <see cref="SpreadsheetDocument.FromJson"/>.
/// </summary>
public static class JsonTemplateRegistry {
    private static readonly Dictionary<string, ITemplateJsonParser> Parsers
        = new(StringComparer.OrdinalIgnoreCase);

    static JsonTemplateRegistry() {
        RegisterBuiltins();
    }

    /// <summary>
    /// Registers a parser. If a parser with the same <see cref="ITemplateJsonParser.TypeName"/> already exists, it is replaced.
    /// </summary>
    /// <param name="parser">The parser to register.</param>
    /// <exception cref="ArgumentNullException"><paramref name="parser"/> is <see langword="null"/>.</exception>
    public static void Register(ITemplateJsonParser parser) {
        ArgumentNullException.ThrowIfNull(parser);
        Parsers[parser.TypeName] = parser;
    }

    /// <summary>
    /// Removes all registered parsers, including the built-ins.
    /// Call <see cref="RegisterBuiltins"/> to restore them.
    /// </summary>
    public static void Clear() {
        Parsers.Clear();
    }

    /// <summary>
    /// Registers all built-in parsers. Called automatically on first use.
    /// </summary>
    public static void RegisterBuiltins() {
        Register(new RecordSetTemplateJsonParser());
        Register(new GridTemplateJsonParser());
    }

    internal static ISheetTemplate Create(string typeName, JsonElement element) {
        if (!Parsers.TryGetValue(typeName, out ITemplateJsonParser? parser)) {
            throw new NotSupportedException(
                $"Template type '{typeName}' is not registered in {nameof(JsonTemplateRegistry)}. " +
                $"Call {nameof(JsonTemplateRegistry)}.{nameof(Register)}() to add a custom parser."
            );
        }
        return parser.Parse(element);
    }
}
