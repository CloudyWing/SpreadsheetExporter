using System;

namespace CloudyWing.SpreadsheetExporter.Templates.Json;

/// <summary>
/// Represents the JSON path context used by template parsers.
/// </summary>
public sealed class JsonParseContext {
    /// <summary>
    /// Initializes a new instance of the <see cref="JsonParseContext"/> class.
    /// </summary>
    /// <param name="path">The JSON path.</param>
    public JsonParseContext(string path) : this(path, JsonStyleResolver.Empty) {
    }

    private JsonParseContext(string path, JsonStyleResolver styleResolver) {
        Path = string.IsNullOrWhiteSpace(path) ? "$" : path;
        StyleResolver = styleResolver;
    }

    /// <summary>
    /// Gets the JSON root context.
    /// </summary>
    public static JsonParseContext Root { get; } = new("$");

    /// <summary>
    /// Gets the JSON path.
    /// </summary>
    public string Path { get; }

    internal JsonStyleResolver StyleResolver { get; }

    /// <summary>
    /// Creates a context for a JSON object property.
    /// </summary>
    /// <param name="name">The property name.</param>
    /// <returns>The property context.</returns>
    public JsonParseContext Property(string name) {
        ArgumentException.ThrowIfNullOrWhiteSpace(name);
        return new JsonParseContext($"{Path}.{name}", StyleResolver);
    }

    /// <summary>
    /// Creates a context for a JSON array item.
    /// </summary>
    /// <param name="index">The zero-based array index.</param>
    /// <returns>The array item context.</returns>
    public JsonParseContext Index(int index) {
        if (index < 0) {
            throw new ArgumentOutOfRangeException(nameof(index), index, "Index must be greater than or equal to 0.");
        }

        return new JsonParseContext($"{Path}[{index}]", StyleResolver);
    }

    internal JsonParseContext WithStyleResolver(JsonStyleResolver styleResolver) {
        return new JsonParseContext(Path, styleResolver);
    }

    /// <inheritdoc/>
    public override string ToString() => Path;
}
