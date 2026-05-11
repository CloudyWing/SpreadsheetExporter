using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;

namespace CloudyWing.SpreadsheetExporter;

/// <summary>
/// Stores named cell styles used by spreadsheet documents and sheets.
/// </summary>
public sealed class SpreadsheetStyleRegistry {
    private readonly Dictionary<string, CellStyle> styles = new(StringComparer.OrdinalIgnoreCase);
    private ReadOnlyDictionary<string, CellStyle>? readOnlyStyles;

    /// <summary>
    /// Gets the registered styles by name.
    /// </summary>
    public IReadOnlyDictionary<string, CellStyle> Styles
        => readOnlyStyles ??= new ReadOnlyDictionary<string, CellStyle>(styles);

    /// <summary>
    /// Sets a named style.
    /// </summary>
    /// <param name="name">The style name.</param>
    /// <param name="style">The style value.</param>
    /// <returns>The current <see cref="SpreadsheetStyleRegistry"/> instance.</returns>
    public SpreadsheetStyleRegistry Set(string name, CellStyle style) {
        ArgumentException.ThrowIfNullOrWhiteSpace(name);
        styles[name] = style;
        readOnlyStyles = null;
        return this;
    }

    /// <summary>
    /// Tries to get a named style.
    /// </summary>
    /// <param name="name">The style name.</param>
    /// <param name="style">The style value if found; otherwise, the default value.</param>
    /// <returns><see langword="true"/> when the style is found; otherwise, <see langword="false"/>.</returns>
    public bool TryGet(string name, out CellStyle style) {
        return styles.TryGetValue(name, out style);
    }

    /// <summary>
    /// Removes all named styles.
    /// </summary>
    public void Clear() {
        styles.Clear();
        readOnlyStyles = null;
    }

    /// <summary>
    /// Copies all styles from another registry into this registry.
    /// </summary>
    /// <param name="registry">The registry to copy from.</param>
    /// <returns>The current <see cref="SpreadsheetStyleRegistry"/> instance.</returns>
    public SpreadsheetStyleRegistry Import(SpreadsheetStyleRegistry registry) {
        ArgumentNullException.ThrowIfNull(registry);

        foreach (KeyValuePair<string, CellStyle> pair in registry.Styles) {
            Set(pair.Key, pair.Value);
        }

        return this;
    }
}
