using System.Collections.Generic;
using System.Collections.ObjectModel;

namespace CloudyWing.SpreadsheetExporter.Util;

internal static class DictionaryExtensions {
    public static IReadOnlyDictionary<TKey, TValue> AsReadOnly<TKey, TValue>(
        this IReadOnlyDictionary<TKey, TValue> readOnlyDictionary
    ) where TKey : notnull {
        return readOnlyDictionary is IDictionary<TKey, TValue> dictionary
            ? new ReadOnlyDictionary<TKey, TValue>(dictionary)
            : readOnlyDictionary;
    }
}
