using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;

namespace CloudyWing.SpreadsheetExporter.Extensions;

/// <summary>
/// Provides extension methods for <see cref="IEnumerable{T}"/>.
/// </summary>
internal static class EnumerableExtensions {
    /// <summary>
    /// Converts the specified <see cref="IEnumerable{T}"/> to a read-only list.
    /// </summary>
    /// <typeparam name="T">The type of the elements in the collection.</typeparam>
    /// <param name="enumerable">The enumerable to convert.</param>
    /// <returns>A read-only list containing the elements of the enumerable.</returns>
    public static IReadOnlyList<T> AsReadOnly<T>(this IEnumerable<T> enumerable) {
        if (enumerable is ReadOnlyCollection<T> readOnlyCollection) {
            return readOnlyCollection;
        }

        if (enumerable is List<T> list) {
            return list.AsReadOnly();
        }

        return enumerable.ToList().AsReadOnly();
    }
}
