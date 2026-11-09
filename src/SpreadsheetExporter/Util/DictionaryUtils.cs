using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Runtime.CompilerServices;

namespace CloudyWing.SpreadsheetExporter.Util;

internal static class DictionaryUtils {
    private static readonly ConditionalWeakTable<Type, PropertyInfo[]> ReadablePropertiesCache = new();

    internal static IDictionary<string, object?> ConvertFrom<T>(T record, int maxNestedLevel) {
        Dictionary<string, object?> dictionary = new(StringComparer.OrdinalIgnoreCase);
        AddPropertyToDictionary(dictionary, "", record, 1, maxNestedLevel);
        return dictionary;
    }

    /// <summary>
    /// Tries to get the value at the specified dot-separated property path.
    /// </summary>
    /// <param name="source">The source object.</param>
    /// <param name="path">The dot-separated property path. Property names are matched case-insensitively.</param>
    /// <param name="value">
    /// The resolved value, or <see langword="null"/> when <paramref name="source"/> or an intermediate object is
    /// <see langword="null"/>.
    /// </param>
    /// <param name="availablePaths">The property paths available at the level where resolution failed.</param>
    /// <returns><see langword="true"/> if the path is resolved; otherwise, <see langword="false"/>.</returns>
    /// <exception cref="ArgumentException">A path segment matches more than one property with different casing.</exception>
    internal static bool TryGetPathValue(
        object? source, string path, out object? value, out IReadOnlyList<string> availablePaths
    ) {
        ArgumentNullException.ThrowIfNull(path);

        object? current = source;
        string prefix = "";
        foreach (string segment in path.Split('.')) {
            if (current is null) {
                value = null;
                availablePaths = [];
                return true;
            }

            PropertyInfo[] properties = GetReadableProperties(current.GetType());
            PropertyInfo? property = FindProperty(properties, segment, path);
            if (property is null) {
                value = null;
                availablePaths = properties
                    .Select(x => prefix + x.Name)
                    .ToList();
                return false;
            }

            current = MemberAccessorCache.GetGetter(property)(current);
            prefix = $"{prefix}{property.Name}.";
        }

        value = current;
        availablePaths = [];
        return true;
    }

    private static PropertyInfo? FindProperty(PropertyInfo[] properties, string name, string path) {
        PropertyInfo? ignoreCaseMatch = null;
        bool isAmbiguous = false;

        foreach (PropertyInfo property in properties) {
            if (string.Equals(property.Name, name, StringComparison.Ordinal)) {
                return property;
            }

            if (string.Equals(property.Name, name, StringComparison.OrdinalIgnoreCase)) {
                isAmbiguous = ignoreCaseMatch is not null;
                ignoreCaseMatch = property;
            }
        }

        if (isAmbiguous) {
            throw new ArgumentException(
                $"Property path '{path}' is ambiguous: segment '{name}' matches multiple properties that differ only by case."
            );
        }

        return ignoreCaseMatch;
    }

    private static void AddPropertyToDictionary(
        IDictionary<string, object?> dictionary, string fullName, object? value, int level, int maxNestedLevel) {
        if (value is null) {
            dictionary.TryAdd(fullName, null);
            return;
        }

        Type type = value.GetType();
        Type underlyingType = Nullable.GetUnderlyingType(type) ?? type;

        if (level > maxNestedLevel || ShouldStoreDirectly(underlyingType)) {
            dictionary.TryAdd(fullName, value);
            return;
        }

        foreach (PropertyInfo prop in GetReadableProperties(underlyingType)) {
            object? propValue = MemberAccessorCache.GetGetter(prop)(value);

            string propFullName = string.IsNullOrEmpty(fullName)
                ? prop.Name
                : $"{fullName}.{prop.Name}";

            AddPropertyToDictionary(
                dictionary,
                propFullName,
                propValue,
                level + 1,
                maxNestedLevel
            );
        }
    }

    private static bool ShouldStoreDirectly(Type type) {
        return type.IsPrimitive
            || type == typeof(string)
            || type == typeof(Guid);
    }

    /// <summary>
    /// Gets the readable instance properties of the specified type, excluding indexers and by-ref properties.
    /// When a derived type hides a base property with <c>new</c>, only the most derived property is kept.
    /// </summary>
    private static PropertyInfo[] GetReadableProperties(Type type) {
        return ReadablePropertiesCache.GetValue(type, static t => t
            .GetProperties(BindingFlags.Public | BindingFlags.Instance)
            .Where(p => p.CanRead
                && p.GetIndexParameters().Length == 0
                && !p.PropertyType.IsByRef
                && !p.PropertyType.IsByRefLike
                && !p.PropertyType.IsPointer
            )
            .GroupBy(p => p.Name, StringComparer.Ordinal)
            .Select(g => g
                .OrderByDescending(p => GetInheritanceDepth(p.DeclaringType))
                .First()
            )
            .ToArray()
        );
    }

    private static int GetInheritanceDepth(Type? type) {
        int depth = 0;
        while (type?.BaseType is not null) {
            depth++;
            type = type.BaseType;
        }

        return depth;
    }
}
