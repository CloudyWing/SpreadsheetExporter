using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text.Json;

namespace CloudyWing.SpreadsheetExporter.Templates.Json;

internal static class JsonElementExtensions {
    internal static bool TryGetPropertyIgnoreCase(
        this JsonElement element, string propertyName, out JsonElement value
    ) {
        if (element.ValueKind == JsonValueKind.Object) {
            foreach (JsonProperty property in element.EnumerateObject()) {
                if (string.Equals(property.Name, propertyName, StringComparison.OrdinalIgnoreCase)) {
                    value = property.Value;
                    return true;
                }
            }
        }

        value = default;
        return false;
    }

    internal static string GetStringValue(this JsonElement element, string propertyName) {
        return element.ValueKind == JsonValueKind.String
            ? element.GetString()!
            : throw new FormatException(
                $"Property '{propertyName}' must be a JSON string."
            );
    }

    internal static int GetInt32Value(this JsonElement element, string propertyName) {
        if (element.ValueKind == JsonValueKind.Number && element.TryGetInt32(out int intValue)) {
            return intValue;
        }

        if (element.ValueKind == JsonValueKind.String && int.TryParse(
            element.GetString()!, NumberStyles.Integer, CultureInfo.InvariantCulture, out int parsedValue
        )) {
            return parsedValue;
        }

        throw new FormatException($"Property '{propertyName}' must be a 32-bit integer.");
    }

    internal static short GetInt16Value(this JsonElement element, string propertyName) {
        if (element.ValueKind == JsonValueKind.Number && element.TryGetInt16(out short int16Value)) {
            return int16Value;
        }

        if (element.ValueKind == JsonValueKind.String && short.TryParse(
            element.GetString()!, NumberStyles.Integer, CultureInfo.InvariantCulture, out short parsedValue
        )) {
            return parsedValue;
        }

        throw new FormatException($"Property '{propertyName}' must be a 16-bit integer.");
    }

    internal static double GetDoubleValue(this JsonElement element, string propertyName) {
        if (element.ValueKind == JsonValueKind.Number) {
            return element.GetDouble();
        }

        if (element.ValueKind == JsonValueKind.String && double.TryParse(
            element.GetString()!, NumberStyles.Float | NumberStyles.AllowThousands,
            CultureInfo.InvariantCulture, out double parsedValue
        )) {
            return parsedValue;
        }

        throw new FormatException($"Property '{propertyName}' must be a number.");
    }

    internal static bool GetBooleanValue(this JsonElement element, string propertyName) {
        if (element.ValueKind == JsonValueKind.True) {
            return true;
        }

        if (element.ValueKind == JsonValueKind.False) {
            return false;
        }

        if (element.ValueKind == JsonValueKind.String && bool.TryParse(element.GetString()!, out bool parsedValue)) {
            return parsedValue;
        }

        throw new FormatException($"Property '{propertyName}' must be a Boolean.");
    }

    internal static object? ToObject(this JsonElement element) {
        return element.ValueKind switch {
            JsonValueKind.Object => element.ToDictionary(),
            JsonValueKind.Array => element.EnumerateArray().Select(ToObject).ToList(),
            JsonValueKind.String => element.GetString(),
            JsonValueKind.Number when element.TryGetInt32(out int intValue) => intValue,
            JsonValueKind.Number when element.TryGetInt64(out long longValue) => longValue,
            JsonValueKind.Number when element.TryGetDecimal(out decimal decimalValue) => decimalValue,
            JsonValueKind.Number => element.GetDouble(),
            JsonValueKind.True => true,
            JsonValueKind.False => false,
            JsonValueKind.Null => null,
            _ => null
        };
    }

    internal static Dictionary<string, object?> ToDictionary(this JsonElement element) {
        if (element.ValueKind != JsonValueKind.Object) {
            throw new FormatException("JSON element must be an object.");
        }

        Dictionary<string, object?> dictionary = new(StringComparer.Ordinal);
        foreach (JsonProperty property in element.EnumerateObject()) {
            dictionary[property.Name] = property.Value.ToObject();
        }

        return dictionary;
    }
}
