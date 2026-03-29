using System;
using System.Drawing;
using System.Text.Json;

namespace CloudyWing.SpreadsheetExporter.Templates.Json;

internal static class JsonStyleParser {
    internal static CellStyle Parse(JsonElement element) {
        if (element.ValueKind == JsonValueKind.Null || element.ValueKind == JsonValueKind.Undefined) {
            return CellStyle.Empty;
        }

        if (element.ValueKind != JsonValueKind.Object) {
            throw new FormatException("Style must be a JSON object.");
        }

        CellStyle style = CellStyle.Empty;

        if (element.TryGetPropertyIgnoreCase(nameof(CellStyle.HorizontalAlignment), out JsonElement horizontalAlignmentElement)) {
            style = style with {
                HorizontalAlignment = ParseEnum<HorizontalAlignment>(
                    horizontalAlignmentElement, nameof(CellStyle.HorizontalAlignment)
                )
            };
        }

        if (element.TryGetPropertyIgnoreCase(nameof(CellStyle.VerticalAlignment), out JsonElement verticalAlignmentElement)) {
            style = style with {
                VerticalAlignment = ParseEnum<VerticalAlignment>(
                    verticalAlignmentElement, nameof(CellStyle.VerticalAlignment)
                )
            };
        }

        if (element.TryGetPropertyIgnoreCase(nameof(CellStyle.HasBorder), out JsonElement hasBorderElement)) {
            style = style with {
                HasBorder = hasBorderElement.GetBooleanValue(nameof(CellStyle.HasBorder))
            };
        }

        if (element.TryGetPropertyIgnoreCase(nameof(CellStyle.WrapText), out JsonElement wrapTextElement)) {
            style = style with {
                WrapText = wrapTextElement.GetBooleanValue(nameof(CellStyle.WrapText))
            };
        }

        if (element.TryGetPropertyIgnoreCase(nameof(CellStyle.BackgroundColor), out JsonElement backgroundColorElement)) {
            style = style with {
                BackgroundColor = ParseColor(backgroundColorElement, nameof(CellStyle.BackgroundColor))
            };
        }

        if (element.TryGetPropertyIgnoreCase(nameof(CellStyle.Font), out JsonElement fontElement)) {
            style = style with {
                Font = ParseFont(fontElement)
            };
        }

        if (element.TryGetPropertyIgnoreCase(nameof(CellStyle.DataFormat), out JsonElement dataFormatElement)) {
            style = style with {
                DataFormat = dataFormatElement.ValueKind == JsonValueKind.Null
                    ? null
                    : dataFormatElement.GetStringValue(nameof(CellStyle.DataFormat))
            };
        }

        if (element.TryGetPropertyIgnoreCase(nameof(CellStyle.IsLocked), out JsonElement isLockedElement)) {
            style = style with {
                IsLocked = isLockedElement.GetBooleanValue(nameof(CellStyle.IsLocked))
            };
        }

        return style;
    }

    internal static CellFont ParseFont(JsonElement element) {
        if (element.ValueKind == JsonValueKind.Null || element.ValueKind == JsonValueKind.Undefined) {
            return CellFont.Empty;
        }

        if (element.ValueKind != JsonValueKind.Object) {
            throw new FormatException("Font must be a JSON object.");
        }

        CellFont font = CellFont.Empty;

        if (element.TryGetPropertyIgnoreCase(nameof(CellFont.Name), out JsonElement nameElement)) {
            font = font with {
                Name = nameElement.ValueKind == JsonValueKind.Null
                    ? null
                    : nameElement.GetStringValue(nameof(CellFont.Name))
            };
        }

        if (element.TryGetPropertyIgnoreCase(nameof(CellFont.Size), out JsonElement sizeElement)) {
            font = font with {
                Size = sizeElement.GetInt16Value(nameof(CellFont.Size))
            };
        }

        if (element.TryGetPropertyIgnoreCase(nameof(CellFont.Color), out JsonElement colorElement)) {
            font = font with {
                Color = ParseColor(colorElement, nameof(CellFont.Color))
            };
        }

        if (element.TryGetPropertyIgnoreCase(nameof(CellFont.Style), out JsonElement styleElement)) {
            font = font with {
                Style = ParseFontStyles(styleElement)
            };
        }

        return font;
    }

    private static TEnum ParseEnum<TEnum>(JsonElement element, string propertyName)
        where TEnum : struct, Enum {
        if (element.ValueKind == JsonValueKind.String) {
            string? raw = element.GetString();
            if (Enum.TryParse(raw, true, out TEnum enumValue)) {
                return enumValue;
            }
        } else if (element.ValueKind == JsonValueKind.Number && element.TryGetInt32(out int numericValue)) {
            return (TEnum)Enum.ToObject(typeof(TEnum), numericValue);
        }

        throw new FormatException(
            $"Property '{propertyName}' must be a valid {typeof(TEnum).Name} value."
        );
    }

    private static FontStyles ParseFontStyles(JsonElement element) {
        if (element.ValueKind == JsonValueKind.Number && element.TryGetInt32(out int parsedNumericValue)) {
            return (FontStyles)parsedNumericValue;
        }

        if (element.ValueKind == JsonValueKind.String) {
            return ParseSingleFontStyle(element.GetString()!);
        }

        if (element.ValueKind != JsonValueKind.Array) {
            throw new FormatException("Font.Style must be a string, number, or array.");
        }

        FontStyles styles = FontStyles.None;
        foreach (JsonElement item in element.EnumerateArray()) {
            styles |= item.ValueKind switch {
                JsonValueKind.String => ParseSingleFontStyle(item.GetString()!),
                JsonValueKind.Number when item.TryGetInt32(out int itemNumericValue) => (FontStyles)itemNumericValue,
                _ => throw new FormatException("Each Font.Style item must be a string or number.")
            };
        }

        return styles;
    }

    private static FontStyles ParseSingleFontStyle(string? value) {
        return Enum.TryParse(value, true, out FontStyles style)
            ? style
            : throw new FormatException($"'{value}' is not a valid {nameof(FontStyles)} value.");
    }

    private static Color ParseColor(JsonElement element, string propertyName) {
        if (element.ValueKind == JsonValueKind.String) {
            string? raw = element.GetString();
            if (string.IsNullOrWhiteSpace(raw)) {
                return default;
            }

            try {
                if (raw.StartsWith('#')) {
                    return ColorTranslator.FromHtml(raw);
                }

                Color namedColor = Color.FromName(raw);
                if (namedColor.IsKnownColor || namedColor.IsNamedColor || namedColor.IsSystemColor) {
                    return namedColor;
                }
            } catch (Exception ex) when (ex is ArgumentException or FormatException) {
                throw new FormatException(
                    $"Property '{propertyName}' contains an invalid color value '{raw}'.", ex
                );
            }

            throw new FormatException(
                $"Property '{propertyName}' contains an invalid color value '{raw}'."
            );
        }

        if (element.ValueKind == JsonValueKind.Object) {
            int a = element.TryGetPropertyIgnoreCase("A", out JsonElement aElement)
                ? aElement.GetInt32Value($"{propertyName}.A")
                : byte.MaxValue;
            int r = element.TryGetPropertyIgnoreCase("R", out JsonElement rElement)
                ? rElement.GetInt32Value($"{propertyName}.R")
                : throw new FormatException($"Property '{propertyName}.R' is required.");
            int g = element.TryGetPropertyIgnoreCase("G", out JsonElement gElement)
                ? gElement.GetInt32Value($"{propertyName}.G")
                : throw new FormatException($"Property '{propertyName}.G' is required.");
            int b = element.TryGetPropertyIgnoreCase("B", out JsonElement bElement)
                ? bElement.GetInt32Value($"{propertyName}.B")
                : throw new FormatException($"Property '{propertyName}.B' is required.");

            return Color.FromArgb(
                ValidateByte(a, $"{propertyName}.A"),
                ValidateByte(r, $"{propertyName}.R"),
                ValidateByte(g, $"{propertyName}.G"),
                ValidateByte(b, $"{propertyName}.B")
            );
        }

        throw new FormatException(
            $"Property '{propertyName}' must be a color string or object."
        );
    }

    private static byte ValidateByte(int value, string propertyName) {
        return value is >= byte.MinValue and <= byte.MaxValue
            ? (byte)value
            : throw new FormatException(
                $"Property '{propertyName}' must be between 0 and 255."
            );
    }
}
