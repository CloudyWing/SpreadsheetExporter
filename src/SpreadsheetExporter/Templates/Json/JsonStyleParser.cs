using System;
using System.Drawing;
using System.Text.Json;

namespace CloudyWing.SpreadsheetExporter.Templates.Json;

internal static class JsonStyleParser {
    internal static CellStyle Parse(JsonElement element) {
        return Parse(element, JsonParseContext.Root);
    }

    internal static CellStyle Parse(JsonElement element, JsonParseContext context) {
        if (element.ValueKind == JsonValueKind.Null || element.ValueKind == JsonValueKind.Undefined) {
            return CellStyle.Empty;
        }

        if (element.ValueKind != JsonValueKind.Object) {
            throw JsonParseExceptionFactory.InvalidType(context, "a JSON object");
        }

        CellStyle style = CellStyle.Empty;

        if (element.TryGetPropertyIgnoreCase(
            nameof(CellStyle.HorizontalAlignment), out JsonElement horizontalAlignmentElement
        )) {
            style = style with {
                HorizontalAlignment = ParseEnum<HorizontalAlignment>(
                    horizontalAlignmentElement, context.Property(nameof(CellStyle.HorizontalAlignment))
                )
            };
        }

        if (element.TryGetPropertyIgnoreCase(
            nameof(CellStyle.VerticalAlignment), out JsonElement verticalAlignmentElement
        )) {
            style = style with {
                VerticalAlignment = ParseEnum<VerticalAlignment>(
                    verticalAlignmentElement, context.Property(nameof(CellStyle.VerticalAlignment))
                )
            };
        }

        if (element.TryGetPropertyIgnoreCase(nameof(CellStyle.HasBorder), out JsonElement hasBorderElement)) {
            style = style with {
                HasBorder = hasBorderElement.GetBooleanValue(context.Property(nameof(CellStyle.HasBorder)))
            };
        }

        if (element.TryGetPropertyIgnoreCase(nameof(CellStyle.WrapText), out JsonElement wrapTextElement)) {
            style = style with {
                WrapText = wrapTextElement.GetBooleanValue(context.Property(nameof(CellStyle.WrapText)))
            };
        }

        if (element.TryGetPropertyIgnoreCase(
            nameof(CellStyle.BackgroundColor), out JsonElement backgroundColorElement
        )) {
            style = style with {
                BackgroundColor = ParseColor(
                    backgroundColorElement, context.Property(nameof(CellStyle.BackgroundColor))
                )
            };
        }

        if (element.TryGetPropertyIgnoreCase(nameof(CellStyle.Font), out JsonElement fontElement)) {
            style = style with {
                Font = ParseFont(fontElement, context.Property(nameof(CellStyle.Font)))
            };
        }

        if (element.TryGetPropertyIgnoreCase(nameof(CellStyle.DataFormat), out JsonElement dataFormatElement)) {
            style = style with {
                DataFormat = dataFormatElement.ValueKind == JsonValueKind.Null
                    ? null
                    : dataFormatElement.GetStringValue(context.Property(nameof(CellStyle.DataFormat)))
            };
        }

        if (element.TryGetPropertyIgnoreCase(nameof(CellStyle.IsLocked), out JsonElement isLockedElement)) {
            style = style with {
                IsLocked = isLockedElement.GetBooleanValue(context.Property(nameof(CellStyle.IsLocked)))
            };
        }

        return style;
    }

    internal static CellFont ParseFont(JsonElement element) {
        return ParseFont(element, JsonParseContext.Root);
    }

    internal static CellFont ParseFont(JsonElement element, JsonParseContext context) {
        if (element.ValueKind == JsonValueKind.Null || element.ValueKind == JsonValueKind.Undefined) {
            return CellFont.Empty;
        }

        if (element.ValueKind != JsonValueKind.Object) {
            throw JsonParseExceptionFactory.InvalidType(context, "a JSON object");
        }

        CellFont font = CellFont.Empty;

        if (element.TryGetPropertyIgnoreCase(nameof(CellFont.Name), out JsonElement nameElement)) {
            font = font with {
                Name = nameElement.ValueKind == JsonValueKind.Null
                    ? null
                    : nameElement.GetStringValue(context.Property(nameof(CellFont.Name)))
            };
        }

        if (element.TryGetPropertyIgnoreCase(nameof(CellFont.Size), out JsonElement sizeElement)) {
            font = font with {
                Size = sizeElement.GetInt16Value(context.Property(nameof(CellFont.Size)))
            };
        }

        if (element.TryGetPropertyIgnoreCase(nameof(CellFont.Color), out JsonElement colorElement)) {
            font = font with {
                Color = ParseColor(colorElement, context.Property(nameof(CellFont.Color)))
            };
        }

        if (element.TryGetPropertyIgnoreCase(nameof(CellFont.Style), out JsonElement styleElement)) {
            font = font with {
                Style = ParseFontStyles(styleElement, context.Property(nameof(CellFont.Style)))
            };
        }

        return font;
    }

    private static TEnum ParseEnum<TEnum>(JsonElement element, JsonParseContext context)
        where TEnum : struct, Enum {
        if (element.ValueKind == JsonValueKind.String) {
            string? raw = element.GetString();
            if (Enum.TryParse(raw, true, out TEnum enumValue)) {
                return enumValue;
            }
        } else if (element.ValueKind == JsonValueKind.Number && element.TryGetInt32(out int numericValue)) {
            return (TEnum)Enum.ToObject(typeof(TEnum), numericValue);
        }

        throw JsonParseExceptionFactory.InvalidType(context, $"a valid {typeof(TEnum).Name} value");
    }

    private static FontStyles ParseFontStyles(JsonElement element, JsonParseContext context) {
        if (element.ValueKind == JsonValueKind.Number && element.TryGetInt32(out int parsedNumericValue)) {
            return (FontStyles)parsedNumericValue;
        }

        if (element.ValueKind == JsonValueKind.String) {
            return ParseSingleFontStyle(element.GetString()!, context);
        }

        if (element.ValueKind != JsonValueKind.Array) {
            throw JsonParseExceptionFactory.InvalidType(context, "a string, number, or array");
        }

        FontStyles styles = FontStyles.None;
        int index = 0;
        foreach (JsonElement item in element.EnumerateArray()) {
            JsonParseContext itemContext = context.Index(index++);
            styles |= item.ValueKind switch {
                JsonValueKind.String => ParseSingleFontStyle(item.GetString()!, itemContext),
                JsonValueKind.Number when item.TryGetInt32(out int itemNumericValue) => (FontStyles)itemNumericValue,
                _ => throw JsonParseExceptionFactory.InvalidType(itemContext, "a string or number")
            };
        }

        return styles;
    }

    private static FontStyles ParseSingleFontStyle(string? value, JsonParseContext context) {
        return Enum.TryParse(value, true, out FontStyles style)
            ? style
            : throw JsonParseExceptionFactory.InvalidValue(
                context, $"contains invalid {nameof(FontStyles)} value '{value}'."
            );
    }

    private static Color ParseColor(JsonElement element, JsonParseContext context) {
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
                throw JsonParseExceptionFactory.InvalidValue(context, $"contains invalid color value '{raw}'.");
            }

            throw JsonParseExceptionFactory.InvalidValue(context, $"contains invalid color value '{raw}'.");
        }

        if (element.ValueKind == JsonValueKind.Object) {
            int a = element.TryGetPropertyIgnoreCase("A", out JsonElement aElement)
                ? aElement.GetInt32Value(context.Property("A"))
                : byte.MaxValue;
            int r = element.TryGetPropertyIgnoreCase("R", out JsonElement rElement)
                ? rElement.GetInt32Value(context.Property("R"))
                : throw JsonParseExceptionFactory.MissingRequiredProperty(context.Property("R"));
            int g = element.TryGetPropertyIgnoreCase("G", out JsonElement gElement)
                ? gElement.GetInt32Value(context.Property("G"))
                : throw JsonParseExceptionFactory.MissingRequiredProperty(context.Property("G"));
            int b = element.TryGetPropertyIgnoreCase("B", out JsonElement bElement)
                ? bElement.GetInt32Value(context.Property("B"))
                : throw JsonParseExceptionFactory.MissingRequiredProperty(context.Property("B"));

            return Color.FromArgb(
                ValidateByte(a, context.Property("A")),
                ValidateByte(r, context.Property("R")),
                ValidateByte(g, context.Property("G")),
                ValidateByte(b, context.Property("B"))
            );
        }

        throw JsonParseExceptionFactory.InvalidType(context, "a color string or object");
    }

    private static byte ValidateByte(int value, JsonParseContext context) {
        return value is >= byte.MinValue and <= byte.MaxValue
            ? (byte)value
            : throw JsonParseExceptionFactory.InvalidValue(
                context, "must be between 0 and 255."
            );
    }
}
