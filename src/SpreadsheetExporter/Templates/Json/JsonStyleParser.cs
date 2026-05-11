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

        return ParseStylePatch(element, context).Apply(CellStyle.Empty);
    }

    internal static CellStyle? ParseOptionalStyle(
        JsonElement ownerElement,
        string stylePropertyName,
        string styleNamePropertyName,
        JsonParseContext ownerContext
    ) {
        CellStyle? style = null;

        if (ownerElement.TryGetPropertyIgnoreCase(styleNamePropertyName, out JsonElement styleNameElement)
            && styleNameElement.ValueKind != JsonValueKind.Null
        ) {
            string styleName = styleNameElement.GetStringValue(ownerContext.Property(styleNamePropertyName));
            if (!ownerContext.StyleResolver.TryGet(styleName, out CellStyle namedStyle)) {
                throw JsonParseExceptionFactory.InvalidValue(
                    ownerContext.Property(styleNamePropertyName),
                    $"references unknown style '{styleName}'."
                );
            }

            style = namedStyle;
        }

        if (ownerElement.TryGetPropertyIgnoreCase(stylePropertyName, out JsonElement styleElement)
            && styleElement.ValueKind != JsonValueKind.Null
        ) {
            style = ParseStylePatch(styleElement, ownerContext.Property(stylePropertyName))
                .Apply(style ?? CellStyle.Empty);
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

        return ParseFontPatch(element, context).Apply(CellFont.Empty);
    }

    private static CellStylePatch ParseStylePatch(JsonElement element, JsonParseContext context) {
        if (element.ValueKind != JsonValueKind.Object) {
            throw JsonParseExceptionFactory.InvalidType(context, "a JSON object");
        }

        CellStylePatch patch = new();

        if (element.TryGetPropertyIgnoreCase(
            nameof(CellStyle.HorizontalAlignment), out JsonElement horizontalAlignmentElement
        )) {
            patch.HorizontalAlignment = ParseEnum<HorizontalAlignment>(
                horizontalAlignmentElement, context.Property(nameof(CellStyle.HorizontalAlignment))
            );
        }

        if (element.TryGetPropertyIgnoreCase(
            nameof(CellStyle.VerticalAlignment), out JsonElement verticalAlignmentElement
        )) {
            patch.VerticalAlignment = ParseEnum<VerticalAlignment>(
                verticalAlignmentElement, context.Property(nameof(CellStyle.VerticalAlignment))
            );
        }

        if (element.TryGetPropertyIgnoreCase(nameof(CellStyle.HasBorder), out JsonElement hasBorderElement)) {
            patch.HasBorder = hasBorderElement.GetBooleanValue(context.Property(nameof(CellStyle.HasBorder)));
        }

        if (element.TryGetPropertyIgnoreCase(nameof(CellStyle.WrapText), out JsonElement wrapTextElement)) {
            patch.WrapText = wrapTextElement.GetBooleanValue(context.Property(nameof(CellStyle.WrapText)));
        }

        if (element.TryGetPropertyIgnoreCase(
            nameof(CellStyle.BackgroundColor), out JsonElement backgroundColorElement
        )) {
            patch.BackgroundColor = ParseColor(
                backgroundColorElement, context.Property(nameof(CellStyle.BackgroundColor))
            );
        }

        if (element.TryGetPropertyIgnoreCase(nameof(CellStyle.Font), out JsonElement fontElement)) {
            patch.HasFont = true;
            patch.Font = fontElement.ValueKind == JsonValueKind.Null
                ? null
                : ParseFontPatch(fontElement, context.Property(nameof(CellStyle.Font)));
        }

        if (element.TryGetPropertyIgnoreCase(nameof(CellStyle.DataFormat), out JsonElement dataFormatElement)) {
            patch.HasDataFormat = true;
            patch.DataFormat = dataFormatElement.ValueKind == JsonValueKind.Null
                ? null
                : dataFormatElement.GetStringValue(context.Property(nameof(CellStyle.DataFormat)));
        }

        if (element.TryGetPropertyIgnoreCase(nameof(CellStyle.IsLocked), out JsonElement isLockedElement)) {
            patch.IsLocked = isLockedElement.GetBooleanValue(context.Property(nameof(CellStyle.IsLocked)));
        }

        return patch;
    }

    private static CellFontPatch ParseFontPatch(JsonElement element, JsonParseContext context) {
        if (element.ValueKind != JsonValueKind.Object) {
            throw JsonParseExceptionFactory.InvalidType(context, "a JSON object");
        }

        CellFontPatch patch = new();

        if (element.TryGetPropertyIgnoreCase(nameof(CellFont.Name), out JsonElement nameElement)) {
            patch.HasName = true;
            patch.Name = nameElement.ValueKind == JsonValueKind.Null
                ? null
                : nameElement.GetStringValue(context.Property(nameof(CellFont.Name)));
        }

        if (element.TryGetPropertyIgnoreCase(nameof(CellFont.Size), out JsonElement sizeElement)) {
            patch.Size = sizeElement.GetInt16Value(context.Property(nameof(CellFont.Size)));
        }

        if (element.TryGetPropertyIgnoreCase(nameof(CellFont.Color), out JsonElement colorElement)) {
            patch.Color = ParseColor(colorElement, context.Property(nameof(CellFont.Color)));
        }

        if (element.TryGetPropertyIgnoreCase(nameof(CellFont.Style), out JsonElement styleElement)) {
            patch.Style = ParseFontStyles(styleElement, context.Property(nameof(CellFont.Style)));
        }

        return patch;
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

    private sealed class CellStylePatch {
        public HorizontalAlignment? HorizontalAlignment { get; set; }

        public VerticalAlignment? VerticalAlignment { get; set; }

        public bool? HasBorder { get; set; }

        public bool? WrapText { get; set; }

        public Color? BackgroundColor { get; set; }

        public bool HasFont { get; set; }

        public CellFontPatch? Font { get; set; }

        public bool HasDataFormat { get; set; }

        public string? DataFormat { get; set; }

        public bool? IsLocked { get; set; }

        public CellStyle Apply(CellStyle baseStyle) {
            CellStyle style = baseStyle;

            if (HorizontalAlignment.HasValue) {
                style = style with { HorizontalAlignment = HorizontalAlignment.Value };
            }

            if (VerticalAlignment.HasValue) {
                style = style with { VerticalAlignment = VerticalAlignment.Value };
            }

            if (HasBorder.HasValue) {
                style = style with { HasBorder = HasBorder.Value };
            }

            if (WrapText.HasValue) {
                style = style with { WrapText = WrapText.Value };
            }

            if (BackgroundColor.HasValue) {
                style = style with { BackgroundColor = BackgroundColor.Value };
            }

            if (HasFont) {
                style = style with { Font = Font?.Apply(style.Font) ?? CellFont.Empty };
            }

            if (HasDataFormat) {
                style = style with { DataFormat = DataFormat };
            }

            if (IsLocked.HasValue) {
                style = style with { IsLocked = IsLocked.Value };
            }

            return style;
        }
    }

    private sealed class CellFontPatch {
        public bool HasName { get; set; }

        public string? Name { get; set; }

        public short? Size { get; set; }

        public Color? Color { get; set; }

        public FontStyles? Style { get; set; }

        public CellFont Apply(CellFont baseFont) {
            CellFont font = baseFont;

            if (HasName) {
                font = font with { Name = Name };
            }

            if (Size.HasValue) {
                font = font with { Size = Size.Value };
            }

            if (Color.HasValue) {
                font = font with { Color = Color.Value };
            }

            if (Style.HasValue) {
                font = font with { Style = Style.Value };
            }

            return font;
        }
    }
}
