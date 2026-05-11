using System;
using System.Collections.Generic;
using System.Text.Json;

namespace CloudyWing.SpreadsheetExporter.Templates.Json;

internal static class JsonDataValidationParser {
    internal static DataValidation? Parse(JsonElement element, string propertyName) {
        return Parse(element, new JsonParseContext(propertyName));
    }

    internal static DataValidation? Parse(JsonElement element, JsonParseContext context) {
        if (element.ValueKind == JsonValueKind.Null || element.ValueKind == JsonValueKind.Undefined) {
            return null;
        }

        if (element.ValueKind != JsonValueKind.Object) {
            throw JsonParseExceptionFactory.InvalidType(context, "a JSON object");
        }

        DataValidation validation = new();

        if (element.TryGetPropertyIgnoreCase(
            nameof(DataValidation.ValidationType), out JsonElement validationTypeElement
        )) {
            validation.ValidationType = ParseEnum<DataValidationType>(
                validationTypeElement, context.Property(nameof(DataValidation.ValidationType))
            );
        }

        if (element.TryGetPropertyIgnoreCase(nameof(DataValidation.Operator), out JsonElement operatorElement)) {
            validation.Operator = operatorElement.ValueKind == JsonValueKind.Null
                ? null
                : ParseEnum<DataValidationOperator>(
                    operatorElement, context.Property(nameof(DataValidation.Operator))
                );
        }

        if (element.TryGetPropertyIgnoreCase(nameof(DataValidation.Value1), out JsonElement value1Element)) {
            validation.Value1 = value1Element.ToObject(context.Property(nameof(DataValidation.Value1)));
        }

        if (element.TryGetPropertyIgnoreCase(nameof(DataValidation.Value2), out JsonElement value2Element)) {
            validation.Value2 = value2Element.ToObject(context.Property(nameof(DataValidation.Value2)));
        }

        if (element.TryGetPropertyIgnoreCase(nameof(DataValidation.ListItems), out JsonElement listItemsElement)) {
            validation.ListItems = ParseListItems(
                listItemsElement, context.Property(nameof(DataValidation.ListItems))
            );
        }

        if (element.TryGetPropertyIgnoreCase(nameof(DataValidation.Formula), out JsonElement formulaElement)) {
            validation.Formula = GetNullableStringValue(
                formulaElement, context.Property(nameof(DataValidation.Formula))
            );
        }

        if (element.TryGetPropertyIgnoreCase(
            nameof(DataValidation.IsDropdownShown), out JsonElement isDropdownShownElement
        )) {
            validation.IsDropdownShown = isDropdownShownElement.GetBooleanValue(
                context.Property(nameof(DataValidation.IsDropdownShown))
            );
        }

        if (element.TryGetPropertyIgnoreCase(
            nameof(DataValidation.IsBlankAllowed), out JsonElement isBlankAllowedElement
        )) {
            validation.IsBlankAllowed = isBlankAllowedElement.GetBooleanValue(
                context.Property(nameof(DataValidation.IsBlankAllowed))
            );
        }

        if (element.TryGetPropertyIgnoreCase(nameof(DataValidation.ErrorTitle), out JsonElement errorTitleElement)) {
            validation.ErrorTitle = GetNullableStringValue(
                errorTitleElement, context.Property(nameof(DataValidation.ErrorTitle))
            );
        }

        if (element.TryGetPropertyIgnoreCase(
            nameof(DataValidation.ErrorMessage), out JsonElement errorMessageElement
        )) {
            validation.ErrorMessage = GetNullableStringValue(
                errorMessageElement, context.Property(nameof(DataValidation.ErrorMessage))
            );
        }

        if (element.TryGetPropertyIgnoreCase(
            nameof(DataValidation.IsErrorAlertShown), out JsonElement isErrorAlertShownElement
        )) {
            validation.IsErrorAlertShown = isErrorAlertShownElement.GetBooleanValue(
                context.Property(nameof(DataValidation.IsErrorAlertShown))
            );
        }

        if (element.TryGetPropertyIgnoreCase(nameof(DataValidation.PromptTitle), out JsonElement promptTitleElement)) {
            validation.PromptTitle = GetNullableStringValue(
                promptTitleElement, context.Property(nameof(DataValidation.PromptTitle))
            );
        }

        if (element.TryGetPropertyIgnoreCase(
            nameof(DataValidation.PromptMessage), out JsonElement promptMessageElement
        )) {
            validation.PromptMessage = GetNullableStringValue(
                promptMessageElement, context.Property(nameof(DataValidation.PromptMessage))
            );
        }

        if (element.TryGetPropertyIgnoreCase(
            nameof(DataValidation.IsInputPromptShown), out JsonElement isInputPromptShownElement
        )) {
            validation.IsInputPromptShown = isInputPromptShownElement.GetBooleanValue(
                context.Property(nameof(DataValidation.IsInputPromptShown))
            );
        }

        return validation;
    }

    private static IReadOnlyList<string>? ParseListItems(JsonElement element, JsonParseContext context) {
        if (element.ValueKind == JsonValueKind.Null) {
            return null;
        }

        if (element.ValueKind != JsonValueKind.Array) {
            throw JsonParseExceptionFactory.InvalidType(context, "an array");
        }

        List<string> items = [];
        int index = 0;
        foreach (JsonElement itemElement in element.EnumerateArray()) {
            items.Add(itemElement.GetStringValue(context.Index(index++)));
        }

        return items.AsReadOnly();
    }

    private static string? GetNullableStringValue(JsonElement element, JsonParseContext context) {
        return element.ValueKind == JsonValueKind.Null
            ? null
            : element.GetStringValue(context);
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
}
