using System;
using System.Collections.Generic;
using System.Text.Json;

namespace CloudyWing.SpreadsheetExporter.Templates.Json;

internal static class JsonDataValidationParser {
    internal static DataValidation? Parse(JsonElement element, string propertyName) {
        if (element.ValueKind == JsonValueKind.Null || element.ValueKind == JsonValueKind.Undefined) {
            return null;
        }

        if (element.ValueKind != JsonValueKind.Object) {
            throw new FormatException($"Property '{propertyName}' must be a JSON object.");
        }

        DataValidation validation = new();

        if (element.TryGetPropertyIgnoreCase(
            nameof(DataValidation.ValidationType), out JsonElement validationTypeElement
        )) {
            validation.ValidationType = ParseEnum<DataValidationType>(
                validationTypeElement, $"{propertyName}.{nameof(DataValidation.ValidationType)}"
            );
        }

        if (element.TryGetPropertyIgnoreCase(nameof(DataValidation.Operator), out JsonElement operatorElement)) {
            validation.Operator = operatorElement.ValueKind == JsonValueKind.Null
                ? null
                : ParseEnum<DataValidationOperator>(
                    operatorElement, $"{propertyName}.{nameof(DataValidation.Operator)}"
                );
        }

        if (element.TryGetPropertyIgnoreCase(nameof(DataValidation.Value1), out JsonElement value1Element)) {
            validation.Value1 = value1Element.ToObject();
        }

        if (element.TryGetPropertyIgnoreCase(nameof(DataValidation.Value2), out JsonElement value2Element)) {
            validation.Value2 = value2Element.ToObject();
        }

        if (element.TryGetPropertyIgnoreCase(nameof(DataValidation.ListItems), out JsonElement listItemsElement)) {
            validation.ListItems = ParseListItems(
                listItemsElement, $"{propertyName}.{nameof(DataValidation.ListItems)}"
            );
        }

        if (element.TryGetPropertyIgnoreCase(nameof(DataValidation.Formula), out JsonElement formulaElement)) {
            validation.Formula = GetNullableStringValue(
                formulaElement, $"{propertyName}.{nameof(DataValidation.Formula)}"
            );
        }

        if (element.TryGetPropertyIgnoreCase(
            nameof(DataValidation.IsDropdownShown), out JsonElement isDropdownShownElement
        )) {
            validation.IsDropdownShown = isDropdownShownElement.GetBooleanValue(
                $"{propertyName}.{nameof(DataValidation.IsDropdownShown)}"
            );
        }

        if (element.TryGetPropertyIgnoreCase(
            nameof(DataValidation.IsBlankAllowed), out JsonElement isBlankAllowedElement
        )) {
            validation.IsBlankAllowed = isBlankAllowedElement.GetBooleanValue(
                $"{propertyName}.{nameof(DataValidation.IsBlankAllowed)}"
            );
        }

        if (element.TryGetPropertyIgnoreCase(nameof(DataValidation.ErrorTitle), out JsonElement errorTitleElement)) {
            validation.ErrorTitle = GetNullableStringValue(
                errorTitleElement, $"{propertyName}.{nameof(DataValidation.ErrorTitle)}"
            );
        }

        if (element.TryGetPropertyIgnoreCase(
            nameof(DataValidation.ErrorMessage), out JsonElement errorMessageElement
        )) {
            validation.ErrorMessage = GetNullableStringValue(
                errorMessageElement, $"{propertyName}.{nameof(DataValidation.ErrorMessage)}"
            );
        }

        if (element.TryGetPropertyIgnoreCase(
            nameof(DataValidation.IsErrorAlertShown), out JsonElement isErrorAlertShownElement
        )) {
            validation.IsErrorAlertShown = isErrorAlertShownElement.GetBooleanValue(
                $"{propertyName}.{nameof(DataValidation.IsErrorAlertShown)}"
            );
        }

        if (element.TryGetPropertyIgnoreCase(nameof(DataValidation.PromptTitle), out JsonElement promptTitleElement)) {
            validation.PromptTitle = GetNullableStringValue(
                promptTitleElement, $"{propertyName}.{nameof(DataValidation.PromptTitle)}"
            );
        }

        if (element.TryGetPropertyIgnoreCase(
            nameof(DataValidation.PromptMessage), out JsonElement promptMessageElement
        )) {
            validation.PromptMessage = GetNullableStringValue(
                promptMessageElement, $"{propertyName}.{nameof(DataValidation.PromptMessage)}"
            );
        }

        if (element.TryGetPropertyIgnoreCase(
            nameof(DataValidation.IsInputPromptShown), out JsonElement isInputPromptShownElement
        )) {
            validation.IsInputPromptShown = isInputPromptShownElement.GetBooleanValue(
                $"{propertyName}.{nameof(DataValidation.IsInputPromptShown)}"
            );
        }

        return validation;
    }

    private static IReadOnlyList<string>? ParseListItems(JsonElement element, string propertyName) {
        if (element.ValueKind == JsonValueKind.Null) {
            return null;
        }

        if (element.ValueKind != JsonValueKind.Array) {
            throw new FormatException($"Property '{propertyName}' must be an array.");
        }

        List<string> items = [];
        foreach (JsonElement itemElement in element.EnumerateArray()) {
            items.Add(itemElement.GetStringValue(propertyName));
        }

        return items.AsReadOnly();
    }

    private static string? GetNullableStringValue(JsonElement element, string propertyName) {
        return element.ValueKind == JsonValueKind.Null
            ? null
            : element.GetStringValue(propertyName);
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
}
