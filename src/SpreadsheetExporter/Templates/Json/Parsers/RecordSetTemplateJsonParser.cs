using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.Json;
using CloudyWing.SpreadsheetExporter.Templates.RecordSet;

namespace CloudyWing.SpreadsheetExporter.Templates.Json.Parsers;

/// <summary>
/// Parses a JSON element into a <see cref="RecordSetTemplate{T}"/> with <see cref="IDictionary{TKey,TValue}"/> records.
/// </summary>
public class RecordSetTemplateJsonParser : ITemplateJsonParserWithContext {
    /// <inheritdoc/>
    public string TypeName => "RecordSet";

    /// <inheritdoc/>
    public ISheetTemplate Parse(JsonElement element) {
        return Parse(element, JsonParseContext.Root);
    }

    /// <inheritdoc/>
    public ISheetTemplate Parse(JsonElement element, JsonParseContext context) {
        if (element.ValueKind != JsonValueKind.Object) {
            throw JsonParseExceptionFactory.InvalidType(context, "a JSON object");
        }

        RecordSetTemplate<IDictionary<string, object?>> template = new([]);

        if (element.TryGetPropertyIgnoreCase(
            nameof(RecordSetTemplate<>.HeaderHeight), out JsonElement headerHeightElement
        )
            && headerHeightElement.ValueKind != JsonValueKind.Null
        ) {
            template.HeaderHeight = headerHeightElement.GetDoubleValue(
                context.Property(nameof(RecordSetTemplate<>.HeaderHeight))
            );
        }

        if (element.TryGetPropertyIgnoreCase(
            nameof(RecordSetTemplate<>.RecordHeight), out JsonElement recordHeightElement
        )
            && recordHeightElement.ValueKind != JsonValueKind.Null
        ) {
            template.RecordHeight = recordHeightElement.GetDoubleValue(
                context.Property(nameof(RecordSetTemplate<>.RecordHeight))
            );
        }

        template.Columns.Clear();
        if (element.TryGetPropertyIgnoreCase(nameof(RecordSetTemplate<>.Columns), out JsonElement columnsElement)) {
            JsonParseContext columnsContext = context.Property(nameof(RecordSetTemplate<>.Columns));
            if (columnsElement.ValueKind != JsonValueKind.Array) {
                throw JsonParseExceptionFactory.InvalidType(columnsContext, "an array");
            }

            ParseColumns(columnsElement, template.Columns, columnsContext);
        }

        if (element.TryGetPropertyIgnoreCase("Records", out JsonElement recordsElement)) {
            JsonParseContext recordsContext = context.Property("Records");
            if (recordsElement.ValueKind != JsonValueKind.Array) {
                throw JsonParseExceptionFactory.InvalidType(recordsContext, "an array");
            }

            List<IDictionary<string, object?>> records = [];
            int recordIndex = 0;
            foreach (JsonElement recordElement in recordsElement.EnumerateArray()) {
                records.Add(recordElement.ToDictionary(recordsContext.Index(recordIndex++)));
            }

            template.DataSource = records;
        }

        return template;
    }

    private static void ParseColumns(
        JsonElement columnsElement,
        RecordSetColumnCollection<IDictionary<string, object?>> collection,
        JsonParseContext columnsContext
    ) {
        int columnIndex = 0;
        foreach (JsonElement columnElement in columnsElement.EnumerateArray()) {
            collection.Add(CreateColumn(columnElement, columnsContext.Index(columnIndex++)));
        }
    }

    private static RecordSetColumnBase<IDictionary<string, object?>> CreateColumn(
        JsonElement columnElement, JsonParseContext columnContext
    ) {
        if (columnElement.ValueKind != JsonValueKind.Object) {
            throw JsonParseExceptionFactory.InvalidType(columnContext, "a JSON object");
        }

        CellStyle? headerStyle = JsonStyleParser.ParseOptionalStyle(
            columnElement, "HeaderStyle", "HeaderStyleName", columnContext
        );
        CellStyle? fieldStyle = JsonStyleParser.ParseOptionalStyle(
            columnElement, "FieldStyle", "FieldStyleName", columnContext
        );

        GeneratorColumn<IDictionary<string, object?>> column = new() {
            HeaderText = columnElement.TryGetPropertyIgnoreCase("HeaderText", out JsonElement headerTextElement)
                ? headerTextElement.GetStringValue(columnContext.Property("HeaderText"))
                : null,
            HeaderStyle = headerStyle ?? SpreadsheetManager.DefaultCellStyles.HeaderStyle
        };

        CellStyle resolvedFieldStyle = fieldStyle ?? SpreadsheetManager.DefaultCellStyles.FieldStyle;
        column.FieldStyleGenerator = context => resolvedFieldStyle;

        DataValidation? dataValidation = columnElement.TryGetPropertyIgnoreCase(
            "DataValidation", out JsonElement dataValidationElement
        )
            ? JsonDataValidationParser.Parse(dataValidationElement, columnContext.Property("DataValidation"))
            : null;
        if (dataValidation is not null) {
            column.FieldDataValidationGenerator = context => dataValidation;
        }

        bool hasFieldKey = columnElement.TryGetPropertyIgnoreCase("FieldKey", out JsonElement fieldKeyElement);
        bool hasFormula = columnElement.TryGetPropertyIgnoreCase("Formula", out JsonElement formulaElement);
        bool hasValue = columnElement.TryGetPropertyIgnoreCase("Value", out JsonElement valueElement);

        if (hasFieldKey && hasValue) {
            throw JsonParseExceptionFactory.InvalidOperation(
                columnContext, "cannot specify both 'FieldKey' and 'Value'."
            );
        }

        if (hasFieldKey && hasFormula) {
            throw JsonParseExceptionFactory.InvalidOperation(
                columnContext, "cannot specify both 'FieldKey' and 'Formula'."
            );
        }

        if (hasValue && hasFormula) {
            throw JsonParseExceptionFactory.InvalidOperation(
                columnContext, "cannot specify both 'Value' and 'Formula'."
            );
        }

        if (hasFieldKey) {
            string fieldKey = fieldKeyElement.GetStringValue(columnContext.Property("FieldKey"));
            column.FieldValueGenerator = context => GetFieldValue(context.Record, fieldKey);
        } else if (hasFormula) {
            string formula = formulaElement.GetStringValue(columnContext.Property("Formula"));
            column.FieldFormulaGenerator = context => formula;
        } else if (hasValue) {
            object? value = valueElement.ToObject(columnContext.Property("Value"));
            column.FieldValueGenerator = context => value;
        }

        if (columnElement.TryGetPropertyIgnoreCase("Children", out JsonElement childrenElement)) {
            JsonParseContext childrenContext = columnContext.Property("Children");
            if (childrenElement.ValueKind != JsonValueKind.Array) {
                throw JsonParseExceptionFactory.InvalidType(childrenContext, "an array");
            }

            ParseColumns(childrenElement, column.ChildColumns, childrenContext);
        }

        return column;
    }

    private static object? GetFieldValue(IDictionary<string, object?> record, string fieldKey) {
        if (TryGetValueByPath(record, fieldKey, out object? value)) {
            return value;
        }

        string availableProperties = string.Join(
            ", ",
            GetAvailableKeys(record).OrderBy(key => key, StringComparer.Ordinal)
        );

        throw new ArgumentException(
            $"Data source does not contain property '{fieldKey}'. Available properties: {availableProperties}"
        );
    }

    private static bool TryGetValueByPath(
        IDictionary<string, object?> dictionary, string fieldKey, out object? value
    ) {
        if (dictionary.TryGetValue(fieldKey, out value)) {
            return true;
        }

        object? current = dictionary;
        foreach (string segment in fieldKey.Split('.')) {
            if (current is IDictionary<string, object?> currentDictionary
                && currentDictionary.TryGetValue(segment, out current)) {
                continue;
            }

            value = null;
            return false;
        }

        value = current;
        return true;
    }

    private static IEnumerable<string> GetAvailableKeys(IDictionary<string, object?> dictionary) {
        foreach (KeyValuePair<string, object?> pair in dictionary) {
            yield return pair.Key;

            if (pair.Value is IDictionary<string, object?> nestedDictionary) {
                foreach (string nestedKey in GetAvailableKeys(nestedDictionary)) {
                    yield return $"{pair.Key}.{nestedKey}";
                }
            }
        }
    }
}
