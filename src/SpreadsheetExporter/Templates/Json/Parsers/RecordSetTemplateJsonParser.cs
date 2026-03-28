using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.Json;
using CloudyWing.SpreadsheetExporter.Templates.RecordSet;

namespace CloudyWing.SpreadsheetExporter.Templates.Json.Parsers;

/// <summary>
/// Parses a JSON element into a <see cref="RecordSetTemplate{T}"/> with <see cref="IDictionary{TKey,TValue}"/> records.
/// </summary>
public class RecordSetTemplateJsonParser : ITemplateJsonParser {
    /// <inheritdoc/>
    public string TypeName => "RecordSet";

    /// <inheritdoc/>
    public ISheetTemplate Parse(JsonElement element) {
        if (element.ValueKind != JsonValueKind.Object) {
            throw new FormatException("RecordSet template JSON must be an object.");
        }

        RecordSetTemplate<IDictionary<string, object?>> template = new([]);

        if (element.TryGetPropertyIgnoreCase(nameof(RecordSetTemplate<>.HeaderHeight), out JsonElement headerHeightElement)
            && headerHeightElement.ValueKind != JsonValueKind.Null
        ) {
            template.HeaderHeight = headerHeightElement.GetDoubleValue(nameof(RecordSetTemplate<>.HeaderHeight));
        }

        if (element.TryGetPropertyIgnoreCase(nameof(RecordSetTemplate<>.RecordHeight), out JsonElement recordHeightElement)
            && recordHeightElement.ValueKind != JsonValueKind.Null
        ) {
            template.RecordHeight = recordHeightElement.GetDoubleValue(nameof(RecordSetTemplate<>.RecordHeight));
        }

        template.Columns.Clear();
        if (element.TryGetPropertyIgnoreCase(nameof(RecordSetTemplate<>.Columns), out JsonElement columnsElement)) {
            if (columnsElement.ValueKind != JsonValueKind.Array) {
                throw new FormatException($"RecordSet template property '{nameof(RecordSetTemplate<>.Columns)}' must be an array.");
            }

            ParseColumns(columnsElement, template.Columns);
        }

        if (element.TryGetPropertyIgnoreCase("Records", out JsonElement recordsElement)) {
            if (recordsElement.ValueKind != JsonValueKind.Array) {
                throw new FormatException("RecordSet template property 'Records' must be an array.");
            }

            List<IDictionary<string, object?>> records = [];
            foreach (JsonElement recordElement in recordsElement.EnumerateArray()) {
                records.Add(recordElement.ToDictionary());
            }

            template.DataSource = records;
        }

        return template;
    }

    private static void ParseColumns(
        JsonElement columnsElement,
        RecordSetColumnCollection<IDictionary<string, object?>> collection
    ) {
        foreach (JsonElement columnElement in columnsElement.EnumerateArray()) {
            collection.Add(CreateColumn(columnElement));
        }
    }

    private static RecordSetColumnBase<IDictionary<string, object?>> CreateColumn(JsonElement columnElement) {
        if (columnElement.ValueKind != JsonValueKind.Object) {
            throw new FormatException("Each record-set column must be a JSON object.");
        }

        GeneratorColumn<IDictionary<string, object?>> column = new() {
            HeaderText = columnElement.TryGetPropertyIgnoreCase("HeaderText", out JsonElement headerTextElement)
                ? headerTextElement.GetStringValue("HeaderText")
                : null,
            HeaderStyle = columnElement.TryGetPropertyIgnoreCase("HeaderStyle", out JsonElement headerStyleElement)
                && headerStyleElement.ValueKind != JsonValueKind.Null
                ? JsonStyleParser.Parse(headerStyleElement)
                : SpreadsheetManager.DefaultCellStyles.HeaderStyle
        };

        CellStyle fieldStyle = columnElement.TryGetPropertyIgnoreCase("FieldStyle", out JsonElement fieldStyleElement)
            && fieldStyleElement.ValueKind != JsonValueKind.Null
            ? JsonStyleParser.Parse(fieldStyleElement)
            : SpreadsheetManager.DefaultCellStyles.FieldStyle;
        column.FieldStyleGenerator = context => fieldStyle;

        bool hasFieldKey = columnElement.TryGetPropertyIgnoreCase("FieldKey", out JsonElement fieldKeyElement);
        bool hasFormula = columnElement.TryGetPropertyIgnoreCase("Formula", out JsonElement formulaElement);
        bool hasValue = columnElement.TryGetPropertyIgnoreCase("Value", out JsonElement valueElement);

        if (hasFieldKey && hasValue) {
            throw new InvalidOperationException(
                "RecordSet column JSON cannot specify both 'FieldKey' and 'Value'."
            );
        }

        if (hasFieldKey && hasFormula) {
            throw new InvalidOperationException(
                "RecordSet column JSON cannot specify both 'FieldKey' and 'Formula'."
            );
        }

        if (hasValue && hasFormula) {
            throw new InvalidOperationException(
                "RecordSet column JSON cannot specify both 'Value' and 'Formula'."
            );
        }

        if (hasFieldKey) {
            string fieldKey = fieldKeyElement.GetStringValue("FieldKey");
            column.FieldValueGenerator = context => GetFieldValue(context.Record, fieldKey);
        } else if (hasFormula) {
            string formula = formulaElement.GetStringValue("Formula");
            column.FieldFormulaGenerator = context => formula;
        } else if (hasValue) {
            object? value = valueElement.ToObject();
            column.FieldValueGenerator = context => value;
        }

        if (columnElement.TryGetPropertyIgnoreCase("Children", out JsonElement childrenElement)) {
            if (childrenElement.ValueKind != JsonValueKind.Array) {
                throw new FormatException("RecordSet column property 'Children' must be an array.");
            }

            ParseColumns(childrenElement, column.ChildColumns);
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
