using System;
using System.Collections.Generic;
using System.Data;
using System.Text.Json;
using CloudyWing.SpreadsheetExporter.Templates.DataTable;

namespace CloudyWing.SpreadsheetExporter.Templates.Json.Parsers;

/// <summary>
/// Parses a JSON element into a <see cref="DataTableTemplate"/>.
/// </summary>
public class DataTableTemplateJsonParser : ITemplateJsonParser {
    /// <inheritdoc/>
    public string TypeName => "DataTable";

    /// <inheritdoc/>
    public ISheetTemplate Parse(JsonElement element) {
        if (element.ValueKind != JsonValueKind.Object) {
            throw new FormatException("DataTable template JSON must be an object.");
        }

        IReadOnlyList<DataTableColumnDefinition> columnDefinitions = ParseColumnDefinitions(element);
        IReadOnlyList<IDictionary<string, object?>> records = ParseRecords(element);
        System.Data.DataTable dataTable = CreateDataTable(columnDefinitions, records);
        DataTableTemplate template = new(dataTable);

        if (element.TryGetPropertyIgnoreCase(
            nameof(DataTableTemplate.HeaderHeight), out JsonElement headerHeightElement
        )
            && headerHeightElement.ValueKind != JsonValueKind.Null
        ) {
            template.HeaderHeight = headerHeightElement.GetDoubleValue(nameof(DataTableTemplate.HeaderHeight));
        }

        if (element.TryGetPropertyIgnoreCase(
            nameof(DataTableTemplate.RecordHeight), out JsonElement recordHeightElement
        )
            && recordHeightElement.ValueKind != JsonValueKind.Null
        ) {
            template.RecordHeight = recordHeightElement.GetDoubleValue(nameof(DataTableTemplate.RecordHeight));
        }

        if (columnDefinitions.Count > 0) {
            template.Columns.Clear();
            foreach (DataTableColumnDefinition columnDefinition in columnDefinitions) {
                template.Columns.Add(CreateTemplateColumn(columnDefinition));
            }
        }

        return template;
    }

    private static IReadOnlyList<DataTableColumnDefinition> ParseColumnDefinitions(JsonElement element) {
        if (!element.TryGetPropertyIgnoreCase(nameof(DataTableTemplate.Columns), out JsonElement columnsElement)) {
            return [];
        }

        if (columnsElement.ValueKind != JsonValueKind.Array) {
            throw new FormatException("DataTable template property 'Columns' must be an array.");
        }

        List<DataTableColumnDefinition> columns = [];
        foreach (JsonElement columnElement in columnsElement.EnumerateArray()) {
            columns.Add(ParseColumnDefinition(columnElement));
        }

        return columns.AsReadOnly();
    }

    private static DataTableColumnDefinition ParseColumnDefinition(JsonElement columnElement) {
        if (columnElement.ValueKind != JsonValueKind.Object) {
            throw new FormatException("Each DataTable column must be a JSON object.");
        }

        if (!columnElement.TryGetPropertyIgnoreCase(
            nameof(DataTableColumn.ColumnName), out JsonElement columnNameElement
        )) {
            throw new FormatException("DataTable column requires a 'ColumnName' property.");
        }

        string columnName = columnNameElement.GetStringValue(nameof(DataTableColumn.ColumnName));
        string? headerText = columnElement.TryGetPropertyIgnoreCase(
            nameof(DataTableColumn.HeaderText), out JsonElement headerTextElement
        )
            ? GetNullableStringValue(headerTextElement, nameof(DataTableColumn.HeaderText))
            : null;
        CellStyle? headerStyle = columnElement.TryGetPropertyIgnoreCase(
            nameof(DataTableColumn.HeaderStyle), out JsonElement headerStyleElement
        )
            && headerStyleElement.ValueKind != JsonValueKind.Null
            ? JsonStyleParser.Parse(headerStyleElement)
            : null;
        CellStyle? fieldStyle = columnElement.TryGetPropertyIgnoreCase(
            "FieldStyle", out JsonElement fieldStyleElement
        )
            && fieldStyleElement.ValueKind != JsonValueKind.Null
            ? JsonStyleParser.Parse(fieldStyleElement)
            : null;
        string? formula = columnElement.TryGetPropertyIgnoreCase("Formula", out JsonElement formulaElement)
            ? GetNullableStringValue(formulaElement, "Formula")
            : null;

        return new DataTableColumnDefinition(columnName, headerText, headerStyle, fieldStyle, formula);
    }

    private static IReadOnlyList<IDictionary<string, object?>> ParseRecords(JsonElement element) {
        if (!element.TryGetPropertyIgnoreCase("Records", out JsonElement recordsElement)) {
            return [];
        }

        if (recordsElement.ValueKind != JsonValueKind.Array) {
            throw new FormatException("DataTable template property 'Records' must be an array.");
        }

        List<IDictionary<string, object?>> records = [];
        foreach (JsonElement recordElement in recordsElement.EnumerateArray()) {
            records.Add(recordElement.ToDictionary());
        }

        return records.AsReadOnly();
    }

    private static System.Data.DataTable CreateDataTable(
        IReadOnlyList<DataTableColumnDefinition> columnDefinitions,
        IReadOnlyList<IDictionary<string, object?>> records
    ) {
        System.Data.DataTable dataTable = new();
        IReadOnlyList<string> columnNames = columnDefinitions.Count > 0
            ? GetColumnNames(columnDefinitions)
            : InferColumnNames(records);

        foreach (string columnName in columnNames) {
            dataTable.Columns.Add(columnName, typeof(object));
        }

        foreach (IDictionary<string, object?> record in records) {
            DataRow row = dataTable.NewRow();

            foreach (DataColumn column in dataTable.Columns) {
                object? value = record.TryGetValue(column.ColumnName, out object? recordValue)
                    ? recordValue
                    : null;

                row[column.ColumnName] = value ?? DBNull.Value;
            }

            dataTable.Rows.Add(row);
        }

        return dataTable;
    }

    private static IReadOnlyList<string> GetColumnNames(IReadOnlyList<DataTableColumnDefinition> columnDefinitions) {
        List<string> columnNames = [];
        foreach (DataTableColumnDefinition columnDefinition in columnDefinitions) {
            columnNames.Add(columnDefinition.ColumnName);
        }

        return columnNames.AsReadOnly();
    }

    private static IReadOnlyList<string> InferColumnNames(IReadOnlyList<IDictionary<string, object?>> records) {
        List<string> columnNames = [];
        HashSet<string> knownColumnNames = new(StringComparer.Ordinal);

        foreach (IDictionary<string, object?> record in records) {
            foreach (string columnName in record.Keys) {
                if (knownColumnNames.Add(columnName)) {
                    columnNames.Add(columnName);
                }
            }
        }

        return columnNames.AsReadOnly();
    }

    private static DataTableColumn CreateTemplateColumn(DataTableColumnDefinition columnDefinition) {
        DataTableColumn column = new() {
            ColumnName = columnDefinition.ColumnName,
            HeaderText = columnDefinition.HeaderText ?? columnDefinition.ColumnName
        };

        if (columnDefinition.HeaderStyle.HasValue) {
            column.HeaderStyle = columnDefinition.HeaderStyle.Value;
        }

        if (columnDefinition.FieldStyle.HasValue) {
            CellStyle fieldStyle = columnDefinition.FieldStyle.Value;
            column.FieldStyleGenerator = value => fieldStyle;
        }

        if (columnDefinition.Formula is not null) {
            column.FieldFormulaGenerator = value => columnDefinition.Formula;
        }

        return column;
    }

    private static string? GetNullableStringValue(JsonElement element, string propertyName) {
        return element.ValueKind == JsonValueKind.Null
            ? null
            : element.GetStringValue(propertyName);
    }

    private sealed record DataTableColumnDefinition(
        string ColumnName,
        string? HeaderText,
        CellStyle? HeaderStyle,
        CellStyle? FieldStyle,
        string? Formula
    );
}
