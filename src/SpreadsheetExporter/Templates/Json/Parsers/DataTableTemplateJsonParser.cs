using System;
using System.Collections.Generic;
using System.Data;
using System.Text.Json;
using CloudyWing.SpreadsheetExporter.Templates.DataTable;

namespace CloudyWing.SpreadsheetExporter.Templates.Json.Parsers;

/// <summary>
/// Parses a JSON element into a <see cref="DataTableTemplate"/>.
/// </summary>
public class DataTableTemplateJsonParser : ITemplateJsonParserWithContext {
    /// <inheritdoc/>
    public string TypeName => "DataTable";

    /// <inheritdoc/>
    public ISheetTemplate Parse(JsonElement element) {
        return Parse(element, JsonParseContext.Root);
    }

    /// <inheritdoc/>
    public ISheetTemplate Parse(JsonElement element, JsonParseContext context) {
        if (element.ValueKind != JsonValueKind.Object) {
            throw JsonParseExceptionFactory.InvalidType(context, "a JSON object");
        }

        IReadOnlyList<DataTableColumnDefinition> columnDefinitions = ParseColumnDefinitions(element, context);
        IReadOnlyList<IDictionary<string, object?>> records = ParseRecords(element, context);
        System.Data.DataTable dataTable = CreateDataTable(columnDefinitions, records);
        DataTableTemplate template = new(dataTable);

        if (element.TryGetPropertyIgnoreCase(
            nameof(DataTableTemplate.HeaderHeight), out JsonElement headerHeightElement
        )
            && headerHeightElement.ValueKind != JsonValueKind.Null
        ) {
            template.HeaderHeight = headerHeightElement.GetDoubleValue(
                context.Property(nameof(DataTableTemplate.HeaderHeight))
            );
        }

        if (element.TryGetPropertyIgnoreCase(
            nameof(DataTableTemplate.RecordHeight), out JsonElement recordHeightElement
        )
            && recordHeightElement.ValueKind != JsonValueKind.Null
        ) {
            template.RecordHeight = recordHeightElement.GetDoubleValue(
                context.Property(nameof(DataTableTemplate.RecordHeight))
            );
        }

        if (columnDefinitions.Count > 0) {
            template.Columns.Clear();
            foreach (DataTableColumnDefinition columnDefinition in columnDefinitions) {
                template.Columns.Add(CreateTemplateColumn(columnDefinition));
            }
        }

        return template;
    }

    private static IReadOnlyList<DataTableColumnDefinition> ParseColumnDefinitions(
        JsonElement element, JsonParseContext context
    ) {
        if (!element.TryGetPropertyIgnoreCase(nameof(DataTableTemplate.Columns), out JsonElement columnsElement)) {
            return [];
        }

        JsonParseContext columnsContext = context.Property(nameof(DataTableTemplate.Columns));
        if (columnsElement.ValueKind != JsonValueKind.Array) {
            throw JsonParseExceptionFactory.InvalidType(columnsContext, "an array");
        }

        List<DataTableColumnDefinition> columns = [];
        int columnIndex = 0;
        foreach (JsonElement columnElement in columnsElement.EnumerateArray()) {
            columns.Add(ParseColumnDefinition(columnElement, columnsContext.Index(columnIndex++)));
        }

        return columns.AsReadOnly();
    }

    private static DataTableColumnDefinition ParseColumnDefinition(
        JsonElement columnElement, JsonParseContext columnContext
    ) {
        if (columnElement.ValueKind != JsonValueKind.Object) {
            throw JsonParseExceptionFactory.InvalidType(columnContext, "a JSON object");
        }

        if (!columnElement.TryGetPropertyIgnoreCase(
            nameof(DataTableColumn.ColumnName), out JsonElement columnNameElement
        )) {
            throw JsonParseExceptionFactory.MissingRequiredProperty(
                columnContext.Property(nameof(DataTableColumn.ColumnName))
            );
        }

        string columnName = columnNameElement.GetStringValue(
            columnContext.Property(nameof(DataTableColumn.ColumnName))
        );
        string? headerText = columnElement.TryGetPropertyIgnoreCase(
            nameof(DataTableColumn.HeaderText), out JsonElement headerTextElement
        )
            ? GetNullableStringValue(headerTextElement, columnContext.Property(nameof(DataTableColumn.HeaderText)))
            : null;
        CellStyle? headerStyle = JsonStyleParser.ParseOptionalStyle(
            columnElement, nameof(DataTableColumn.HeaderStyle), "HeaderStyleName", columnContext
        );
        CellStyle? fieldStyle = JsonStyleParser.ParseOptionalStyle(
            columnElement, "FieldStyle", "FieldStyleName", columnContext
        );
        string? formula = columnElement.TryGetPropertyIgnoreCase("Formula", out JsonElement formulaElement)
            ? GetNullableStringValue(formulaElement, columnContext.Property("Formula"))
            : null;

        return new DataTableColumnDefinition(columnName, headerText, headerStyle, fieldStyle, formula);
    }

    private static IReadOnlyList<IDictionary<string, object?>> ParseRecords(
        JsonElement element, JsonParseContext context
    ) {
        if (!element.TryGetPropertyIgnoreCase("Records", out JsonElement recordsElement)) {
            return [];
        }

        JsonParseContext recordsContext = context.Property("Records");
        if (recordsElement.ValueKind != JsonValueKind.Array) {
            throw JsonParseExceptionFactory.InvalidType(recordsContext, "an array");
        }

        List<IDictionary<string, object?>> records = [];
        int recordIndex = 0;
        foreach (JsonElement recordElement in recordsElement.EnumerateArray()) {
            records.Add(recordElement.ToDictionary(recordsContext.Index(recordIndex++)));
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

    private static string? GetNullableStringValue(JsonElement element, JsonParseContext context) {
        return element.ValueKind == JsonValueKind.Null
            ? null
            : element.GetStringValue(context);
    }

    private sealed record DataTableColumnDefinition(
        string ColumnName,
        string? HeaderText,
        CellStyle? HeaderStyle,
        CellStyle? FieldStyle,
        string? Formula
    );
}
