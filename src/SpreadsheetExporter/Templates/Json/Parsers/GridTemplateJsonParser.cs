using System;
using System.Text.Json;
using CloudyWing.SpreadsheetExporter.Templates.Grid;

namespace CloudyWing.SpreadsheetExporter.Templates.Json.Parsers;

/// <summary>
/// Parses a JSON element into a <see cref="GridTemplate"/>.
/// </summary>
public class GridTemplateJsonParser : ITemplateJsonParserWithContext {
    /// <inheritdoc/>
    public string TypeName => "Grid";

    /// <inheritdoc/>
    public ISheetTemplate Parse(JsonElement element) {
        return Parse(element, JsonParseContext.Root);
    }

    /// <inheritdoc/>
    public ISheetTemplate Parse(JsonElement element, JsonParseContext context) {
        if (element.ValueKind != JsonValueKind.Object) {
            throw JsonParseExceptionFactory.InvalidType(context, "a JSON object");
        }

        GridTemplate template = new();

        if (!element.TryGetPropertyIgnoreCase("Rows", out JsonElement rowsElement)) {
            return template;
        }

        JsonParseContext rowsContext = context.Property("Rows");
        if (rowsElement.ValueKind != JsonValueKind.Array) {
            throw JsonParseExceptionFactory.InvalidType(rowsContext, "an array");
        }

        int rowIndex = 0;
        foreach (JsonElement rowElement in rowsElement.EnumerateArray()) {
            JsonParseContext rowContext = rowsContext.Index(rowIndex++);
            if (rowElement.ValueKind != JsonValueKind.Object) {
                throw JsonParseExceptionFactory.InvalidType(rowContext, "a JSON object");
            }

            double? height = rowElement.TryGetPropertyIgnoreCase("Height", out JsonElement heightElement)
                && heightElement.ValueKind != JsonValueKind.Null
                ? heightElement.GetDoubleValue(rowContext.Property("Height"))
                : null;

            template.CreateRow(height);

            if (!rowElement.TryGetPropertyIgnoreCase("Cells", out JsonElement cellsElement)) {
                continue;
            }

            JsonParseContext cellsContext = rowContext.Property("Cells");
            if (cellsElement.ValueKind != JsonValueKind.Array) {
                throw JsonParseExceptionFactory.InvalidType(cellsContext, "an array");
            }

            int cellIndex = 0;
            foreach (JsonElement cellElement in cellsElement.EnumerateArray()) {
                ParseCell(template, cellElement, cellsContext.Index(cellIndex++));
            }
        }

        return template;
    }

    private static void ParseCell(GridTemplate template, JsonElement cellElement, JsonParseContext cellContext) {
        if (cellElement.ValueKind != JsonValueKind.Object) {
            throw JsonParseExceptionFactory.InvalidType(cellContext, "a JSON object");
        }

        bool hasValue = cellElement.TryGetPropertyIgnoreCase("Value", out JsonElement valueElement);
        bool hasFormula = cellElement.TryGetPropertyIgnoreCase("Formula", out JsonElement formulaElement);

        if (hasValue && hasFormula) {
            throw JsonParseExceptionFactory.InvalidOperation(
                cellContext, "cannot specify both 'Value' and 'Formula'."
            );
        }

        int columnSpan = cellElement.TryGetPropertyIgnoreCase(
            nameof(GridTemplate.ColumnSpan), out JsonElement columnSpanElement
        )
            ? columnSpanElement.GetInt32Value(cellContext.Property(nameof(GridTemplate.ColumnSpan)))
            : 1;
        int rowSpan = cellElement.TryGetPropertyIgnoreCase(nameof(GridTemplate.RowSpan), out JsonElement rowSpanElement)
            ? rowSpanElement.GetInt32Value(cellContext.Property(nameof(GridTemplate.RowSpan)))
            : 1;
        CellStyle? style = cellElement.TryGetPropertyIgnoreCase("Style", out JsonElement styleElement)
            && styleElement.ValueKind != JsonValueKind.Null
            ? JsonStyleParser.Parse(styleElement, cellContext.Property("Style"))
            : null;
        DataValidation? dataValidation = cellElement.TryGetPropertyIgnoreCase(
            "DataValidation", out JsonElement dataValidationElement
        )
            ? JsonDataValidationParser.Parse(dataValidationElement, cellContext.Property("DataValidation"))
            : null;

        template.CreateCell(
            cell => {
                if (hasValue) {
                    object? value = valueElement.ToObject(cellContext.Property("Value"));
                    cell.ValueGenerator = (cellIndex, rowIndex) => value;
                }

                if (hasFormula) {
                    string formula = formulaElement.GetStringValue(cellContext.Property("Formula"));
                    cell.FormulaGenerator = (cellIndex, rowIndex) => formula;
                }

                if (dataValidation is not null) {
                    cell.DataValidationGenerator = (cellIndex, rowIndex) => dataValidation;
                }
            },
            columnSpan,
            rowSpan,
            style
        );
    }
}
