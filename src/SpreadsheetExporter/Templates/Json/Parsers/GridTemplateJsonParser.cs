using System;
using System.Text.Json;
using CloudyWing.SpreadsheetExporter.Templates.Grid;

namespace CloudyWing.SpreadsheetExporter.Templates.Json.Parsers;

/// <summary>
/// Parses a JSON element into a <see cref="GridTemplate"/>.
/// </summary>
public class GridTemplateJsonParser : ITemplateJsonParser {
    /// <inheritdoc/>
    public string TypeName => "Grid";

    /// <inheritdoc/>
    public ISheetTemplate Parse(JsonElement element) {
        if (element.ValueKind != JsonValueKind.Object) {
            throw new FormatException("Grid template JSON must be an object.");
        }

        GridTemplate template = new();

        if (!element.TryGetPropertyIgnoreCase("Rows", out JsonElement rowsElement)) {
            return template;
        }

        if (rowsElement.ValueKind != JsonValueKind.Array) {
            throw new FormatException("Grid template property 'Rows' must be an array.");
        }

        foreach (JsonElement rowElement in rowsElement.EnumerateArray()) {
            if (rowElement.ValueKind != JsonValueKind.Object) {
                throw new FormatException("Each grid row must be a JSON object.");
            }

            double? height = rowElement.TryGetPropertyIgnoreCase("Height", out JsonElement heightElement)
                && heightElement.ValueKind != JsonValueKind.Null
                ? heightElement.GetDoubleValue("Rows[].Height")
                : null;

            template.CreateRow(height);

            if (!rowElement.TryGetPropertyIgnoreCase("Cells", out JsonElement cellsElement)) {
                continue;
            }

            if (cellsElement.ValueKind != JsonValueKind.Array) {
                throw new FormatException("Grid row property 'Cells' must be an array.");
            }

            foreach (JsonElement cellElement in cellsElement.EnumerateArray()) {
                ParseCell(template, cellElement);
            }
        }

        return template;
    }

    private static void ParseCell(GridTemplate template, JsonElement cellElement) {
        if (cellElement.ValueKind != JsonValueKind.Object) {
            throw new FormatException("Each grid cell must be a JSON object.");
        }

        bool hasValue = cellElement.TryGetPropertyIgnoreCase("Value", out JsonElement valueElement);
        bool hasFormula = cellElement.TryGetPropertyIgnoreCase("Formula", out JsonElement formulaElement);

        if (hasValue && hasFormula) {
            throw new InvalidOperationException(
                "Grid cell JSON cannot specify both 'Value' and 'Formula'."
            );
        }

        int columnSpan = cellElement.TryGetPropertyIgnoreCase(nameof(GridTemplate.ColumnSpan), out JsonElement columnSpanElement)
            ? columnSpanElement.GetInt32Value(nameof(GridTemplate.ColumnSpan))
            : 1;
        int rowSpan = cellElement.TryGetPropertyIgnoreCase(nameof(GridTemplate.RowSpan), out JsonElement rowSpanElement)
            ? rowSpanElement.GetInt32Value(nameof(GridTemplate.RowSpan))
            : 1;
        CellStyle? style = cellElement.TryGetPropertyIgnoreCase("Style", out JsonElement styleElement)
            && styleElement.ValueKind != JsonValueKind.Null
            ? JsonStyleParser.Parse(styleElement)
            : null;

        template.CreateCell(
            cell => {
                if (hasValue) {
                    object? value = valueElement.ToObject();
                    cell.ValueGenerator = (cellIndex, rowIndex) => value;
                }

                if (hasFormula) {
                    string formula = formulaElement.GetStringValue("Formula");
                    cell.FormulaGenerator = (cellIndex, rowIndex) => formula;
                }
            },
            columnSpan,
            rowSpan,
            style
        );
    }
}
