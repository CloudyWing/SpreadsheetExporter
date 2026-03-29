using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text.Json;
using CloudyWing.SpreadsheetExporter.Exceptions;
using CloudyWing.SpreadsheetExporter.Templates;
using CloudyWing.SpreadsheetExporter.Templates.Json;

namespace CloudyWing.SpreadsheetExporter;

/// <summary>
/// Represents a spreadsheet document that manages sheet definitions and coordinates rendering.
/// </summary>
public sealed class SpreadsheetDocument {
    private readonly ISpreadsheetRenderer renderer;
    private readonly List<SheetDefinition> sheets = [];

    /// <summary>
    /// Initializes a new instance of the <see cref="SpreadsheetDocument"/> class.
    /// </summary>
    /// <param name="renderer">The renderer used to produce the output binary.</param>
    /// <exception cref="ArgumentNullException"><paramref name="renderer"/> is <see langword="null"/>.</exception>
    public SpreadsheetDocument(ISpreadsheetRenderer renderer) {
        ArgumentNullException.ThrowIfNull(renderer);
        this.renderer = renderer;
    }

    /// <summary>
    /// Gets the MIME content-type from the associated renderer.
    /// </summary>
    public string ContentType => renderer.ContentType;

    /// <summary>
    /// Gets the file name extension from the associated renderer.
    /// </summary>
    public string FileNameExtension => renderer.FileNameExtension;

    /// <summary>
    /// Gets or sets the default sheet name prefix used when creating sheets without an explicit name.
    /// </summary>
    /// <value>The default sheet name prefix. Default is "工作表".</value>
    public string DefaultSheetName { get; set; } = "工作表";

    /// <summary>
    /// Gets or sets the default font applied to all cells unless overridden at the sheet or cell level.
    /// </summary>
    /// <value>The default font, or <see langword="null"/> to use the renderer default.</value>
    public CellFont? DefaultFont { get; set; }

    /// <summary>
    /// Sets the default sheet name prefix used when creating sheets without an explicit name.
    /// </summary>
    /// <param name="defaultSheetName">The default sheet name prefix.</param>
    /// <returns>The current <see cref="SpreadsheetDocument"/> instance.</returns>
    public SpreadsheetDocument SetDefaultSheetName(string defaultSheetName) {
        DefaultSheetName = defaultSheetName;
        return this;
    }

    /// <summary>
    /// Sets the default font applied to all cells unless overridden at the sheet or cell level.
    /// </summary>
    /// <param name="defaultFont">The default font, or <see langword="null"/> to use the renderer default.</param>
    /// <returns>The current <see cref="SpreadsheetDocument"/> instance.</returns>
    public SpreadsheetDocument SetDefaultFont(CellFont? defaultFont) {
        DefaultFont = defaultFont;
        return this;
    }

    /// <summary>
    /// Gets the last sheet definition. Creates a new sheet if none exist.
    /// </summary>
    public SheetDefinition LastSheet => sheets.Count > 0 ? sheets[^1] : CreateSheet();

    /// <summary>
    /// Creates a new sheet definition and adds it to this document.
    /// </summary>
    /// <param name="sheetName">The sheet name. If <see langword="null"/> or whitespace, a default name is generated.</param>
    /// <param name="defaultRowHeight">The default row height for the sheet.</param>
    /// <returns>The created sheet definition.</returns>
    public SheetDefinition CreateSheet(string? sheetName = null, double? defaultRowHeight = null) {
        sheetName = GetUniqueSheetName(sheetName);
        SheetDefinition sheet = new(sheetName) {
            DefaultRowHeight = defaultRowHeight
        };
        sheets.Add(sheet);
        return sheet;
    }

    /// <summary>
    /// Gets the sheet definition at the specified zero-based index.
    /// </summary>
    /// <param name="index">The zero-based index.</param>
    /// <returns>The sheet definition.</returns>
    /// <exception cref="ArgumentOutOfRangeException">
    /// <paramref name="index"/> is less than 0 or greater than or equal to the number of sheets.
    /// </exception>
    public SheetDefinition GetSheet(int index) {
        if (index < 0 || index >= sheets.Count) {
            throw new ArgumentOutOfRangeException(
                nameof(index), index,
                $"Index must be between 0 and {sheets.Count - 1}."
            );
        }
        return sheets[index];
    }

    /// <summary>
    /// Tries to get the sheet definition at the specified index.
    /// </summary>
    /// <param name="index">The zero-based index.</param>
    /// <param name="sheet">The sheet definition if found; otherwise, <see langword="null"/>.</param>
    /// <returns><see langword="true"/> if the index is valid; otherwise, <see langword="false"/>.</returns>
    public bool TryGetSheet(int index, out SheetDefinition? sheet) {
        if (index >= 0 && index < sheets.Count) {
            sheet = sheets[index];
            return true;
        }
        sheet = null;
        return false;
    }

    /// <summary>
    /// Exports this document as a byte array using the associated renderer.
    /// </summary>
    /// <returns>The rendered spreadsheet bytes.</returns>
    /// <exception cref="SheetDefinitionNotFoundException">No sheets have been created.</exception>
    public byte[] Export() {
        if (sheets.Count == 0) {
            throw new SheetDefinitionNotFoundException();
        }

        IEnumerable<SheetLayout> layouts = sheets.Select(
            s => new SheetLayout(s, DefaultFont)
        );
        return renderer.Render(layouts);
    }

    /// <summary>
    /// Exports this document to a file at the specified path.
    /// </summary>
    /// <param name="path">The output file path.</param>
    /// <param name="fileMode">The file creation mode.</param>
    public void ExportFile(string path, SpreadsheetFileMode fileMode = SpreadsheetFileMode.Create) {
        byte[] bytes = Export();
        FileMode mode = fileMode == SpreadsheetFileMode.Create
            ? FileMode.Create
            : FileMode.CreateNew;

        using FileStream stream = new(path, mode, FileAccess.Write);
        stream.Write(bytes, 0, bytes.Length);
    }

    /// <summary>
    /// Creates a <see cref="SpreadsheetDocument"/> from a JSON string using the renderer
    /// registered via <see cref="SpreadsheetManager.SetRenderer"/>.
    /// </summary>
    /// <param name="json">The JSON string describing sheets and templates.</param>
    /// <returns>A configured <see cref="SpreadsheetDocument"/>.</returns>
    public static SpreadsheetDocument FromJson(string json) {
        ArgumentNullException.ThrowIfNull(json);

        SpreadsheetDocument document = SpreadsheetManager.CreateDocument();

        using JsonDocument jsonDocument = JsonDocument.Parse(json);
        JsonElement root = jsonDocument.RootElement;
        if (root.ValueKind != JsonValueKind.Array) {
            throw new FormatException("Root JSON element must be an array of sheet definitions.");
        }

        foreach (JsonElement sheetElement in root.EnumerateArray()) {
            if (sheetElement.ValueKind != JsonValueKind.Object) {
                throw new FormatException("Each sheet definition must be a JSON object.");
            }

            string? sheetName = sheetElement.TryGetPropertyIgnoreCase("SheetName", out JsonElement sheetNameElement)
                && sheetNameElement.ValueKind != JsonValueKind.Null
                ? sheetNameElement.GetStringValue("SheetName")
                : null;
            double? defaultRowHeight = sheetElement.TryGetPropertyIgnoreCase("DefaultRowHeight", out JsonElement defaultRowHeightElement)
                && defaultRowHeightElement.ValueKind != JsonValueKind.Null
                ? defaultRowHeightElement.GetDoubleValue("DefaultRowHeight")
                : null;

            SheetDefinition sheet = document.CreateSheet(sheetName, defaultRowHeight);
            PopulateSheet(sheet, sheetElement);
        }

        return document;
    }

    private string GetUniqueSheetName(string? sheetName) {
        if (string.IsNullOrWhiteSpace(sheetName)) {
            return GenerateDefaultSheetName();
        }
        if (IsSheetNameExists(sheetName)) {
            return GenerateUniqueSheetName(sheetName);
        }
        return sheetName;
    }

    private string GenerateDefaultSheetName() {
        string baseName = DefaultSheetName;
        int i = 1;
        string name;
        do {
            name = $"{baseName}{i++}";
        } while (IsSheetNameExists(name));
        return name;
    }

    private string GenerateUniqueSheetName(string sheetName) {
        int i = 1;
        string name;
        do {
            name = $"{sheetName}({i++})";
        } while (IsSheetNameExists(name));
        return name;
    }

    private bool IsSheetNameExists(string sheetName) {
        return sheets.Select(x => x.SheetName).Contains(sheetName);
    }

    private static void PopulateSheet(SheetDefinition sheet, JsonElement sheetElement) {
        if (sheetElement.TryGetPropertyIgnoreCase(nameof(SheetDefinition.Password), out JsonElement passwordElement)
            && passwordElement.ValueKind != JsonValueKind.Null) {
            sheet.Password = passwordElement.GetStringValue(nameof(SheetDefinition.Password));
        }

        if (sheetElement.TryGetPropertyIgnoreCase(nameof(SheetDefinition.FreezePanes), out JsonElement freezePanesElement)
            && freezePanesElement.ValueKind == JsonValueKind.Object) {
            if (!freezePanesElement.TryGetPropertyIgnoreCase("Row", out JsonElement rowElement)) {
                throw new FormatException("FreezePanes requires a 'Row' property.");
            }
            if (!freezePanesElement.TryGetPropertyIgnoreCase("Column", out JsonElement columnElement)) {
                throw new FormatException("FreezePanes requires a 'Column' property.");
            }
            sheet.FreezePanes = new Point(
                columnElement.GetInt32Value("FreezePanes.Column"),
                rowElement.GetInt32Value("FreezePanes.Row")
            );
        }

        if (sheetElement.TryGetPropertyIgnoreCase(nameof(SheetDefinition.IsAutoFilterEnabled), out JsonElement isAutoFilterEnabledElement)) {
            sheet.IsAutoFilterEnabled = isAutoFilterEnabledElement.GetBooleanValue(nameof(SheetDefinition.IsAutoFilterEnabled));
        }

        if (sheetElement.TryGetPropertyIgnoreCase("ColumnWidths", out JsonElement columnWidthsElement)) {
            PopulateColumnWidths(sheet, columnWidthsElement);
        }

        if (sheetElement.TryGetPropertyIgnoreCase(nameof(SheetDefinition.PageSettings), out JsonElement pageSettingsElement)
            && pageSettingsElement.ValueKind != JsonValueKind.Null) {
            PopulatePageSettings(sheet.PageSettings, pageSettingsElement);
        }

        if (sheetElement.TryGetPropertyIgnoreCase(nameof(SheetDefinition.Metadata), out JsonElement metadataElement)
            && metadataElement.ValueKind != JsonValueKind.Null) {
            foreach (KeyValuePair<string, object?> pair in metadataElement.ToDictionary()) {
                sheet.Metadata[pair.Key] = pair.Value;
            }
        }

        if (!sheetElement.TryGetPropertyIgnoreCase(nameof(SheetDefinition.Templates), out JsonElement templatesElement)) {
            return;
        }

        if (templatesElement.ValueKind != JsonValueKind.Array) {
            throw new FormatException("Sheet property 'Templates' must be an array.");
        }

        foreach (JsonElement templateElement in templatesElement.EnumerateArray()) {
            sheet.AddTemplate(CreateTemplate(templateElement));
        }
    }

    private static void PopulateColumnWidths(SheetDefinition sheet, JsonElement columnWidthsElement) {
        if (columnWidthsElement.ValueKind == JsonValueKind.Object) {
            foreach (JsonProperty property in columnWidthsElement.EnumerateObject()) {
                if (!int.TryParse(property.Name, NumberStyles.Integer, CultureInfo.InvariantCulture, out int index)) {
                    throw new FormatException(
                        $"Column width key '{property.Name}' must be a zero-based integer index."
                    );
                }

                sheet.SetColumnWidth(index, property.Value.GetDoubleValue($"ColumnWidths.{property.Name}"));
            }

            return;
        }

        if (columnWidthsElement.ValueKind == JsonValueKind.Array) {
            foreach (JsonElement columnWidthElement in columnWidthsElement.EnumerateArray()) {
                if (columnWidthElement.ValueKind != JsonValueKind.Object) {
                    throw new FormatException("Each 'ColumnWidths' array item must be an object.");
                }

                if (!columnWidthElement.TryGetPropertyIgnoreCase("Index", out JsonElement indexElement)) {
                    throw new FormatException("Column width item requires an 'Index' property.");
                }

                if (!columnWidthElement.TryGetPropertyIgnoreCase("Width", out JsonElement widthElement)) {
                    throw new FormatException("Column width item requires a 'Width' property.");
                }

                sheet.SetColumnWidth(
                    indexElement.GetInt32Value("ColumnWidths[].Index"),
                    widthElement.GetDoubleValue("ColumnWidths[].Width")
                );
            }

            return;
        }

        throw new FormatException("Sheet property 'ColumnWidths' must be an object or array.");
    }

    private static void PopulatePageSettings(PageSettings pageSettings, JsonElement pageSettingsElement) {
        if (pageSettingsElement.ValueKind != JsonValueKind.Object) {
            throw new FormatException("PageSettings must be a JSON object.");
        }

        if (pageSettingsElement.TryGetPropertyIgnoreCase(nameof(PageSettings.PageOrientation), out JsonElement orientationElement)) {
            pageSettings.PageOrientation = ParseEnum<PageOrientation>(
                orientationElement,
                $"{nameof(PageSettings)}.{nameof(PageSettings.PageOrientation)}"
            );
        }

        if (pageSettingsElement.TryGetPropertyIgnoreCase(nameof(PageSettings.PaperSize), out JsonElement paperSizeElement)) {
            pageSettings.PaperSize = ParsePaperSize(
                paperSizeElement,
                $"{nameof(PageSettings)}.{nameof(PageSettings.PaperSize)}"
            );
        }
    }

    private static ISheetTemplate CreateTemplate(JsonElement templateElement) {
        if (templateElement.ValueKind != JsonValueKind.Object) {
            throw new FormatException("Each template definition must be a JSON object.");
        }

        if (!templateElement.TryGetPropertyIgnoreCase("Type", out JsonElement typeElement)) {
            throw new FormatException("Template definition requires a 'Type' property.");
        }

        string typeName = typeElement.GetStringValue("Type");
        return JsonTemplateRegistry.Create(typeName, templateElement);
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

    private static PaperSize ParsePaperSize(JsonElement element, string propertyName) {
        if (element.ValueKind == JsonValueKind.String) {
            string? raw = element.GetString();
            if (int.TryParse(raw, NumberStyles.Integer, CultureInfo.InvariantCulture, out int numericStringValue)) {
                return ParsePaperSizeFromValue(numericStringValue);
            }

            PaperSize namedPaperSize = GetKnownPaperSizes()
                .FirstOrDefault(pair => string.Equals(pair.Key, raw, StringComparison.OrdinalIgnoreCase))
                .Value;

            return namedPaperSize ?? throw new FormatException(
                $"Property '{propertyName}' contains unknown paper size '{raw}'."
            );
        }

        if (element.ValueKind == JsonValueKind.Number && element.TryGetInt32(out int numericValue)) {
            return ParsePaperSizeFromValue(numericValue);
        }

        throw new FormatException(
            $"Property '{propertyName}' must be a paper size name or numeric value."
        );
    }

    private static PaperSize ParsePaperSizeFromValue(int value) {
        return GetKnownPaperSizes()
            .Select(pair => pair.Value)
            .FirstOrDefault(paperSize => paperSize.Value == value)
            ?? new PaperSize(value);
    }

    private static Dictionary<string, PaperSize> GetKnownPaperSizes() {
        return typeof(PaperSize)
            .GetProperties(BindingFlags.Public | BindingFlags.Static)
            .Where(property => property.PropertyType == typeof(PaperSize))
            .ToDictionary(
                property => property.Name,
                property => (PaperSize)property.GetValue(null)!,
                StringComparer.OrdinalIgnoreCase
            );
    }
}
