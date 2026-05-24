using System.Collections.Generic;

namespace CloudyWing.SpreadsheetExporter;

internal static class RendererCapabilityValidator {
    internal static IReadOnlyList<LayoutDiagnostic> GetDiagnostics(
        IReadOnlyList<SheetDefinition> sheets,
        CellFont? defaultFont,
        RendererCapabilities capabilities
    ) {
        List<LayoutDiagnostic> diagnostics = [];
        HashSet<string> reportedCapabilities = [];

        AppendDocumentDiagnostics(diagnostics, reportedCapabilities, sheets, defaultFont, capabilities);

        for (int sheetIndex = 0; sheetIndex < sheets.Count; sheetIndex++) {
            AppendSheetDiagnostics(
                diagnostics,
                reportedCapabilities,
                new SheetLayout(sheets[sheetIndex], defaultFont),
                sheetIndex,
                capabilities
            );
        }

        return diagnostics.AsReadOnly();
    }

    private static void AppendDocumentDiagnostics(
        List<LayoutDiagnostic> diagnostics,
        HashSet<string> reportedCapabilities,
        IReadOnlyList<SheetDefinition> sheets,
        CellFont? defaultFont,
        RendererCapabilities capabilities
    ) {
        if (sheets.Count > 1) {
            AddUnsupportedDiagnostic(
                diagnostics,
                reportedCapabilities,
                capabilities.SupportsMultipleSheets,
                nameof(RendererCapabilities.SupportsMultipleSheets),
                "Renderer capability 'SupportsMultipleSheets' is disabled, but the document contains multiple sheets.",
                null,
                null
            );
        }

        if (defaultFont is not null) {
            AddUnsupportedDiagnostic(
                diagnostics,
                reportedCapabilities,
                capabilities.SupportsStyles,
                nameof(RendererCapabilities.SupportsStyles),
                "Renderer capability 'SupportsStyles' is disabled, but the document default font is set.",
                null,
                null
            );
        }
    }

    private static void AppendSheetDiagnostics(
        List<LayoutDiagnostic> diagnostics,
        HashSet<string> reportedCapabilities,
        SheetLayout layout,
        int sheetIndex,
        RendererCapabilities capabilities
    ) {
        CellSource sheetSource = CreateSheetSource(layout, sheetIndex);

        if (layout.FreezePanes is not null) {
            AddUnsupportedDiagnostic(
                diagnostics,
                reportedCapabilities,
                capabilities.SupportsFreezePanes,
                nameof(RendererCapabilities.SupportsFreezePanes),
                "Renderer capability 'SupportsFreezePanes' is disabled, "
                + $"but sheet '{layout.SheetName}' uses freeze panes.",
                null,
                sheetSource
            );
        }

        if (layout.IsAutoFilterEnabled) {
            AddUnsupportedDiagnostic(
                diagnostics,
                reportedCapabilities,
                capabilities.SupportsAutoFilter,
                nameof(RendererCapabilities.SupportsAutoFilter),
                "Renderer capability 'SupportsAutoFilter' is disabled, "
                + $"but sheet '{layout.SheetName}' uses auto filter.",
                null,
                sheetSource
            );
        }

        if (layout.IsProtected) {
            AddUnsupportedDiagnostic(
                diagnostics,
                reportedCapabilities,
                capabilities.SupportsWorksheetProtection,
                nameof(RendererCapabilities.SupportsWorksheetProtection),
                "Renderer capability 'SupportsWorksheetProtection' is disabled, "
                + $"but sheet '{layout.SheetName}' is protected.",
                null,
                sheetSource
            );
        }

        if (UsesPageSettings(layout.PageSettings)) {
            AddUnsupportedDiagnostic(
                diagnostics,
                reportedCapabilities,
                capabilities.SupportsPageSettings,
                nameof(RendererCapabilities.SupportsPageSettings),
                "Renderer capability 'SupportsPageSettings' is disabled, "
                + $"but sheet '{layout.SheetName}' uses page settings.",
                null,
                sheetSource
            );
        }

        AppendCellDiagnostics(diagnostics, reportedCapabilities, layout, sheetIndex, capabilities);
    }

    private static void AppendCellDiagnostics(
        List<LayoutDiagnostic> diagnostics,
        HashSet<string> reportedCapabilities,
        SheetLayout layout,
        int sheetIndex,
        RendererCapabilities capabilities
    ) {
        for (int cellIndex = 0; cellIndex < layout.Cells.Count; cellIndex++) {
            Cell cell = layout.Cells[cellIndex];
            CellRange range = new(cell.Point.X, cell.Point.Y, cell.Size.Width, cell.Size.Height);
            CellSource cellSource = CreateCellSource(layout, sheetIndex, cellIndex);

            if (cell.Size.Width > 1 || cell.Size.Height > 1) {
                AddUnsupportedDiagnostic(
                    diagnostics,
                    reportedCapabilities,
                    capabilities.SupportsMergedCells,
                    nameof(RendererCapabilities.SupportsMergedCells),
                    "Renderer capability 'SupportsMergedCells' is disabled, "
                    + $"but sheet '{layout.SheetName}' uses merged cells.",
                    range,
                    cellSource
                );
            }

            if (cell.FormulaGenerator is not null) {
                AddUnsupportedDiagnostic(
                    diagnostics,
                    reportedCapabilities,
                    capabilities.SupportsFormulas,
                    nameof(RendererCapabilities.SupportsFormulas),
                    "Renderer capability 'SupportsFormulas' is disabled, "
                    + $"but sheet '{layout.SheetName}' contains formulas.",
                    range,
                    cellSource
                );
            }

            if (cell.DataValidationGenerator is not null) {
                AddUnsupportedDiagnostic(
                    diagnostics,
                    reportedCapabilities,
                    capabilities.SupportsDataValidation,
                    nameof(RendererCapabilities.SupportsDataValidation),
                    "Renderer capability 'SupportsDataValidation' is disabled, "
                    + $"but sheet '{layout.SheetName}' uses data validation.",
                    range,
                    cellSource
                );
            }

            if (cell.CellStyleGenerator is not null) {
                AddUnsupportedDiagnostic(
                    diagnostics,
                    reportedCapabilities,
                    capabilities.SupportsStyles,
                    nameof(RendererCapabilities.SupportsStyles),
                    "Renderer capability 'SupportsStyles' is disabled, "
                    + $"but sheet '{layout.SheetName}' contains styled cells.",
                    range,
                    cellSource
                );
            }
        }
    }

    private static void AddUnsupportedDiagnostic(
        List<LayoutDiagnostic> diagnostics,
        HashSet<string> reportedCapabilities,
        bool isSupported,
        string capabilityName,
        string message,
        CellRange? range,
        CellSource? source
    ) {
        if (isSupported || !reportedCapabilities.Add(capabilityName)) {
            return;
        }

        diagnostics.Add(new LayoutDiagnostic(
            LayoutDiagnosticCodes.UnsupportedRendererCapability,
            message,
            range,
            source
        ));
    }

    private static bool UsesPageSettings(PageSettings pageSettings) {
        return pageSettings.PageOrientation != PageOrientation.Portrait
            || pageSettings.PaperSize.Value != PaperSize.A4.Value;
    }

    private static CellSource CreateSheetSource(SheetLayout layout, int sheetIndex) {
        return new CellSource(
            sheetIndex,
            layout.SheetName,
            -1,
            nameof(SheetDefinition),
            $"Sheets[{sheetIndex}]"
        );
    }

    private static CellSource CreateCellSource(SheetLayout layout, int sheetIndex, int cellIndex) {
        return new CellSource(
            sheetIndex,
            layout.SheetName,
            -1,
            nameof(Cell),
            $"Sheets[{sheetIndex}].Cells[{cellIndex}]"
        );
    }
}
