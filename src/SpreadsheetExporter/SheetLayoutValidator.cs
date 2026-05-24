using System;
using System.Collections.Generic;
using System.Drawing;
using CloudyWing.SpreadsheetExporter.Templates;

namespace CloudyWing.SpreadsheetExporter;

internal static class SheetLayoutValidator {
    internal static IReadOnlyList<LayoutDiagnostic> GetDiagnostics(IReadOnlyList<SheetDefinition> sheets) {
        List<LayoutDiagnostic> diagnostics = [];

        for (int sheetIndex = 0; sheetIndex < sheets.Count; sheetIndex++) {
            AppendSheetDiagnostics(diagnostics, sheets[sheetIndex], sheetIndex);
        }

        return diagnostics.AsReadOnly();
    }

    private static void AppendSheetDiagnostics(
        List<LayoutDiagnostic> diagnostics, SheetDefinition sheet, int sheetIndex
    ) {
        Dictionary<Point, OccupiedCell> occupiedPoints = [];
        int rowOffset = 0;

        AppendColumnWidthDiagnostics(diagnostics, sheet, sheetIndex);

        for (int templateIndex = 0; templateIndex < sheet.Templates.Count; templateIndex++) {
            ISheetTemplate template = sheet.Templates[templateIndex];
            TemplateLayout layout = template.GetLayout();
            CellSource source = CreateSource(sheet, sheetIndex, templateIndex, template);

            AppendRowHeightDiagnostics(diagnostics, layout, rowOffset, source);
            AppendCellDiagnostics(diagnostics, occupiedPoints, layout, rowOffset, source);

            rowOffset += layout.RowSpan;
        }
    }

    private static void AppendColumnWidthDiagnostics(
        List<LayoutDiagnostic> diagnostics, SheetDefinition sheet, int sheetIndex
    ) {
        CellSource source = new(sheetIndex, sheet.SheetName, -1, nameof(SheetDefinition), "ColumnWidths");

        foreach (KeyValuePair<int, double> pair in sheet.ColumnWidths) {
            if (pair.Key < 0 || !IsAllowedColumnWidth(pair.Value)) {
                diagnostics.Add(new LayoutDiagnostic(
                    LayoutDiagnosticCodes.InvalidDimension,
                    $"Column width at index {pair.Key} has invalid value {pair.Value}.",
                    null,
                    source
                ));
            }
        }
    }

    private static void AppendRowHeightDiagnostics(
        List<LayoutDiagnostic> diagnostics, TemplateLayout layout, int rowOffset, CellSource source
    ) {
        foreach (KeyValuePair<int, double?> pair in layout.RowHeights) {
            int rowIndex = rowOffset + pair.Key;

            if (pair.Key < 0 || !IsAllowedRowHeight(pair.Value)) {
                diagnostics.Add(new LayoutDiagnostic(
                    LayoutDiagnosticCodes.InvalidDimension,
                    $"Row height at index {rowIndex} has invalid value {pair.Value}.",
                    null,
                    source
                ));
            }
        }
    }

    private static void AppendCellDiagnostics(
        List<LayoutDiagnostic> diagnostics,
        Dictionary<Point, OccupiedCell> occupiedPoints,
        TemplateLayout layout,
        int rowOffset,
        CellSource source
    ) {
        foreach (Cell cell in layout.Cells) {
            CellRange range = new(cell.Point.X, cell.Point.Y + rowOffset, cell.Size.Width, cell.Size.Height);
            AppendCellValueFormulaDiagnostics(diagnostics, cell, range, source);

            if (!range.IsValid) {
                diagnostics.Add(new LayoutDiagnostic(
                    LayoutDiagnosticCodes.InvalidCellRange,
                    $"Cell range ({range.Column}, {range.Row}, {range.ColumnSpan}, {range.RowSpan}) is invalid.",
                    range,
                    source
                ));
                continue;
            }

            OccupiedCell current = new(range, source);
            OccupiedCell? existing = FindOverlappingCell(occupiedPoints, range);
            if (existing is not null) {
                diagnostics.Add(new LayoutDiagnostic(
                    LayoutDiagnosticCodes.OverlappingCellRange,
                    $"Cell range ({range.Column}, {range.Row})-({range.LastColumn}, {range.LastRow}) overlaps " +
                    $"with range ({existing.Range.Column}, {existing.Range.Row})-" +
                    $"({existing.Range.LastColumn}, {existing.Range.LastRow}).",
                    range,
                    source
                ));
            }

            AddOccupiedPoints(occupiedPoints, current);
        }
    }

    private static void AppendCellValueFormulaDiagnostics(
        List<LayoutDiagnostic> diagnostics, Cell cell, CellRange range, CellSource source
    ) {
        if (cell.ValueGenerator is not null && cell.FormulaGenerator is not null) {
            diagnostics.Add(new LayoutDiagnostic(
                LayoutDiagnosticCodes.ValueAndFormulaConflict,
                "Cell defines both a value and a formula. Existing renderers keep formula output unchanged.",
                range,
                source
            ));
        }
    }

    private static OccupiedCell? FindOverlappingCell(
        Dictionary<Point, OccupiedCell> occupiedPoints, CellRange range
    ) {
        for (int column = range.Column; column <= range.LastColumn; column++) {
            for (int row = range.Row; row <= range.LastRow; row++) {
                Point point = new(column, row);
                if (occupiedPoints.TryGetValue(point, out OccupiedCell? occupiedCell)) {
                    return occupiedCell;
                }
            }
        }

        return null;
    }

    private static void AddOccupiedPoints(Dictionary<Point, OccupiedCell> occupiedPoints, OccupiedCell cell) {
        for (int column = cell.Range.Column; column <= cell.Range.LastColumn; column++) {
            for (int row = cell.Range.Row; row <= cell.Range.LastRow; row++) {
                occupiedPoints[new Point(column, row)] = cell;
            }
        }
    }

    private static CellSource CreateSource(
        SheetDefinition sheet, int sheetIndex, int templateIndex, ISheetTemplate template
    ) {
        Type templateType = template.GetType();
        string templateTypeName = templateType.FullName ?? templateType.Name;

        return new CellSource(
            sheetIndex,
            sheet.SheetName,
            templateIndex,
            templateTypeName,
            $"Sheets[{sheetIndex}].Templates[{templateIndex}]"
        );
    }

    private static bool IsAllowedColumnWidth(double width) {
        return width > 0
            || width == Constants.HiddenColumn
            || width == Constants.AutoFitColumnWidth;
    }

    private static bool IsAllowedRowHeight(double? height) {
        return !height.HasValue
            || height.Value > 0
            || height.Value == Constants.HiddenRow
            || height.Value == Constants.AutoFitRowHeight;
    }

    private sealed record OccupiedCell(CellRange Range, CellSource Source);
}
