using System.Collections.Generic;
using System.Drawing;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using OfficeOpenXml.Style.XmlAccess;

namespace CloudyWing.SpreadsheetExporter.Excel.EPPlus;

/// <summary>
/// The excel exporter, using epplus.
/// </summary>
/// <seealso cref="ExporterBase" />
public class ExcelExporter : ExporterBase {
    private readonly Dictionary<HorizontalAlignment, ExcelHorizontalAlignment> horizontalAlignmentMap = new() {
        [HorizontalAlignment.General] = ExcelHorizontalAlignment.General,
        [HorizontalAlignment.Left] = ExcelHorizontalAlignment.Left,
        [HorizontalAlignment.Center] = ExcelHorizontalAlignment.Center,
        [HorizontalAlignment.Right] = ExcelHorizontalAlignment.Right,
        [HorizontalAlignment.Justify] = ExcelHorizontalAlignment.Justify
    };

    private readonly Dictionary<VerticalAlignment, ExcelVerticalAlignment> verticalAlignmentMap = new() {
        [VerticalAlignment.Top] = ExcelVerticalAlignment.Top,
        [VerticalAlignment.Middle] = ExcelVerticalAlignment.Center,
        [VerticalAlignment.Bottom] = ExcelVerticalAlignment.Bottom
    };

    /// <inheritdoc/>
    public override string ContentType => "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

    /// <inheritdoc/>
    public override string FileNameExtension => ".xlsx";

    /// <inheritdoc/>
    protected override byte[] ExecuteExport(IEnumerable<SheeterContext> contexts) {
        using ExcelPackage package = new();

        if (DefaultFont.HasValue) {
            SetDefaultFont(package, DefaultFont.Value);
        }

        foreach (SheeterContext context in contexts) {
            CreateSheetToWorkbook(package.Workbook, context);
        }

        return HasPassword
            ? package.GetAsByteArray(Password)
            : package.GetAsByteArray();
    }

    private void SetDefaultFont(ExcelPackage package, CellFont font) {
        ExcelFontXml defaultFont = package.Workbook.Styles.Fonts[0];
        if (!string.IsNullOrWhiteSpace(font.Name)) {
            defaultFont.Name = font.Name;
        }

        if (font.Size != 0) {
            defaultFont.Size = font.Size;
        }

        if (font.Color != Color.Empty) {
            defaultFont.Color.SetColor(font.Color);
        }

        defaultFont.Bold = (font.Style & FontStyles.IsBold) == FontStyles.IsBold;
        defaultFont.Italic = (font.Style & FontStyles.IsItalic) == FontStyles.IsItalic;
        defaultFont.UnderLine = (font.Style & FontStyles.HasUnderline) == FontStyles.HasUnderline;
        defaultFont.Strike = (font.Style & FontStyles.IsStrikeout) == FontStyles.IsStrikeout;
    }

    private void CreateSheetToWorkbook(ExcelWorkbook workbook, SheeterContext context) {
        ExcelWorksheet sheet = workbook.Worksheets.Add(context.SheetName);

        if (context.DefaultRowHeight.HasValue) {
            sheet.DefaultRowHeight = context.DefaultRowHeight.Value;
        }

        SetSheetCells(sheet, context.Cells);
        SetSheetColumnWidths(sheet, context.ColumnWidths);
        SetSheetRowHeights(sheet, context.RowHeights);

        if (context.FreezePanes.HasValue) {
            // EPPlus uses 1-based index, FreezePanes specifies the first unfrozen cell
            sheet.View.FreezePanes(context.FreezePanes.Value.Y + 1, context.FreezePanes.Value.X + 1);
        }

        if (context.IsAutoFilterEnabled && context.Cells.Count > 0) {
            // Calculate the range for auto filter based on all cells
            int maxRow = 0;
            int maxCol = 0;
            foreach (Cell cell in context.Cells) {
                int endRow = cell.Point.Y + cell.Size.Height - 1;
                int endCol = cell.Point.X + cell.Size.Width - 1;
                if (endRow > maxRow) maxRow = endRow;
                if (endCol > maxCol) maxCol = endCol;
            }
            // EPPlus uses 1-based index
            sheet.Cells[1, 1, maxRow + 1, maxCol + 1].AutoFilter = true;
        }

        if (context.IsProtected) {
            sheet.Protection.SetPassword(context.Password);
        }

        sheet.PrinterSettings.Orientation = context.PageSettings.PageOrientation == PageOrientation.Portrait
            ? eOrientation.Portrait : eOrientation.Landscape;
        sheet.PrinterSettings.PaperSize = (ePaperSize)context.PageSettings.PaperSize.Value;

        sheet.Calculate();

        OnSheetCreated(new SheetCreatedEventArgs(sheet, context));
    }

    private void SetSheetCells(ExcelWorksheet sheet, IReadOnlyList<Cell> cells) {
        foreach (Cell cell in cells) {
            int startRow = cell.Point.Y + 1;
            int startColumn = cell.Point.X + 1;
            int endRow = cell.Point.Y + cell.Size.Height;
            int endColumn = cell.Point.X + cell.Size.Width;

            ExcelRange range = sheet.Cells[startRow, startColumn, endRow, endColumn];
            range.Merge = startRow != endRow || startColumn != endColumn;

            string formula = cell.GetFormula();

            if (string.IsNullOrWhiteSpace(formula)) {
                range.Value = cell.GetValue();
            } else {
                range.Formula = formula;
            }

            SetCellStyleToExcel(range.Style, cell.GetCellStyle());
        }
    }

    private void SetCellStyleToExcel(ExcelStyle excelCellStyle, CellStyle cellStyle) {
        excelCellStyle.HorizontalAlignment = horizontalAlignmentMap[cellStyle.HorizontalAlignment];
        excelCellStyle.VerticalAlignment = verticalAlignmentMap[cellStyle.VerticalAlignment];

        if (cellStyle.HasBorder) {
            excelCellStyle.Border.Bottom.Style = ExcelBorderStyle.Thin;
            excelCellStyle.Border.Left.Style = ExcelBorderStyle.Thin;
            excelCellStyle.Border.Right.Style = ExcelBorderStyle.Thin;
            excelCellStyle.Border.Top.Style = ExcelBorderStyle.Thin;
        }

        excelCellStyle.WrapText = cellStyle.WrapText;
        if (cellStyle.BackgroundColor != Color.Empty) {
            excelCellStyle.Fill.PatternType = ExcelFillStyle.Solid;
            excelCellStyle.Fill.BackgroundColor.SetColor(cellStyle.BackgroundColor);
        }

        if (cellStyle.Font != CellFont.Empty) {
            SetFontToExcel(excelCellStyle.Font, cellStyle.Font);
        }

        if (!string.IsNullOrWhiteSpace(cellStyle.DataFormat)) {
            excelCellStyle.Numberformat.Format = cellStyle.DataFormat;
        }

        excelCellStyle.Locked = cellStyle.IsLocked;
    }

    private void SetFontToExcel(ExcelFont excelFont, CellFont font) {
        if (!string.IsNullOrWhiteSpace(font.Name)) {
            excelFont.Name = font.Name;
        }

        if (font.Size != 0) {
            excelFont.Size = font.Size;
        }

        if (font.Color != Color.Empty) {
            excelFont.Color.SetColor(font.Color);
        }

        excelFont.Bold = (font.Style & FontStyles.IsBold) == FontStyles.IsBold;
        excelFont.Italic = (font.Style & FontStyles.IsItalic) == FontStyles.IsItalic;
        excelFont.UnderLine = (font.Style & FontStyles.HasUnderline) == FontStyles.HasUnderline;
        excelFont.Strike = (font.Style & FontStyles.IsStrikeout) == FontStyles.IsStrikeout;
    }

    private void SetSheetColumnWidths(ExcelWorksheet sheet, IReadOnlyDictionary<int, double> columnWidths) {
        foreach (KeyValuePair<int, double> pair in columnWidths) {
            // EPPlus 從 1 開始算
            ExcelColumn column = sheet.Column(pair.Key + 1);
            if (pair.Value <= Constants.AutoFitColumnWidth) {
                column.AutoFit();
            } else if (pair.Value == Constants.HiddenColumn) {
                column.Hidden = true;
            } else {
                column.Width = pair.Value;
            }
        }
    }

    private void SetSheetRowHeights(ExcelWorksheet sheet, IReadOnlyDictionary<int, double?> rowHeights) {
        foreach (KeyValuePair<int, double?> pair in rowHeights) {
            // EPPlus 從 1 開始算
            ExcelRow row = sheet.Row(pair.Key + 1);
            if (pair.Value <= Constants.AutoFiteRowHeight) {
                row.CustomHeight = false;
            } else if (pair.Value == Constants.HiddenRow) {
                row.Hidden = true;
            } else if (pair.Value.HasValue) {
                row.Height = pair.Value.Value;
            }
        }
    }
}
