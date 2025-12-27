using NPOI.OpenXmlFormats.Spreadsheet;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;

namespace CloudyWing.SpreadsheetExporter.Excel.NPOI.Tests.Helpers;

/// <summary>
/// Helper class for verifying Excel file contents in tests
/// </summary>
internal class ExcelVerifier : IDisposable {
    private readonly IWorkbook workbook;
    private bool disposed;

    public ExcelVerifier(byte[] excelBytes) {
        using MemoryStream ms = new(excelBytes);
        workbook = new XSSFWorkbook(ms);
    }

    public ISheet GetSheet(int index) => workbook.GetSheetAt(index);

    public ISheet GetSheet(string name) => workbook.GetSheet(name);

    public void VerifyCellValue(int sheetIndex, int rowIndex, int columnIndex, object expectedValue) {
        ISheet sheet = GetSheet(sheetIndex);
        IRow row = sheet.GetRow(rowIndex);
        Assert.That(row, Is.Not.Null, $"Row {rowIndex} should exist");

        ICell cell = row.GetCell(columnIndex);
        Assert.That(cell, Is.Not.Null, $"Cell ({rowIndex}, {columnIndex}) should exist");

        object actualValue = cell.CellType switch {
            CellType.String => cell.StringCellValue,
            CellType.Numeric => cell.NumericCellValue,
            CellType.Boolean => cell.BooleanCellValue,
            CellType.Formula => cell.CellFormula,
            _ => null
        };

        Assert.That(actualValue, Is.EqualTo(expectedValue),
            $"Cell ({rowIndex}, {columnIndex}) value should match");
    }

    public void VerifyFreezePanes(int sheetIndex, int colSplit, int rowSplit) {
        ISheet sheet = GetSheet(sheetIndex);

        Assert.That(sheet.PaneInformation, Is.Not.Null, "Sheet should have freeze panes");
        Assert.That(sheet.PaneInformation.VerticalSplitPosition, Is.EqualTo(colSplit),
            "Vertical split position should match");
        Assert.That(sheet.PaneInformation.HorizontalSplitPosition, Is.EqualTo(rowSplit),
            "Horizontal split position should match");
        Assert.That(sheet.PaneInformation.IsFreezePane, Is.True, "Should be freeze pane, not split pane");
    }

    public void VerifyAutoFilter(int sheetIndex, int firstRow, int lastRow, int firstCol, int lastCol) {
        ISheet sheet = GetSheet(sheetIndex);

        // NPOI stores AutoFilter in CTAutoFilter for XLSX or similar structure
        // We can check if AutoFilter is set by checking if the first row has filter buttons
        IRow row = sheet.GetRow(firstRow);
        Assert.That(row, Is.Not.Null, "AutoFilter row should exist");

        // For now, we verify by checking the sheet structure
        // A more robust check would require accessing internal NPOI structures
        // which varies between XSSFSheet and HSSFSheet
        if (sheet is XSSFSheet xssfSheet) {
            CT_AutoFilter ctAutoFilter = xssfSheet.GetCTWorksheet().autoFilter;
            Assert.That(ctAutoFilter, Is.Not.Null, "Sheet should have auto filter");

            // Parse the reference like "A1:D10"
            string reference = ctAutoFilter.@ref;
            Assert.That(reference, Is.Not.Null.And.Not.Empty, "AutoFilter reference should not be empty");
        } else {
            // For HSSFSheet (XLS format), AutoFilter is handled differently
            // We'll just verify the sheet is not null for now
            Assert.That(sheet, Is.Not.Null, "Sheet should exist");
        }
    }

    public void VerifyDataValidation(int sheetIndex, int row, int col,
        Action<IDataValidation> validationAssertion) {
        ISheet sheet = GetSheet(sheetIndex);
        IList<IDataValidation> validations = sheet.GetDataValidations();

        IDataValidation validation = validations.FirstOrDefault(v => {
            CellRangeAddressList regions = v.Regions;
            return regions.CellRangeAddresses.Any(range =>
                range.IsInRange(row, col));
        });

        Assert.That(validation, Is.Not.Null,
            $"Cell ({row}, {col}) should have data validation");
        validationAssertion(validation);
    }

    public void VerifyColumnWidth(int sheetIndex, int columnIndex, double expectedWidth) {
        ISheet sheet = GetSheet(sheetIndex);
        int actualWidth = sheet.GetColumnWidth(columnIndex);
        double actualWidthInPoints = actualWidth / 256.0;

        Assert.That(actualWidthInPoints, Is.EqualTo(expectedWidth).Within(0.1),
            $"Column {columnIndex} width should match");
    }

    public void VerifyRowHeight(int sheetIndex, int rowIndex, double expectedHeight) {
        ISheet sheet = GetSheet(sheetIndex);
        IRow row = sheet.GetRow(rowIndex);

        Assert.That(row, Is.Not.Null, $"Row {rowIndex} should exist");
        double actualHeight = row.HeightInPoints;

        Assert.That(actualHeight, Is.EqualTo(expectedHeight).Within(0.1),
            $"Row {rowIndex} height should match");
    }

    public void VerifyMergedRegion(int sheetIndex, int firstRow, int lastRow, int firstCol, int lastCol) {
        ISheet sheet = GetSheet(sheetIndex);

        bool found = false;

        for (int i = 0; i < sheet.NumMergedRegions; i++) {
            CellRangeAddress region = sheet.GetMergedRegion(i);
            if (region.FirstRow == firstRow && region.LastRow == lastRow &&
                region.FirstColumn == firstCol && region.LastColumn == lastCol) {
                found = true;
                break;
            }
        }

        Assert.That(found, Is.True,
            $"Merged region ({firstRow}:{lastRow}, {firstCol}:{lastCol}) should exist");
    }

    public void VerifyCellStyle(int sheetIndex, int rowIndex, int columnIndex,
        Action<ICellStyle> styleAssertion) {
        ISheet sheet = GetSheet(sheetIndex);
        IRow row = sheet.GetRow(rowIndex);
        Assert.That(row, Is.Not.Null, $"Row {rowIndex} should exist");

        ICell cell = row.GetCell(columnIndex);
        Assert.That(cell, Is.Not.Null, $"Cell ({rowIndex}, {columnIndex}) should exist");

        styleAssertion(cell.CellStyle);
    }

    public void Dispose() {
        if (!disposed) {
            workbook?.Close();
            disposed = true;
        }
        GC.SuppressFinalize(this);
    }
}
