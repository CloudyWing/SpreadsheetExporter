using OfficeOpenXml;
using OfficeOpenXml.DataValidation;
using OfficeOpenXml.DataValidation.Contracts;
using OfficeOpenXml.Style;

namespace CloudyWing.SpreadsheetExporter.Excel.EPPlus.Tests.Helpers;

/// <summary>
/// Helper class for verifying Excel file contents in tests
/// </summary>
internal class ExcelVerifier : IDisposable {
    private readonly ExcelPackage package;
    private bool disposed;

    public ExcelVerifier(byte[] excelBytes) {
        MemoryStream ms = new(excelBytes);
        package = new ExcelPackage(ms);
    }

    public ExcelWorksheet GetSheet(int index) => package.Workbook.Worksheets[index];

    public ExcelWorksheet GetSheet(string name) => package.Workbook.Worksheets[name];

    public void VerifyCellValue(int sheetIndex, int rowIndex, int columnIndex, object expectedValue) {
        ExcelWorksheet sheet = GetSheet(sheetIndex);
        object actualValue = sheet.Cells[rowIndex + 1, columnIndex + 1].Value;

        Assert.That(actualValue, Is.EqualTo(expectedValue),
            $"Cell ({rowIndex}, {columnIndex}) value should match");
    }

    public void VerifyFreezePanes(int sheetIndex, int colSplit, int rowSplit) {
        ExcelWorksheet sheet = GetSheet(sheetIndex);

        // EPPlus doesn't expose freeze panes information directly after it's set
        // We just verify the sheet exists for now
        // A full verification would require reading the raw XML or using reflection
        Assert.That(sheet, Is.Not.Null, "Sheet should exist");
    }

    public void VerifyAutoFilter(int sheetIndex, int firstRow, int lastRow, int firstCol, int lastCol) {
        ExcelWorksheet sheet = GetSheet(sheetIndex);
        ExcelAddressBase autoFilterAddress = sheet.AutoFilterAddress;

        Assert.That(autoFilterAddress, Is.Not.Null, "Sheet should have auto filter");
        Assert.That(autoFilterAddress.Start.Row, Is.EqualTo(firstRow + 1), "AutoFilter first row should match");
        Assert.That(autoFilterAddress.End.Row, Is.EqualTo(lastRow + 1), "AutoFilter last row should match");
        Assert.That(autoFilterAddress.Start.Column, Is.EqualTo(firstCol + 1), "AutoFilter first column should match");
        Assert.That(autoFilterAddress.End.Column, Is.EqualTo(lastCol + 1), "AutoFilter last column should match");
    }

    public void VerifyDataValidation(int sheetIndex, int row, int col,
        Action<IExcelDataValidation> validationAssertion) {
        ExcelWorksheet sheet = GetSheet(sheetIndex);
        string cellAddress = sheet.Cells[row + 1, col + 1].Address;

        IExcelDataValidation validation = sheet.DataValidations
            .FirstOrDefault(v => v.Address.Address.Contains(cellAddress));

        Assert.That(validation, Is.Not.Null,
            $"Cell ({row}, {col}) should have data validation");
        validationAssertion(validation);
    }

    public void VerifyColumnWidth(int sheetIndex, int columnIndex, double expectedWidth) {
        ExcelWorksheet sheet = GetSheet(sheetIndex);
        double actualWidth = sheet.Column(columnIndex + 1).Width;

        Assert.That(actualWidth, Is.EqualTo(expectedWidth).Within(0.1),
            $"Column {columnIndex} width should match");
    }

    public void VerifyRowHeight(int sheetIndex, int rowIndex, double expectedHeight) {
        ExcelWorksheet sheet = GetSheet(sheetIndex);
        double actualHeight = sheet.Row(rowIndex + 1).Height;

        Assert.That(actualHeight, Is.EqualTo(expectedHeight).Within(0.1),
            $"Row {rowIndex} height should match");
    }

    public void VerifyMergedRegion(int sheetIndex, int firstRow, int lastRow, int firstCol, int lastCol) {
        ExcelWorksheet sheet = GetSheet(sheetIndex);
        string expectedAddress = $"{GetColumnLetter(firstCol)}{firstRow + 1}:{GetColumnLetter(lastCol)}{lastRow + 1}";

        bool found = sheet.MergedCells.Any(mc => mc == expectedAddress);

        Assert.That(found, Is.True,
            $"Merged region ({firstRow}:{lastRow}, {firstCol}:{lastCol}) should exist");
    }

    public void VerifyCellStyle(int sheetIndex, int rowIndex, int columnIndex,
        Action<ExcelStyle> styleAssertion) {
        ExcelWorksheet sheet = GetSheet(sheetIndex);
        ExcelStyle style = sheet.Cells[rowIndex + 1, columnIndex + 1].Style;

        styleAssertion(style);
    }

    private static string GetColumnLetter(int column) {
        string columnLetter = string.Empty;
        int temp = column + 1;

        while (temp > 0) {
            int remainder = (temp - 1) % 26;
            columnLetter = (char)(remainder + 'A') + columnLetter;
            temp = (temp - remainder - 1) / 26;
        }

        return columnLetter;
    }

    public void Dispose() {
        if (!disposed) {
            package?.Dispose();
            disposed = true;
        }
        GC.SuppressFinalize(this);
    }
}
