using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.Linq;
using OfficeOpenXml;
using OfficeOpenXml.DataValidation;
using OfficeOpenXml.DataValidation.Contracts;
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

    private static readonly Dictionary<DataValidationOperator, ExcelDataValidationOperator> validationOperatorMap = new() {
        [DataValidationOperator.Between] = ExcelDataValidationOperator.between,
        [DataValidationOperator.NotBetween] = ExcelDataValidationOperator.notBetween,
        [DataValidationOperator.Equal] = ExcelDataValidationOperator.equal,
        [DataValidationOperator.NotEqual] = ExcelDataValidationOperator.notEqual,
        [DataValidationOperator.GreaterThan] = ExcelDataValidationOperator.greaterThan,
        [DataValidationOperator.LessThan] = ExcelDataValidationOperator.lessThan,
        [DataValidationOperator.GreaterThanOrEqual] = ExcelDataValidationOperator.greaterThanOrEqual,
        [DataValidationOperator.LessThanOrEqual] = ExcelDataValidationOperator.lessThanOrEqual
    };

    /// <summary>
    /// Gets or sets a value indicating whether this instance is closed not implemented exception.
    /// </summary>
    /// <value>
    ///   <c>true</c> if this instance is closed not implemented exception; otherwise, <c>false</c>.</value>
    public bool IsClosedNotImplementedException { get; set; }

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
                if (endRow > maxRow)
                    maxRow = endRow;
                if (endCol > maxCol)
                    maxCol = endCol;
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

            DataValidation dataValidation = cell.GetDataValidation();
            if (dataValidation is not null) {
                SetDataValidation(sheet, range, dataValidation);
            }
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
            if (pair.Value <= Constants.AutoFitRowHeight) {
                row.CustomHeight = false;
            } else if (pair.Value == Constants.HiddenRow) {
                row.Hidden = true;
            } else if (pair.Value.HasValue) {
                row.Height = pair.Value.Value;
            }
        }
    }

    private void SetDataValidation(ExcelWorksheet sheet, ExcelRange range, DataValidation validation) {
        IExcelDataValidation excelValidation = CreateDataValidation(sheet, range, validation);

        excelValidation.AllowBlank = validation.IsBlankAllowed;
        excelValidation.ShowErrorMessage = validation.IsErrorAlertShown;
        excelValidation.ShowInputMessage = validation.IsInputPromptShown;

        if (!string.IsNullOrWhiteSpace(validation.ErrorTitle)) {
            excelValidation.ErrorTitle = validation.ErrorTitle;
        }

        if (!string.IsNullOrWhiteSpace(validation.ErrorMessage)) {
            excelValidation.Error = validation.ErrorMessage;
        }

        if (!string.IsNullOrWhiteSpace(validation.PromptTitle)) {
            excelValidation.PromptTitle = validation.PromptTitle;
        }

        if (!string.IsNullOrWhiteSpace(validation.PromptMessage)) {
            excelValidation.Prompt = validation.PromptMessage;
        }
    }

    private IExcelDataValidation CreateDataValidation(ExcelWorksheet sheet, ExcelRange range, DataValidation validation) {
        return validation.ValidationType switch {
            DataValidationType.List => CreateListValidation(sheet, range, validation),
            DataValidationType.Integer => CreateIntegerValidation(sheet, range, validation),
            DataValidationType.Decimal => CreateDecimalValidation(sheet, range, validation),
            DataValidationType.Date => CreateDateValidation(sheet, range, validation),
            DataValidationType.Time => CreateTimeValidation(sheet, range, validation),
            DataValidationType.TextLength => CreateTextLengthValidation(sheet, range, validation),
            DataValidationType.Custom => CreateCustomValidation(sheet, range, validation),
            _ => throw new ArgumentException($"Unsupported validation type: {validation.ValidationType}")
        };
    }

    private IExcelDataValidationList CreateListValidation(ExcelWorksheet sheet, ExcelRange range, DataValidation validation) {
        if (validation.ListItems is null || !validation.ListItems.Any()) {
            throw new ArgumentException("ListItems cannot be null or empty for List validation type.");
        }

        var listValidation = sheet.DataValidations.AddListValidation(range.Address);
        foreach (string item in validation.ListItems) {
            listValidation.Formula.Values.Add(item);
        }

        // EPPlus 4.x limitation: ShowDropDown property is not exposed in the public API
        // The dropdown arrow is always shown by default for list validation
        if (!validation.IsDropdownShown && !IsClosedNotImplementedException) {
            throw new NotImplementedException(
                "EPPlus 4.x does not support hiding the dropdown arrow for list validation. "
                + "Set IsClosedNotImplementedException = true to suppress this exception."
            );
        }

        return listValidation;
    }

    private IExcelDataValidationInt CreateIntegerValidation(ExcelWorksheet sheet, ExcelRange range, DataValidation validation) {
        if (!validation.Operator.HasValue) {
            throw new ArgumentException("Operator is required for Integer validation type.");
        }

        IExcelDataValidationInt intValidation = sheet.DataValidations.AddIntegerValidation(range.Address);
        SetValidationOperatorAndValue(intValidation, validation);

        return intValidation;
    }

    private IExcelDataValidationDecimal CreateDecimalValidation(ExcelWorksheet sheet, ExcelRange range, DataValidation validation) {
        if (!validation.Operator.HasValue) {
            throw new ArgumentException("Operator is required for Decimal validation type.");
        }

        IExcelDataValidationDecimal decimalValidation = sheet.DataValidations.AddDecimalValidation(range.Address);
        SetValidationOperatorAndValue(decimalValidation, validation);

        return decimalValidation;
    }

    private IExcelDataValidationDateTime CreateDateValidation(ExcelWorksheet sheet, ExcelRange range, DataValidation validation) {
        if (!validation.Operator.HasValue) {
            throw new ArgumentException("Operator is required for Date validation type.");
        }

        IExcelDataValidationDateTime dateValidation = sheet.DataValidations.AddDateTimeValidation(range.Address);
        SetValidationOperatorAndValue(dateValidation, validation);

        return dateValidation;
    }

    private IExcelDataValidationTime CreateTimeValidation(ExcelWorksheet sheet, ExcelRange range, DataValidation validation) {
        if (!validation.Operator.HasValue) {
            throw new ArgumentException("Operator is required for Time validation type.");
        }

        IExcelDataValidationTime timeValidation = sheet.DataValidations.AddTimeValidation(range.Address);
        SetValidationOperatorAndValue(timeValidation, validation);

        return timeValidation;
    }

    private IExcelDataValidationInt CreateTextLengthValidation(ExcelWorksheet sheet, ExcelRange range, DataValidation validation) {
        if (!validation.Operator.HasValue) {
            throw new ArgumentException("Operator is required for TextLength validation type.");
        }

        IExcelDataValidationInt textLengthValidation = sheet.DataValidations.AddTextLengthValidation(range.Address);
        SetValidationOperatorAndValue(textLengthValidation, validation);

        return textLengthValidation;
    }

    private IExcelDataValidationCustom CreateCustomValidation(ExcelWorksheet sheet, ExcelRange range, DataValidation validation) {
        if (string.IsNullOrWhiteSpace(validation.Formula)) {
            throw new ArgumentException("Formula is required for Custom validation type.");
        }

        IExcelDataValidationCustom customValidation = sheet.DataValidations.AddCustomValidation(range.Address);
        customValidation.Formula.ExcelFormula = validation.Formula;

        return customValidation;
    }

    private static void SetValidationOperatorAndValue(IExcelDataValidationInt validation, DataValidation dataValidation) {
        validation.Operator = ConvertOperator(dataValidation.Operator.Value);

        if (validation.Operator == ExcelDataValidationOperator.between
            || validation.Operator == ExcelDataValidationOperator.notBetween
        ) {
            validation.Formula.Value = Convert.ToInt32(dataValidation.Value1);
            validation.Formula2.Value = Convert.ToInt32(dataValidation.Value2);
        } else {
            validation.Formula.Value = Convert.ToInt32(dataValidation.Value1);
        }
    }

    private static void SetValidationOperatorAndValue(IExcelDataValidationDecimal validation, DataValidation dataValidation) {
        validation.Operator = ConvertOperator(dataValidation.Operator.Value);

        if (validation.Operator == ExcelDataValidationOperator.between
            || validation.Operator == ExcelDataValidationOperator.notBetween
        ) {
            validation.Formula.Value = Convert.ToDouble(dataValidation.Value1);
            validation.Formula2.Value = Convert.ToDouble(dataValidation.Value2);
        } else {
            validation.Formula.Value = Convert.ToDouble(dataValidation.Value1);
        }
    }

    private static void SetValidationOperatorAndValue(IExcelDataValidationDateTime validation, DataValidation dataValidation) {
        validation.Operator = ConvertOperator(dataValidation.Operator.Value);

        if (validation.Operator == ExcelDataValidationOperator.between
            || validation.Operator == ExcelDataValidationOperator.notBetween
        ) {
            validation.Formula.Value = Convert.ToDateTime(dataValidation.Value1);
            validation.Formula2.Value = Convert.ToDateTime(dataValidation.Value2);
        } else {
            validation.Formula.Value = Convert.ToDateTime(dataValidation.Value1);
        }
    }

    private static void SetValidationOperatorAndValue(IExcelDataValidationTime validation, DataValidation dataValidation) {
        validation.Operator = ConvertOperator(dataValidation.Operator.Value);

        if (validation.Operator == ExcelDataValidationOperator.between
            || validation.Operator == ExcelDataValidationOperator.notBetween
        ) {
            TimeSpan ts1 = (TimeSpan)dataValidation.Value1;
            TimeSpan ts2 = (TimeSpan)dataValidation.Value2;
            validation.Formula.Value = new ExcelTime((decimal)ts1.TotalDays);
            validation.Formula2.Value = new ExcelTime((decimal)ts2.TotalDays);
        } else {
            TimeSpan ts = (TimeSpan)dataValidation.Value1;
            validation.Formula.Value = new ExcelTime((decimal)ts.TotalDays);
        }
    }

    private static ExcelDataValidationOperator ConvertOperator(DataValidationOperator @operator) {
        if (!validationOperatorMap.TryGetValue(@operator, out ExcelDataValidationOperator result)) {
            throw new ArgumentException($"Unsupported operator: {@operator}");
        }

        return result;
    }
}
