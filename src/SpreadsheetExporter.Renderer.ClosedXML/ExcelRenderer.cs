using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using ClosedXML.Excel;

namespace CloudyWing.SpreadsheetExporter.Renderer.ClosedXML;

/// <summary>
/// Renders spreadsheet data to Excel (.xlsx) format using ClosedXML.
/// </summary>
public class ExcelRenderer : ISpreadsheetRenderer {
    /// <inheritdoc/>
    public string ContentType => "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

    /// <inheritdoc/>
    public string FileNameExtension => ".xlsx";

    /// <inheritdoc/>
    public byte[] Render(IEnumerable<SheetLayout> contexts) {
        ArgumentNullException.ThrowIfNull(contexts);

        SheetLayout[] contextArray = [.. contexts];

        using XLWorkbook workbook = new();

        SheetLayout? firstContext = contextArray.FirstOrDefault();
        if (firstContext?.DefaultFont.HasValue == true) {
            SetDefaultFont(workbook, firstContext.DefaultFont.Value);
        }

        foreach (SheetLayout context in contextArray) {
            ArgumentNullException.ThrowIfNull(context);

            IXLWorksheet worksheet = workbook.Worksheets.Add(context.SheetName);
            ConfigureWorksheet(worksheet, context);
            OnSheetCreated(worksheet, context);
        }

        using MemoryStream stream = new();
        workbook.SaveAs(stream);
        return stream.ToArray();
    }

    /// <summary>
    /// Called after a worksheet is created and configured.
    /// Override this method to perform additional customization using the ClosedXML API.
    /// </summary>
    /// <param name="worksheet">The ClosedXML worksheet.</param>
    /// <param name="context">The sheet context.</param>
    protected virtual void OnSheetCreated(IXLWorksheet worksheet, SheetLayout context) { }

    private static void ConfigureWorksheet(IXLWorksheet worksheet, SheetLayout context) {
        SetDefaultRowHeight(worksheet, context);
        SetSheetCells(worksheet, context.Cells);
        SetColumnWidths(worksheet, context.ColumnWidths);
        SetRowHeights(worksheet, context.RowHeights);
        SetPageSettings(worksheet, context.PageSettings);
        SetFreezePanes(worksheet, context.FreezePanes);
        SetAutoFilter(worksheet, context);
        SetProtection(worksheet, context);
    }

    private static void SetDefaultFont(XLWorkbook workbook, CellFont font) {
        SetFont(workbook.Style.Font, font);
    }

    private static void SetDefaultRowHeight(IXLWorksheet worksheet, SheetLayout context) {
        if (context.DefaultRowHeight.HasValue && context.DefaultRowHeight.Value > 0) {
            worksheet.RowHeight = context.DefaultRowHeight.Value;
        }
    }

    private static void SetSheetCells(IXLWorksheet worksheet, IReadOnlyList<Cell> cells) {
        foreach (Cell cell in cells) {
            int firstRow = cell.Point.Y + 1;
            int firstColumn = cell.Point.X + 1;
            int lastRow = firstRow + cell.Size.Height - 1;
            int lastColumn = firstColumn + cell.Size.Width - 1;

            IXLCell xlCell = worksheet.Cell(firstRow, firstColumn);
            IXLRange range = worksheet.Range(firstRow, firstColumn, lastRow, lastColumn);

            string? formula = cell.GetFormula();
            if (!string.IsNullOrWhiteSpace(formula)) {
                xlCell.FormulaA1 = formula;
            } else {
                SetCellValue(xlCell, cell.GetValue());
            }

            if (cell.Size.Width > 1 || cell.Size.Height > 1) {
                range.Merge();
            }

            SetCellStyle(range.Style, cell.GetCellStyle());

            if (cell.GetDataValidation() is DataValidation dataValidation) {
                SetDataValidation(range, dataValidation);
            }
        }
    }

    private static void SetCellStyle(IXLStyle xlStyle, CellStyle style) {
        xlStyle.Alignment.SetHorizontal(ToHorizontalAlignment(style.HorizontalAlignment));
        xlStyle.Alignment.SetVertical(ToVerticalAlignment(style.VerticalAlignment));
        xlStyle.Alignment.SetWrapText(style.WrapText);
        xlStyle.Protection.SetLocked(style.IsLocked);

        if (style.HasBorder) {
            xlStyle.Border.SetOutsideBorder(XLBorderStyleValues.Thin);
        }

        if (HasColor(style.BackgroundColor)) {
            xlStyle.Fill.SetBackgroundColor(XLColor.FromColor(style.BackgroundColor));
        }

        if (!string.IsNullOrWhiteSpace(style.DataFormat)) {
            xlStyle.NumberFormat.Format = style.DataFormat;
        }

        SetFont(xlStyle.Font, style.Font);
    }

    private static void SetFont(IXLFont xlFont, CellFont font) {
        if (!string.IsNullOrWhiteSpace(font.Name)) {
            xlFont.SetFontName(font.Name);
        }

        if (font.Size > 0) {
            xlFont.SetFontSize(font.Size);
        }

        if (HasColor(font.Color)) {
            xlFont.SetFontColor(XLColor.FromColor(font.Color));
        }

        if (font.Style.HasFlag(FontStyles.Bold)) {
            xlFont.SetBold();
        }

        if (font.Style.HasFlag(FontStyles.Italic)) {
            xlFont.SetItalic();
        }

        if (font.Style.HasFlag(FontStyles.Underline)) {
            xlFont.SetUnderline(XLFontUnderlineValues.Single);
        }

        if (font.Style.HasFlag(FontStyles.Strikeout)) {
            xlFont.SetStrikethrough();
        }
    }

    private static void SetColumnWidths(IXLWorksheet worksheet, IReadOnlyDictionary<int, double> widths) {
        foreach (KeyValuePair<int, double> pair in widths) {
            IXLColumn column = worksheet.Column(pair.Key + 1);

            if (pair.Value == Constants.HiddenColumn) {
                column.Hide();
            } else if (pair.Value == Constants.AutoFitColumnWidth) {
                column.AdjustToContents();
            } else if (pair.Value > 0) {
                column.Width = pair.Value;
            }
        }
    }

    private static void SetRowHeights(IXLWorksheet worksheet, IReadOnlyDictionary<int, double?> heights) {
        foreach (KeyValuePair<int, double?> pair in heights) {
            if (!pair.Value.HasValue) {
                continue;
            }

            IXLRow row = worksheet.Row(pair.Key + 1);

            if (pair.Value == Constants.HiddenRow) {
                row.Hide();
            } else if (pair.Value == Constants.AutoFitRowHeight) {
                row.AdjustToContents();
            } else if (pair.Value > 0) {
                row.Height = pair.Value.Value;
            }
        }
    }

    private static void SetPageSettings(IXLWorksheet worksheet, PageSettings settings) {
        ArgumentNullException.ThrowIfNull(settings);

        worksheet.PageSetup.PageOrientation = settings.PageOrientation == PageOrientation.Landscape
            ? XLPageOrientation.Landscape
            : XLPageOrientation.Portrait;
        worksheet.PageSetup.PaperSize = (XLPaperSize)settings.PaperSize.Value;
    }

    private static void SetFreezePanes(IXLWorksheet worksheet, Point? freezePanes) {
        if (!freezePanes.HasValue) {
            return;
        }

        if (freezePanes.Value.Y > 0) {
            worksheet.SheetView.FreezeRows(freezePanes.Value.Y);
        }

        if (freezePanes.Value.X > 0) {
            worksheet.SheetView.FreezeColumns(freezePanes.Value.X);
        }
    }

    private static void SetAutoFilter(IXLWorksheet worksheet, SheetLayout context) {
        if (!context.IsAutoFilterEnabled) {
            return;
        }

        if (worksheet.RangeUsed() is IXLRange usedRange) {
            usedRange.SetAutoFilter();
        }
    }

    private static void SetProtection(IXLWorksheet worksheet, SheetLayout context) {
        if (!context.IsProtected) {
            return;
        }

        worksheet.Protect(context.Password, XLProtectionAlgorithm.Algorithm.SHA512);
    }

    private static void SetDataValidation(IXLRange range, DataValidation validation) {
        IXLDataValidation xlValidation = range.CreateDataValidation();

        switch (validation.ValidationType) {
            case DataValidationType.List:
                string listString = string.Join(",", validation.ListItems ?? []);
                xlValidation.List($"\"{listString}\"", validation.IsDropdownShown);
                break;
            case DataValidationType.Integer:
                SetWholeNumberValidation(xlValidation, validation);
                break;
            case DataValidationType.Decimal:
                SetDecimalValidation(xlValidation, validation);
                break;
            case DataValidationType.Date:
                SetDateValidation(xlValidation, validation);
                break;
            case DataValidationType.Time:
                SetTimeValidation(xlValidation, validation);
                break;
            case DataValidationType.TextLength:
                SetTextLengthValidation(xlValidation, validation);
                break;
            case DataValidationType.Custom:
                xlValidation.Custom(NormalizeCustomFormula(validation.Formula));
                break;
            default:
                throw new NotSupportedException(
                    $"Unsupported data validation type: {validation.ValidationType}."
                );
        }

        xlValidation.IgnoreBlanks = validation.IsBlankAllowed;
        xlValidation.InCellDropdown = validation.IsDropdownShown;
        xlValidation.ShowErrorMessage = validation.IsErrorAlertShown;
        xlValidation.ShowInputMessage = validation.IsInputPromptShown;

        if (!string.IsNullOrWhiteSpace(validation.ErrorTitle)) {
            xlValidation.ErrorTitle = validation.ErrorTitle;
        }

        if (!string.IsNullOrWhiteSpace(validation.ErrorMessage)) {
            xlValidation.ErrorMessage = validation.ErrorMessage;
        }

        if (!string.IsNullOrWhiteSpace(validation.PromptTitle)) {
            xlValidation.InputTitle = validation.PromptTitle;
        }

        if (!string.IsNullOrWhiteSpace(validation.PromptMessage)) {
            xlValidation.InputMessage = validation.PromptMessage;
        }
    }

    private static void SetWholeNumberValidation(IXLDataValidation xlValidation, DataValidation validation) {
        if (!string.IsNullOrWhiteSpace(validation.Formula)) {
            SetFormulaValidation(xlValidation, validation, XLAllowedValues.WholeNumber);
            return;
        }

        switch (validation.Operator ?? DataValidationOperator.Equal) {
            case DataValidationOperator.Between:
                xlValidation.WholeNumber.Between(ToInt32(validation.Value1), ToInt32(validation.Value2));
                break;
            case DataValidationOperator.NotBetween:
                xlValidation.WholeNumber.NotBetween(ToInt32(validation.Value1), ToInt32(validation.Value2));
                break;
            case DataValidationOperator.Equal:
                xlValidation.WholeNumber.EqualTo(ToInt32(validation.Value1));
                break;
            case DataValidationOperator.NotEqual:
                xlValidation.WholeNumber.NotEqualTo(ToInt32(validation.Value1));
                break;
            case DataValidationOperator.GreaterThan:
                xlValidation.WholeNumber.GreaterThan(ToInt32(validation.Value1));
                break;
            case DataValidationOperator.LessThan:
                xlValidation.WholeNumber.LessThan(ToInt32(validation.Value1));
                break;
            case DataValidationOperator.GreaterThanOrEqual:
                xlValidation.WholeNumber.EqualOrGreaterThan(ToInt32(validation.Value1));
                break;
            case DataValidationOperator.LessThanOrEqual:
                xlValidation.WholeNumber.EqualOrLessThan(ToInt32(validation.Value1));
                break;
            default:
                throw new NotSupportedException(
                    $"Unsupported data validation operator: {validation.Operator}."
                );
        }
    }

    private static void SetDecimalValidation(IXLDataValidation xlValidation, DataValidation validation) {
        if (!string.IsNullOrWhiteSpace(validation.Formula)) {
            SetFormulaValidation(xlValidation, validation, XLAllowedValues.Decimal);
            return;
        }

        switch (validation.Operator ?? DataValidationOperator.Equal) {
            case DataValidationOperator.Between:
                xlValidation.Decimal.Between(ToDouble(validation.Value1), ToDouble(validation.Value2));
                break;
            case DataValidationOperator.NotBetween:
                xlValidation.Decimal.NotBetween(ToDouble(validation.Value1), ToDouble(validation.Value2));
                break;
            case DataValidationOperator.Equal:
                xlValidation.Decimal.EqualTo(ToDouble(validation.Value1));
                break;
            case DataValidationOperator.NotEqual:
                xlValidation.Decimal.NotEqualTo(ToDouble(validation.Value1));
                break;
            case DataValidationOperator.GreaterThan:
                xlValidation.Decimal.GreaterThan(ToDouble(validation.Value1));
                break;
            case DataValidationOperator.LessThan:
                xlValidation.Decimal.LessThan(ToDouble(validation.Value1));
                break;
            case DataValidationOperator.GreaterThanOrEqual:
                xlValidation.Decimal.EqualOrGreaterThan(ToDouble(validation.Value1));
                break;
            case DataValidationOperator.LessThanOrEqual:
                xlValidation.Decimal.EqualOrLessThan(ToDouble(validation.Value1));
                break;
            default:
                throw new NotSupportedException(
                    $"Unsupported data validation operator: {validation.Operator}."
                );
        }
    }

    private static void SetDateValidation(IXLDataValidation xlValidation, DataValidation validation) {
        if (!string.IsNullOrWhiteSpace(validation.Formula)) {
            SetFormulaValidation(xlValidation, validation, XLAllowedValues.Date);
            return;
        }

        switch (validation.Operator ?? DataValidationOperator.Equal) {
            case DataValidationOperator.Between:
                xlValidation.Date.Between(ToDateTime(validation.Value1), ToDateTime(validation.Value2));
                break;
            case DataValidationOperator.NotBetween:
                xlValidation.Date.NotBetween(ToDateTime(validation.Value1), ToDateTime(validation.Value2));
                break;
            case DataValidationOperator.Equal:
                xlValidation.Date.EqualTo(ToDateTime(validation.Value1));
                break;
            case DataValidationOperator.NotEqual:
                xlValidation.Date.NotEqualTo(ToDateTime(validation.Value1));
                break;
            case DataValidationOperator.GreaterThan:
                xlValidation.Date.GreaterThan(ToDateTime(validation.Value1));
                break;
            case DataValidationOperator.LessThan:
                xlValidation.Date.LessThan(ToDateTime(validation.Value1));
                break;
            case DataValidationOperator.GreaterThanOrEqual:
                xlValidation.Date.EqualOrGreaterThan(ToDateTime(validation.Value1));
                break;
            case DataValidationOperator.LessThanOrEqual:
                xlValidation.Date.EqualOrLessThan(ToDateTime(validation.Value1));
                break;
            default:
                throw new NotSupportedException(
                    $"Unsupported data validation operator: {validation.Operator}."
                );
        }
    }

    private static void SetTimeValidation(IXLDataValidation xlValidation, DataValidation validation) {
        if (!string.IsNullOrWhiteSpace(validation.Formula)) {
            SetFormulaValidation(xlValidation, validation, XLAllowedValues.Time);
            return;
        }

        switch (validation.Operator ?? DataValidationOperator.Equal) {
            case DataValidationOperator.Between:
                xlValidation.Time.Between(ToTimeSpan(validation.Value1), ToTimeSpan(validation.Value2));
                break;
            case DataValidationOperator.NotBetween:
                xlValidation.Time.NotBetween(ToTimeSpan(validation.Value1), ToTimeSpan(validation.Value2));
                break;
            case DataValidationOperator.Equal:
                xlValidation.Time.EqualTo(ToTimeSpan(validation.Value1));
                break;
            case DataValidationOperator.NotEqual:
                xlValidation.Time.NotEqualTo(ToTimeSpan(validation.Value1));
                break;
            case DataValidationOperator.GreaterThan:
                xlValidation.Time.GreaterThan(ToTimeSpan(validation.Value1));
                break;
            case DataValidationOperator.LessThan:
                xlValidation.Time.LessThan(ToTimeSpan(validation.Value1));
                break;
            case DataValidationOperator.GreaterThanOrEqual:
                xlValidation.Time.EqualOrGreaterThan(ToTimeSpan(validation.Value1));
                break;
            case DataValidationOperator.LessThanOrEqual:
                xlValidation.Time.EqualOrLessThan(ToTimeSpan(validation.Value1));
                break;
            default:
                throw new NotSupportedException(
                    $"Unsupported data validation operator: {validation.Operator}."
                );
        }
    }

    private static void SetTextLengthValidation(IXLDataValidation xlValidation, DataValidation validation) {
        if (!string.IsNullOrWhiteSpace(validation.Formula)) {
            SetFormulaValidation(xlValidation, validation, XLAllowedValues.TextLength);
            return;
        }

        switch (validation.Operator ?? DataValidationOperator.Equal) {
            case DataValidationOperator.Between:
                xlValidation.TextLength.Between(ToInt32(validation.Value1), ToInt32(validation.Value2));
                break;
            case DataValidationOperator.NotBetween:
                xlValidation.TextLength.NotBetween(ToInt32(validation.Value1), ToInt32(validation.Value2));
                break;
            case DataValidationOperator.Equal:
                xlValidation.TextLength.EqualTo(ToInt32(validation.Value1));
                break;
            case DataValidationOperator.NotEqual:
                xlValidation.TextLength.NotEqualTo(ToInt32(validation.Value1));
                break;
            case DataValidationOperator.GreaterThan:
                xlValidation.TextLength.GreaterThan(ToInt32(validation.Value1));
                break;
            case DataValidationOperator.LessThan:
                xlValidation.TextLength.LessThan(ToInt32(validation.Value1));
                break;
            case DataValidationOperator.GreaterThanOrEqual:
                xlValidation.TextLength.EqualOrGreaterThan(ToInt32(validation.Value1));
                break;
            case DataValidationOperator.LessThanOrEqual:
                xlValidation.TextLength.EqualOrLessThan(ToInt32(validation.Value1));
                break;
            default:
                throw new NotSupportedException(
                    $"Unsupported data validation operator: {validation.Operator}."
                );
        }
    }

    private static void SetFormulaValidation(
        IXLDataValidation xlValidation, DataValidation validation, XLAllowedValues allowedValues
    ) {
        xlValidation.AllowedValues = allowedValues;
        xlValidation.Operator = ToValidationOperator(validation.Operator ?? DataValidationOperator.Equal);

        string? formula = EnsureFormulaPrefix(validation.Formula);
        switch (validation.Operator ?? DataValidationOperator.Equal) {
            case DataValidationOperator.Between:
            case DataValidationOperator.NotBetween:
                xlValidation.MinValue = formula;
                xlValidation.MaxValue = EnsureFormulaPrefix(ToInvariantString(validation.Value2));
                break;
            default:
                xlValidation.Value = formula;
                break;
        }
    }

    private static void SetCellValue(IXLCell cell, object? value) {
        cell.Value = value is null
            ? Blank.Value
            : XLCellValue.FromObject(value, CultureInfo.InvariantCulture);
    }

    private static XLAlignmentHorizontalValues ToHorizontalAlignment(HorizontalAlignment alignment) {
        return alignment switch {
            HorizontalAlignment.Left => XLAlignmentHorizontalValues.Left,
            HorizontalAlignment.Center => XLAlignmentHorizontalValues.Center,
            HorizontalAlignment.Right => XLAlignmentHorizontalValues.Right,
            HorizontalAlignment.Justify => XLAlignmentHorizontalValues.Justify,
            _ => XLAlignmentHorizontalValues.General
        };
    }

    private static XLAlignmentVerticalValues ToVerticalAlignment(VerticalAlignment alignment) {
        return alignment switch {
            VerticalAlignment.Middle => XLAlignmentVerticalValues.Center,
            VerticalAlignment.Bottom => XLAlignmentVerticalValues.Bottom,
            _ => XLAlignmentVerticalValues.Top
        };
    }

    private static XLOperator ToValidationOperator(DataValidationOperator validationOperator) {
        return validationOperator switch {
            DataValidationOperator.Between => XLOperator.Between,
            DataValidationOperator.NotBetween => XLOperator.NotBetween,
            DataValidationOperator.Equal => XLOperator.EqualTo,
            DataValidationOperator.NotEqual => XLOperator.NotEqualTo,
            DataValidationOperator.GreaterThan => XLOperator.GreaterThan,
            DataValidationOperator.LessThan => XLOperator.LessThan,
            DataValidationOperator.GreaterThanOrEqual => XLOperator.EqualOrGreaterThan,
            DataValidationOperator.LessThanOrEqual => XLOperator.EqualOrLessThan,
            _ => throw new NotSupportedException(
                $"Unsupported data validation operator: {validationOperator}."
            )
        };
    }

    private static int ToInt32(object? value) {
        return Convert.ToInt32(value, CultureInfo.InvariantCulture);
    }

    private static double ToDouble(object? value) {
        return Convert.ToDouble(value, CultureInfo.InvariantCulture);
    }

    private static DateTime ToDateTime(object? value) {
        return value switch {
            DateTime dateTime => dateTime,
            string text => DateTime.Parse(text, CultureInfo.InvariantCulture),
            _ => Convert.ToDateTime(value, CultureInfo.InvariantCulture)
        };
    }

    private static TimeSpan ToTimeSpan(object? value) {
        return value switch {
            TimeSpan timeSpan => timeSpan,
            string text => TimeSpan.Parse(text, CultureInfo.InvariantCulture),
            _ => TimeSpan.FromTicks(Convert.ToInt64(value, CultureInfo.InvariantCulture))
        };
    }

    private static string NormalizeCustomFormula(string? formula) {
        if (string.IsNullOrWhiteSpace(formula)) {
            throw new ArgumentException("Custom validation formula cannot be null or whitespace.", nameof(formula));
        }

        return formula.Trim().TrimStart('=');
    }

    private static string? EnsureFormulaPrefix(string? value) {
        if (string.IsNullOrWhiteSpace(value)) {
            return value;
        }

        string normalized = value.Trim();
        return normalized.StartsWith('=')
            ? normalized
            : $"={normalized}";
    }

    private static string? ToInvariantString(object? value) {
        return value switch {
            null => null,
            IFormattable formattable => formattable.ToString(null, CultureInfo.InvariantCulture),
            _ => value.ToString()
        };
    }

    private static bool HasColor(Color color) {
        return !color.IsEmpty;
    }
}
