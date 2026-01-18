using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using NPOI.HSSF.UserModel;
using NPOI.HSSF.Util;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;

namespace CloudyWing.SpreadsheetExporter.Excel.NPOI;

/// <summary>
/// The excel exporter, using npoi.
/// </summary>
/// <seealso cref="ExporterBase" />
/// <param name="excelFormat">The excel format.</param>
public sealed class ExcelExporter(ExcelFormat excelFormat = ExcelFormat.OfficeOpenXmlDocument) : ExporterBase {
    private IWorkbook workbook;
    private readonly IDictionary<CellStyle, ICellStyle> cellStyles = new Dictionary<CellStyle, ICellStyle>();
    private readonly IDictionary<CellFont, IFont> fonts = new Dictionary<CellFont, IFont>();
    private static readonly IDictionary<ExcelFormat, string> filenameExtensionMap = new Dictionary<ExcelFormat, string> {
        [ExcelFormat.ExcelBinaryFileFormat] = ".xls",
        [ExcelFormat.OfficeOpenXmlDocument] = ".xlsx"
    };

    private static readonly IDictionary<ExcelFormat, string> contentTypeMap = new Dictionary<ExcelFormat, string> {
        [ExcelFormat.ExcelBinaryFileFormat] = "application/vnd.ms-excel",
        [ExcelFormat.OfficeOpenXmlDocument] = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    };

    private static readonly IDictionary<DataValidationOperator, int> validationOperatorMap = new Dictionary<DataValidationOperator, int> {
        [DataValidationOperator.Between] = OperatorType.BETWEEN,
        [DataValidationOperator.NotBetween] = OperatorType.NOT_BETWEEN,
        [DataValidationOperator.Equal] = OperatorType.EQUAL,
        [DataValidationOperator.NotEqual] = OperatorType.NOT_EQUAL,
        [DataValidationOperator.GreaterThan] = OperatorType.GREATER_THAN,
        [DataValidationOperator.LessThan] = OperatorType.LESS_THAN,
        [DataValidationOperator.GreaterThanOrEqual] = OperatorType.GREATER_OR_EQUAL,
        [DataValidationOperator.LessThanOrEqual] = OperatorType.LESS_OR_EQUAL
    };

    private readonly Dictionary<HorizontalAlignment, global::NPOI.SS.UserModel.HorizontalAlignment> horizontalAlignmentMap = new() {
        [HorizontalAlignment.General] = global::NPOI.SS.UserModel.HorizontalAlignment.General,
        [HorizontalAlignment.Left] = global::NPOI.SS.UserModel.HorizontalAlignment.Left,
        [HorizontalAlignment.Center] = global::NPOI.SS.UserModel.HorizontalAlignment.Center,
        [HorizontalAlignment.Right] = global::NPOI.SS.UserModel.HorizontalAlignment.Right,
        [HorizontalAlignment.Justify] = global::NPOI.SS.UserModel.HorizontalAlignment.Justify
    };

    /// <summary>
    /// Excel uses a unit of 1/256th of a character width for column width.
    /// </summary>
    private const int ExcelColumnWidthUnit = 256;

    private readonly Dictionary<VerticalAlignment, global::NPOI.SS.UserModel.VerticalAlignment> verticalAlignmentMap = new() {
        [VerticalAlignment.Top] = global::NPOI.SS.UserModel.VerticalAlignment.Top,
        [VerticalAlignment.Middle] = global::NPOI.SS.UserModel.VerticalAlignment.Center,
        [VerticalAlignment.Bottom] = global::NPOI.SS.UserModel.VerticalAlignment.Bottom
    };

    private readonly object thisLock = new();

    /// <summary>
    /// Gets or sets the excel format.
    /// </summary>
    /// <value>
    /// The excel format.
    /// </value>
    public ExcelFormat ExcelFormat { get; set; } = excelFormat;

    /// <summary>
    /// Gets a value indicating whether this instance is office open XML document.
    /// </summary>
    /// <value>
    ///   <c>true</c> if this instance is office open XML document; otherwise, <c>false</c>.</value>
    public bool IsOfficeOpenXmlDocument => ExcelFormat == ExcelFormat.OfficeOpenXmlDocument;

    /// <summary>
    /// Gets or sets a value indicating whether this instance is closed not implemented exception.
    /// </summary>
    /// <value>
    ///   <c>true</c> if this instance is closed not implemented exception; otherwise, <c>false</c>.</value>
    public bool IsClosedNotImplementedException { get; set; }

    /// <inheritdoc/>
    public override string ContentType => contentTypeMap[ExcelFormat];

    /// <inheritdoc/>
    public override string FileNameExtension => filenameExtensionMap[ExcelFormat];

    /// <inheritdoc/>
    protected override byte[] ExecuteExport(IEnumerable<SheeterContext> contexts) {
        lock (thisLock) {
            // 因為 ParseCellStyle 和 ParseFont 會用到，所以用 field 處理
            workbook = IsOfficeOpenXmlDocument ? new XSSFWorkbook() : new HSSFWorkbook();

            try {
                if (DefaultFont.HasValue) {
                    SetDefaultFont(workbook, DefaultFont.Value);
                }

                foreach (SheeterContext context in contexts) {
                    CreateSheet(context);
                }

                using MemoryStream ms = new();

                if (HasPassword) {
                    // NPOI 目前不支援在 xlsx 格式使用密碼保護 Workbook
                    if (IsOfficeOpenXmlDocument) {
                        if (!IsClosedNotImplementedException) {
                            throw new NotImplementedException("If no other packages are installed, NPOI currently does not support the output of xlsx file with passwords.");
                        }
                    } else {
                        HSSFWorkbook wb = (HSSFWorkbook)workbook;
                        // 因為 NPOI 的 Bug，在 2.5.5 版以前要先 call InternalWorkbook.WriteAccess 才可以正常，後續不知是否有修正
                        _ = wb.InternalWorkbook.WriteAccess;
                        wb.WriteProtectWorkbook(Password, "");
                    }
                }

                workbook.Write(ms);

                return ms.ToArray();
            } finally {
                workbook.Close();
                cellStyles.Clear();
                fonts.Clear();
                workbook = null;
            }
        }
    }

    private void SetDefaultFont(IWorkbook workbook, CellFont font) {
        IFont defaultFont = workbook.GetFontAt(0);
        if (!string.IsNullOrWhiteSpace(font.Name)) {
            defaultFont.FontName = font.Name;
        }

        // NPOI 可能只有 Size 可以生效
        if (font.Size != 0) {
            defaultFont.FontHeightInPoints = font.Size;
        }

        if (font.Color != Color.Empty) {
            if (defaultFont is HSSFFont) {
                defaultFont.Color = ParseColor(font.Color);
            } else {
                ((XSSFFont)defaultFont).SetColor(new XSSFColor(font.Color));
            }
        }
        defaultFont.IsBold = (font.Style & FontStyles.IsBold) == FontStyles.IsBold;
        defaultFont.IsItalic = (font.Style & FontStyles.IsItalic) == FontStyles.IsItalic;
        if ((font.Style & FontStyles.HasUnderline) == FontStyles.HasUnderline) {
            defaultFont.Underline = FontUnderlineType.Single;
        }
        defaultFont.IsStrikeout = (font.Style & FontStyles.IsStrikeout) == FontStyles.IsStrikeout;
    }

    private void CreateSheet(SheeterContext context) {
        ISheet sheet = workbook.CreateSheet(context.SheetName);
        if (context.DefaultRowHeight.HasValue) {
            sheet.DefaultRowHeightInPoints = (float)context.DefaultRowHeight;
        }
        SetSheetCells(sheet, context.Cells);
        SetSheetColumnWidths(sheet, context.ColumnWidths);
        SetSheetRowHeights(sheet, context.RowHeights);

        if (context.FreezePanes.HasValue) {
            // NPOI CreateFreezePane: colSplit and rowSplit specify the first unfrozen column/row
            sheet.CreateFreezePane(context.FreezePanes.Value.X, context.FreezePanes.Value.Y);
        }

        if (context.IsAutoFilterEnabled && context.Cells.Count > 0) {
            // Calculate the range for auto filter based on all cells
            int maxRow = 0;
            int maxCol = 0;
            foreach (Cell cell in context.Cells) {
                int endRow = cell.Point.Y + cell.Size.Height - 1;
                int endCol = cell.Point.X + cell.Size.Width - 1;
                if (endRow > maxRow) {
                    maxRow = endRow;
                }
                if (endCol > maxCol) {
                    maxCol = endCol;
                }
            }
            // NPOI uses 0-based index for SetAutoFilter
            sheet.SetAutoFilter(new CellRangeAddress(0, maxRow, 0, maxCol));
        }

        if (context.IsProtected) {
            sheet.ProtectSheet(context.Password);
        }

        sheet.PrintSetup.Landscape = context.PageSettings.PageOrientation == PageOrientation.Landscape;
        sheet.PrintSetup.PaperSize = (short)context.PageSettings.PaperSize;

        sheet.ForceFormulaRecalculation = true;

        OnSheetCreated(new SheetCreatedEventArgs(sheet, context));
    }

    private void SetSheetCells(ISheet sheet, IReadOnlyList<Cell> cells) {
        foreach (Cell cell in cells) {
            IRow excelRow = sheet.GetRow(cell.Point.Y) ?? sheet.CreateRow(cell.Point.Y);
            ICell excelCell = excelRow.GetCell(cell.Point.X) ?? excelRow.CreateCell(cell.Point.X);
            string formula = cell.GetFormula();

            if (string.IsNullOrWhiteSpace(formula)) {
                SetValueToCell(excelCell, cell.GetValue());
            } else {
                excelCell.CellFormula = formula;
            }

            excelCell.CellStyle = ParseCellStyle(cell.GetCellStyle());

            DataValidation dataValidation = cell.GetDataValidation();
            if (dataValidation is not null) {
                SetDataValidation(sheet, cell.Point, cell.Size, dataValidation);
            }

            if (cell.Size.Width > 1 || cell.Size.Height > 1) {
                MergedRegion(
                    sheet,
                    cell.Point.X, cell.Point.X + cell.Size.Width - 1,
                    cell.Point.Y, cell.Point.Y + cell.Size.Height - 1
                );
            }
        }
    }

    private ICellStyle ParseCellStyle(CellStyle cellStyle) {
        ICellStyle excelCellStyle;

        // 聽說 CellStyle 建立太多，可能會出現問題，所以如果有已存在的，就直接使用
        if (cellStyles.ContainsKey(cellStyle)) {
            return cellStyles[cellStyle];
        }
        excelCellStyle = workbook.CreateCellStyle();

        excelCellStyle.Alignment = horizontalAlignmentMap[cellStyle.HorizontalAlignment];
        excelCellStyle.VerticalAlignment = verticalAlignmentMap[cellStyle.VerticalAlignment];

        if (cellStyle.HasBorder) {
            excelCellStyle.BorderBottom = BorderStyle.Thin;
            excelCellStyle.BorderLeft = BorderStyle.Thin;
            excelCellStyle.BorderRight = BorderStyle.Thin;
            excelCellStyle.BorderTop = BorderStyle.Thin;
        }

        excelCellStyle.WrapText = cellStyle.WrapText;
        if (cellStyle.BackgroundColor != Color.Empty) {
            excelCellStyle.FillPattern = FillPattern.SolidForeground;
            if (IsOfficeOpenXmlDocument) {
                ((XSSFCellStyle)excelCellStyle).SetFillForegroundColor(new XSSFColor(cellStyle.BackgroundColor));
            } else {
                excelCellStyle.FillForegroundColor = ParseColor(cellStyle.BackgroundColor);
            }
        }

        if (cellStyle.Font != CellFont.Empty) {
            excelCellStyle.SetFont(ParseFont(cellStyle.Font));
        }

        if (!string.IsNullOrWhiteSpace(cellStyle.DataFormat)) {
            excelCellStyle.DataFormat = ParseDataFormat(cellStyle.DataFormat);
        }

        excelCellStyle.IsLocked = cellStyle.IsLocked;

        cellStyles.Add(cellStyle, excelCellStyle);

        return excelCellStyle;
    }

    private IFont ParseFont(CellFont font) {
        // 聽說 CellStyle 建立太多，可能會出現問題，所以如果有已存在的，就直接使用
        if (fonts.ContainsKey(font)) {
            return fonts[font];
        }

        IFont excelFont = workbook.CreateFont();

        if (!string.IsNullOrWhiteSpace(font.Name)) {
            excelFont.FontName = font.Name;
        }

        if (font.Size != 0) {
            excelFont.FontHeightInPoints = font.Size;
        }

        if (font.Color != Color.Empty) {
            if (excelFont is HSSFFont) {
                excelFont.Color = ParseColor(font.Color);
            } else {
                ((XSSFFont)excelFont).SetColor(new XSSFColor(font.Color));
            }
        }
        excelFont.IsBold = (font.Style & FontStyles.IsBold) == FontStyles.IsBold;
        excelFont.IsItalic = (font.Style & FontStyles.IsItalic) == FontStyles.IsItalic;
        if ((font.Style & FontStyles.HasUnderline) == FontStyles.HasUnderline) {
            excelFont.Underline = FontUnderlineType.Single;
        }
        excelFont.IsStrikeout = (font.Style & FontStyles.IsStrikeout) == FontStyles.IsStrikeout;

        fonts.Add(font, excelFont);

        return excelFont;
    }

    private short ParseColor(Color color) {
        // HSSFColor 無法用 Color 設定顏色，所以如果 xls 版本就使用 Reflection 找出對應顏色 Index
        // 找出此 Type 下公用的 class
        Type[] colorTypes = typeof(HSSFColor).GetNestedTypes(BindingFlags.Public);

        foreach (Type type in colorTypes) {
            if (!TryGetColorIndex(type, color, out short colorIndex)) {
                continue;
            }

            return colorIndex;
        }

        return 0;
    }

    private static bool TryGetColorIndex(Type colorType, Color color, out short colorIndex) {
        colorIndex = 0;

        // Automatic 這個顏色沒有 Triplet
        FieldInfo tripletField = colorType.GetField("Triplet", BindingFlags.Public | BindingFlags.Static);
        if (tripletField is null) {
            return false;
        }

        byte[] rgb = tripletField.GetValue(null) as byte[];
        if (rgb is null || rgb[0] != color.R || rgb[1] != color.G || rgb[2] != color.B) {
            return false;
        }

        FieldInfo indexField = colorType.GetField("Index", BindingFlags.Public | BindingFlags.Static);
        colorIndex = (short)indexField.GetValue(null);
        return true;
    }

    private short ParseDataFormat(string formatStr) {
        IDataFormat dataFormat = workbook.CreateDataFormat();
        return dataFormat.GetFormat(formatStr);
    }

    private void MergedRegion(ISheet sheet, int firstColnum, int lastColumn, int firstRow, int lastRow) {
        // 合併儲存格，並把第一個儲存格的樣式複製到其他儲存格
        ICellStyle cellStyle = sheet.GetRow(firstRow).GetCell(firstColnum).CellStyle;
        for (int column = firstColnum; column <= lastColumn; column++) {
            for (int row = firstRow; row <= lastRow; row++) {
                if (column == firstColnum && row == firstRow) {
                    continue;
                }
                (sheet.GetRow(row) ?? sheet.CreateRow(row)).CreateCell(column).CellStyle = cellStyle;
            }
        }

        sheet.AddMergedRegion(
            new CellRangeAddress(firstRow, lastRow, firstColnum, lastColumn)
        );
    }

    private void SetValueToCell(ICell cell, object value) {
        if (value is null) {
            cell.SetCellValue("");
        } else if (value is bool _bool) {
            cell.SetCellValue(_bool);
        } else if (value is DateTime _dateTime) {
            cell.SetCellValue(_dateTime);
        } else if (value is double _double) {
            cell.SetCellValue(_double);
        } else {
            if (value is not string and IConvertible) {
                try {
                    cell.SetCellValue(Convert.ToDouble(value));
                } catch (Exception ex) when (ex is FormatException or InvalidCastException or OverflowException) {
                    cell.SetCellValue(value.ToString());
                }
            } else {
                cell.SetCellValue(value.ToString());
            }
        }
    }

    private void SetSheetColumnWidths(ISheet sheet, IReadOnlyDictionary<int, double> columnWidths) {
        foreach (KeyValuePair<int, double> pair in columnWidths) {
            if (pair.Value <= Constants.AutoFitColumnWidth) {
                sheet.AutoSizeColumn(pair.Key);
            } else if (pair.Value == Constants.HiddenColumn) {
                sheet.SetColumnHidden(pair.Key, true);
            } else {
                sheet.SetColumnWidth(pair.Key, (short)(pair.Value * ExcelColumnWidthUnit));
            }
        }
    }

    private void SetSheetRowHeights(ISheet sheet, IReadOnlyDictionary<int, double?> rowHeights) {
        foreach (KeyValuePair<int, double?> pair in rowHeights) {
            IRow row = sheet.GetRow(pair.Key) ?? sheet.CreateRow(pair.Key);
            if (pair.Value <= Constants.AutoFitRowHeight) {
                row.Height = -1;
            } else if (pair.Value == Constants.HiddenRow) {
                row.ZeroHeight = true;
            } else if (pair.Value.HasValue) {
                row.HeightInPoints = (float)pair.Value.Value;
            }
        }
    }

    private void SetDataValidation(ISheet sheet, Point point, Size size, DataValidation validation) {
        CellRangeAddressList addressList = new(
            point.Y,
            point.Y + size.Height - 1,
            point.X,
            point.X + size.Width - 1
        );

        IDataValidationHelper validationHelper = sheet.GetDataValidationHelper();
        IDataValidationConstraint constraint = CreateValidationConstraint(validationHelper, validation);
        IDataValidation dataValidation = validationHelper.CreateValidation(constraint, addressList);

        dataValidation.EmptyCellAllowed = validation.IsBlankAllowed;
        dataValidation.ShowErrorBox = validation.IsErrorAlertShown;
        dataValidation.ShowPromptBox = validation.IsInputPromptShown;

        if (!string.IsNullOrWhiteSpace(validation.ErrorTitle)
            || !string.IsNullOrWhiteSpace(validation.ErrorMessage)
        ) {
            dataValidation.CreateErrorBox(
                validation.ErrorTitle ?? "",
                validation.ErrorMessage ?? ""
            );
        }

        if (!string.IsNullOrWhiteSpace(validation.PromptTitle)
            || !string.IsNullOrWhiteSpace(validation.PromptMessage)
        ) {
            dataValidation.CreatePromptBox(
                validation.PromptTitle ?? "",
                validation.PromptMessage ?? ""
            );
        }

        if (validation.ValidationType == DataValidationType.List
            && dataValidation is XSSFDataValidation xssfValidation
        ) {
            xssfValidation.SuppressDropDownArrow = !validation.IsDropdownShown;
        }

        sheet.AddValidationData(dataValidation);
    }

    private IDataValidationConstraint CreateValidationConstraint(IDataValidationHelper helper, DataValidation validation) {
        return validation.ValidationType switch {
            DataValidationType.List => CreateListConstraint(helper, validation),
            DataValidationType.Integer => CreateNumericConstraint(helper, validation, ValidationType.INTEGER),
            DataValidationType.Decimal => CreateNumericConstraint(helper, validation, ValidationType.DECIMAL),
            DataValidationType.Date => CreateDateConstraint(helper, validation),
            DataValidationType.Time => CreateTimeConstraint(helper, validation),
            DataValidationType.TextLength => CreateNumericConstraint(helper, validation, ValidationType.TEXT_LENGTH),
            DataValidationType.Custom => helper.CreateCustomConstraint(validation.Formula),
            _ => throw new ArgumentException($"Unsupported validation type: {validation.ValidationType}")
        };
    }

    private IDataValidationConstraint CreateListConstraint(IDataValidationHelper helper, DataValidation validation) {
        if (validation.ListItems is null || !validation.ListItems.Any()) {
            throw new ArgumentException("ListItems cannot be null or empty for List validation type.");
        }

        string[] items = validation.ListItems.ToArray();
        return helper.CreateExplicitListConstraint(items);
    }

    private IDataValidationConstraint CreateNumericConstraint(
        IDataValidationHelper helper, DataValidation validation, int validationType
    ) {
        (int operatorType, string formula1, string formula2) = PrepareConstraintParameters(validation);
        return helper.CreateNumericConstraint(validationType, operatorType, formula1, formula2);
    }

    private IDataValidationConstraint CreateDateConstraint(IDataValidationHelper helper, DataValidation validation) {
        (int operatorType, string formula1, string formula2) = PrepareConstraintParameters(validation);
        return helper.CreateDateConstraint(operatorType, formula1, formula2, null);
    }

    private IDataValidationConstraint CreateTimeConstraint(IDataValidationHelper helper, DataValidation validation) {
        (int operatorType, string formula1, string formula2) = PrepareConstraintParameters(validation);
        return helper.CreateTimeConstraint(operatorType, formula1, formula2);
    }

    private (int operatorType, string formula1, string formula2) PrepareConstraintParameters(DataValidation validation) {
        if (!validation.Operator.HasValue) {
            throw new ArgumentException($"Operator is required for {validation.ValidationType} validation type.");
        }

        int operatorType = ConvertOperator(validation.Operator.Value);
        string formula1 = !string.IsNullOrWhiteSpace(validation.Formula)
            ? EnsureFormulaPrefix(validation.Formula)
            : validation.Value1?.ToString() ?? "";
        string formula2 = validation.Value2?.ToString() ?? "";

        if (string.IsNullOrWhiteSpace(formula1)) {
            throw new ArgumentException($"Either Formula or Value1 is required for {validation.ValidationType} validation type.");
        }

        return (operatorType, formula1, formula2);
    }

    private static string EnsureFormulaPrefix(string formula) {
        return formula.StartsWith("=") ? formula : $"={formula}";
    }

    private static int ConvertOperator(DataValidationOperator @operator) {
        if (!validationOperatorMap.TryGetValue(@operator, out int result)) {
            throw new ArgumentException($"Unsupported operator: {@operator}");
        }

        return result;
    }
}
