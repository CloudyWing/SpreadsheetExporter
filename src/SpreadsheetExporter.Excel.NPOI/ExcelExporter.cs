using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Reflection;
using NPOI.HSSF.UserModel;
using NPOI.HSSF.Util;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;

namespace CloudyWing.SpreadsheetExporter.Excel.NPOI {
    /// <summary>The excel exporter, using npoi.</summary>
    /// <seealso cref="ExporterBase" />
    public sealed class ExcelExporter : ExporterBase {
        private IWorkbook workbook;
        private readonly IDictionary<CellStyle, ICellStyle> cellStyles = new Dictionary<CellStyle, ICellStyle>();
        private readonly IDictionary<CellFont, IFont> fonts = new Dictionary<CellFont, IFont>();
        private static readonly IDictionary<ExcelFormat, string> filenameExtensionMaps = new Dictionary<ExcelFormat, string> {
            [ExcelFormat.ExcelBinaryFileFormat] = ".xls",
            [ExcelFormat.OfficeOpenXmlDocument] = ".xlsx"
        };

        private static readonly IDictionary<ExcelFormat, string> contentTypeMaps = new Dictionary<ExcelFormat, string> {
            [ExcelFormat.ExcelBinaryFileFormat] = "application/vnd.ms-excel",
            [ExcelFormat.OfficeOpenXmlDocument] = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        };

        private readonly Dictionary<HorizontalAlignment, global::NPOI.SS.UserModel.HorizontalAlignment> horizontalAlignmentMaps = new() {
            [HorizontalAlignment.General] = global::NPOI.SS.UserModel.HorizontalAlignment.General,
            [HorizontalAlignment.Left] = global::NPOI.SS.UserModel.HorizontalAlignment.Left,
            [HorizontalAlignment.Center] = global::NPOI.SS.UserModel.HorizontalAlignment.Center,
            [HorizontalAlignment.Right] = global::NPOI.SS.UserModel.HorizontalAlignment.Right,
            [HorizontalAlignment.Justify] = global::NPOI.SS.UserModel.HorizontalAlignment.Justify
        };

        private readonly Dictionary<VerticalAlignment, global::NPOI.SS.UserModel.VerticalAlignment> verticalAlignmentMaps = new() {
            [VerticalAlignment.Top] = global::NPOI.SS.UserModel.VerticalAlignment.Top,
            [VerticalAlignment.Middle] = global::NPOI.SS.UserModel.VerticalAlignment.Center,
            [VerticalAlignment.Bottom] = global::NPOI.SS.UserModel.VerticalAlignment.Bottom
        };

        private readonly object thisLock = new();

        /// <summary>
        /// Initializes a new instance of the <see cref="ExcelExporter"/> class.
        /// </summary>
        /// <param name="excelFormat">The excel format.</param>
        public ExcelExporter(ExcelFormat excelFormat = ExcelFormat.OfficeOpenXmlDocument) {
            ExcelFormat = excelFormat;
        }

        /// <summary>Gets or sets the excel format.</summary>
        /// <value>The excel format.</value>
        public ExcelFormat ExcelFormat { get; set; }

        /// <summary>Gets a value indicating whether this instance is office open XML document.</summary>
        /// <value>
        ///   <c>true</c> if this instance is office open XML document; otherwise, <c>false</c>.</value>
        public bool IsOfficeOpenXmlDocument => ExcelFormat == ExcelFormat.OfficeOpenXmlDocument;

        /// <summary>Gets or sets a value indicating whether this instance is closed not implemented exception.</summary>
        /// <value>
        ///   <c>true</c> if this instance is closed not implemented exception; otherwise, <c>false</c>.</value>
        public bool IsClosedNotImplementedException { get; set; }

        /// <inheritdoc/>
        public override string ContentType => contentTypeMaps[ExcelFormat];

        /// <inheritdoc/>
        public override string FileNameExtension => filenameExtensionMaps[ExcelFormat];

        /// <inheritdoc/>
        protected override byte[] ExecuteExport(IEnumerable<SheeterContext> contexts) {
            lock (thisLock) {
                // 因為 ParseCellStyle 和 ParseFont 會用到，所以用 field 處理
                workbook = IsOfficeOpenXmlDocument ? new XSSFWorkbook() : new HSSFWorkbook();
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

                // Dispose() 和 Close() 疑似有問題，所以設為 null
                workbook = null;
                cellStyles.Clear();
                fonts.Clear();

                return ms.ToArray();
            }
        }

        private void CreateSheet(SheeterContext context) {
            ISheet sheet = workbook.CreateSheet(context.SheetName);
            // 不知道為什麼預設給很小，所以設定 default
            sheet.DefaultRowHeight = 330;
            SetSheetCells(sheet, context.Cells);
            SetSheetColumnWidths(sheet, context.ColumnWidths);
            SetSheetRowHeights(sheet, context.RowHeights);

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

            excelCellStyle.Alignment = horizontalAlignmentMaps[cellStyle.HorizontalAlignment];
            excelCellStyle.VerticalAlignment = verticalAlignmentMaps[cellStyle.VerticalAlignment];

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
                FieldInfo field = type.GetField("Triplet", BindingFlags.Public | BindingFlags.Static);
                // Automatic 這個顏色沒有 Triplet
                if (field != null) {
                    byte[] rgb = field.GetValue(null) as byte[];
                    if (rgb[0] == color.R && rgb[1] == color.G && rgb[2] == color.B) {
                        return (short)type.GetField("Index").GetValue(null);
                    }
                }
            }
            return 0;
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
            if (value == null) {
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
                    } catch {
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
                    sheet.SetColumnWidth(pair.Key, (short)(pair.Value * 256));
                }
            }
        }

        private void SetSheetRowHeights(ISheet sheet, IReadOnlyDictionary<int, double> rowHeights) {
            foreach (KeyValuePair<int, double> pair in rowHeights) {
                IRow row = sheet.GetRow(pair.Key) ?? sheet.CreateRow(pair.Key);
                if (pair.Value <= Constants.AutoFiteRowHeight) {
                    row.Height = -1;
                } else if (pair.Value == Constants.HiddenRow) {
                    row.ZeroHeight = true;
                } else {
                    row.Height = (short)(pair.Value * 20);
                }
            }
        }
    }
}
