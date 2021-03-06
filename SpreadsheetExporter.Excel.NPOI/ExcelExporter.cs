﻿using System;
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
    public sealed class ExcelExporter : ExporterBase {
        private IWorkbook workbook;
        private readonly IDictionary<CellStyle, ICellStyle> cellStyles = new Dictionary<CellStyle, ICellStyle>();
        private readonly IDictionary<CellFont, IFont> fonts = new Dictionary<CellFont, IFont>();
        private static readonly IDictionary<ExcelFormat, string> filenameExtensionMap = new Dictionary<ExcelFormat, string> {
            [ExcelFormat.ExcelBinaryFileFormat] = ".xls",
            [ExcelFormat.OfficeOpenXmlDocument] = ".xlsx"
        };

        private readonly Dictionary<HorizontalAlignment, global::NPOI.SS.UserModel.HorizontalAlignment> horizontalAlignmentMap = new Dictionary<HorizontalAlignment, global::NPOI.SS.UserModel.HorizontalAlignment>() {
            [HorizontalAlignment.None] = global::NPOI.SS.UserModel.HorizontalAlignment.General,
            [HorizontalAlignment.Left] = global::NPOI.SS.UserModel.HorizontalAlignment.Left,
            [HorizontalAlignment.Center] = global::NPOI.SS.UserModel.HorizontalAlignment.Center,
            [HorizontalAlignment.Right] = global::NPOI.SS.UserModel.HorizontalAlignment.Right,
            [HorizontalAlignment.Justify] = global::NPOI.SS.UserModel.HorizontalAlignment.Justify
        };

        private readonly Dictionary<VerticalAlignment, global::NPOI.SS.UserModel.VerticalAlignment> verticalAlignmentMap = new Dictionary<VerticalAlignment, global::NPOI.SS.UserModel.VerticalAlignment>() {
            [VerticalAlignment.Top] = global::NPOI.SS.UserModel.VerticalAlignment.Top,
            [VerticalAlignment.Middle] = global::NPOI.SS.UserModel.VerticalAlignment.Center,
            [VerticalAlignment.Bottom] = global::NPOI.SS.UserModel.VerticalAlignment.Bottom
        };

        private readonly object thisLock = new object();

        public ExcelExporter(ExcelFormat excelFormat = ExcelFormat.OfficeOpenXmlDocument) {
            ExcelFormat = excelFormat;
        }

        public ExcelFormat ExcelFormat { get; set; }

        public bool IsOfficeOpenXmlDocument => ExcelFormat == ExcelFormat.OfficeOpenXmlDocument;

        public override string ContentType => "application/ms-excel";

        public override string FileNameExtension => filenameExtensionMap[ExcelFormat];

        protected override byte[] ExecuteExport(SheeterContext[] contexts) {
            lock (thisLock) {
                // 因為ParseCellStyle和ParseFont會用到，所以用field處理
                workbook = IsOfficeOpenXmlDocument ? new XSSFWorkbook() : (IWorkbook)new HSSFWorkbook();
                foreach (SheeterContext context in contexts) {
                    CreateSheet(context);
                }

                using MemoryStream ms = new MemoryStream();
                workbook.Write(ms);
                // Dispose()和Close()疑似有問題，所以設為null
                workbook = null;
                cellStyles.Clear();
                fonts.Clear();

                return ms.ToArray();
            }
        }

        private void CreateSheet(SheeterContext context) {
            ISheet sheet = workbook.CreateSheet(context.SheetName);
            // 不知道為什麼預設給很小，所以設定default
            sheet.DefaultRowHeight = 330;
            SetSheetCell(sheet, context.Cells);
            SetSheetColumnWidths(sheet, context.ColumnWidths);
            SetSheetRowHeights(sheet, context.RowHeights);

            if (context.IsProtected) {
                sheet.ProtectSheet(context.Password);
            }
        }

        private void SetSheetCell(ISheet sheet, IReadOnlyList<Cell> cells) {
            foreach (Cell cell in cells) {
                IRow excelRow = sheet.GetRow(cell.Point.Y) ?? sheet.CreateRow(cell.Point.Y);
                ICell excelCell = excelRow.GetCell(cell.Point.X) ?? excelRow.CreateCell(cell.Point.X);
                SetValueToCell(excelCell, cell.Value);
                excelCell.CellStyle = ParseCellStyle(cell.CellStyle);
                excelCell.CellFormula = cell.Formula;

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

            // 聽說CellStyle建立太多，可能會出現問題，所以如果有已存在的，就直接使用
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

            cellStyles.Add(cellStyle, excelCellStyle);

            return excelCellStyle;
        }

        private IFont ParseFont(CellFont font) {
            // 聽說CellStyle建立太多，可能會出現問題，所以如果有已存在的，就直接使用
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

        /// <summary>
        /// HSSFColor無法用Color設定顏色，所以如果xls版本就使用Reflection找出對應顏色Index
        /// </summary>
        private short ParseColor(Color color) {
            //找出此Type下公用的class
            Type[] colorTypes = typeof(HSSFColor).GetNestedTypes(BindingFlags.Public);

            foreach (Type type in colorTypes) {
                FieldInfo field = type.GetField("Triplet", BindingFlags.Public | BindingFlags.Static);
                // Automatic這個顏色沒有Triplet
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

        /// <summary>
        /// 合併儲存格，並把第一個儲存格的樣式複製到其他儲存格
        /// </summary>
        private void MergedRegion(ISheet sheet, int firstColnum, int lastColumn, int firstRow, int lastRow) {
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
                if (!(value is string) && value is IConvertible) {
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
                if (pair.Value <= Constant.AutoSizeColumn) {
                    sheet.AutoSizeColumn(pair.Key);
                } else if (pair.Value == Constant.HiddenColumn) {
                    sheet.SetColumnHidden(pair.Key, true);
                } else {
                    sheet.SetColumnWidth(pair.Key, (short)(pair.Value * 256));
                }
            }
        }

        private void SetSheetRowHeights(ISheet sheet, IReadOnlyDictionary<int, double> rowHeights) {
            foreach (KeyValuePair<int, double> pair in rowHeights) {
                IRow row = sheet.GetRow(pair.Key) ?? sheet.CreateRow(pair.Key);
                if (pair.Value <= Constant.AutoSizeRow) {
                    row.Height = -1;
                } else if (pair.Value == Constant.HiddenRow) {
                    row.ZeroHeight = true;
                } else {
                    row.Height = (short)(pair.Value * 20);
                }
            }
        }
    }
}
