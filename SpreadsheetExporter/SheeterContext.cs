using System.Collections.Generic;
using System.Diagnostics;
using CloudyWing.SpreadsheetExporter.Extensions;

namespace CloudyWing.SpreadsheetExporter {
    public class SheeterContext {
        public SheeterContext(
            string sheetName,
            IEnumerable<TemplateContext> templateContexts,
            IReadOnlyDictionary<int, double> columnWidths
        ) {
            SheetName = sheetName;
            InitializeCellsAndRowHeights();
            ColumnWidths = columnWidths.AsReadOnly();

            void InitializeCellsAndRowHeights() {
                List<Cell> cells = new List<Cell>();
                Dictionary<int, double> rowHeights = new Dictionary<int, double>();
                int rowIndex = 0;

                foreach (TemplateContext context in templateContexts) {
                    foreach (KeyValuePair<int, double> pair in context.RowHeights) {
                        rowHeights.Add(rowIndex + pair.Key, pair.Value);
                    }

                    foreach (Cell cell in context.Cells) {
                        Cell fixedCell = new Cell() {
                            Value = cell.Value,
                            Point = new System.Drawing.Point() {
                                X = cell.Point.X,
                                Y = cell.Point.Y + rowIndex
                            },
                            Size = cell.Size,
                            CellStyle = cell.CellStyle
                        };

                        Debug.Assert(fixedCell.Point.Y == cell.Point.Y + rowIndex);

                        cells.Add(fixedCell);
                    }

                    rowIndex += context.RowSpan;
                }

                Cells = cells.AsReadOnly();
                RowHeights = rowHeights.AsReadOnly();
            }
        }

        public string SheetName { get; }

        public IReadOnlyList<Cell> Cells { get; private set; }

        public IReadOnlyDictionary<int, double> ColumnWidths { get; }

        public IReadOnlyDictionary<int, double> RowHeights { get; private set; }

        public string Password { get; set; }

        public bool IsProtected => !string.IsNullOrEmpty(Password);
    }
}
