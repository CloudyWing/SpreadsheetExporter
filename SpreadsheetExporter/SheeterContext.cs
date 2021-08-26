using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using CloudyWing.SpreadsheetExporter.Extensions;

namespace CloudyWing.SpreadsheetExporter {
    public class SheeterContext {
        public SheeterContext(Sheeter sheeter) {
            SheetName = sheeter.SheetName;
            InitializeCellsAndRowHeights();
            ColumnWidths = sheeter.ColumnWidths.ToDictionary(x => x.Key, x => x.Value).AsReadOnly();
            Password = sheeter.Password;

            void InitializeCellsAndRowHeights() {
                List<Cell> cells = new List<Cell>();
                Dictionary<int, double> rowHeights = new Dictionary<int, double>();
                int rowIndex = 0;

                foreach (TemplateContext context in sheeter.Templates.Select(x => x.GetContext())) {
                    foreach (KeyValuePair<int, double> pair in context.RowHeights) {
                        rowHeights.Add(rowIndex + pair.Key, pair.Value);
                    }

                    foreach (Cell cell in context.Cells) {
                        Cell fixedCell = cell.ShallowCopy();
                        fixedCell.Point = new System.Drawing.Point() {
                            X = cell.Point.X,
                            Y = cell.Point.Y + rowIndex
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

        public string Password { get; }

        public bool IsProtected => !string.IsNullOrEmpty(Password);
    }
}
