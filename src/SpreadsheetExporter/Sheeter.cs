using System.Collections.Generic;
using System.Collections.ObjectModel;
using CloudyWing.SpreadsheetExporter.Templates;

namespace CloudyWing.SpreadsheetExporter {
    /// <summary>
    /// The sheeter.
    /// </summary>
    public class Sheeter {
        private readonly Dictionary<int, double> columnWidths = [];
        private readonly List<ITemplate> templates = [];

        internal Sheeter(string sheetName) {
            SheetName = sheetName;
        }

        /// <summary>
        /// Gets or sets the name of the sheet.
        /// </summary>
        /// <value>
        /// The name of the sheet.
        /// </value>
        public string SheetName { get; set; }

        /// <summary>
        /// Gets or sets the default height of the row.
        /// </summary>
        /// <value>
        /// The default height of the row.
        /// </value>
        public double? DefaultRowHeight { get; set; }

        /// <summary>
        /// Gets or sets the password.
        /// </summary>
        /// <value>
        /// The password.
        /// </value>
        public string Password { get; set; }

        /// <summary>
        /// Gets the page settings.
        /// </summary>
        /// <value>
        /// The page settings.
        /// </value>
        public PageSettings PageSettings { get; } = new PageSettings();

        /// <summary>
        /// Gets the width of the columns.
        /// </summary>
        /// <value>
        /// The width of columns.
        /// </value>
        public IReadOnlyDictionary<int, double> ColumnWidths => new ReadOnlyDictionary<int, double>(columnWidths);

        /// <summary>
        /// Gets the templates.
        /// </summary>
        /// <value>
        /// The templates.
        /// </value>
        public IReadOnlyList<ITemplate> Templates => new ReadOnlyCollection<ITemplate>(templates);

        /// <summary>
        /// Sets the width of the column.
        /// </summary>
        /// <param name="index">The index.</param>
        /// <param name="width">The width. If width is <c>0</c>, hide width. if the width is <c>-1</c>, the column width will be adjusted automatically.</param>
        public void SetColumnWidth(int index, double width) {
            columnWidths[index] = width;
        }

        /// <summary>
        /// Adds the template.
        /// </summary>
        /// <param name="template">The template.</param>
        public void AddTemplate(ITemplate template) {
            templates.Add(template);
        }

        /// <summary>
        /// Adds the templates.
        /// </summary>
        /// <param name="templates">The templates.</param>
        public void AddTemplates(IEnumerable<ITemplate> templates) {
            foreach (ITemplate template in templates) {
                AddTemplate(template);
            }
        }

        /// <summary>
        /// Adds the templates.
        /// </summary>
        /// <param name="templates">The templates.</param>
        public void AddTemplates(params ITemplate[] templates) {
            AddTemplates(templates as IEnumerable<ITemplate>);
        }
    }
}
