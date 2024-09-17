using System;

namespace CloudyWing.SpreadsheetExporter.Config {
    /// <summary>
    /// The cell style setuper.
    /// </summary>
    /// <param name="configuration">
    /// The configuration.
    /// </param>
    /// <exception cref="ArgumentNullException">
    /// configuration
    /// </exception>
    public sealed class CellStyleSetuper(CellStyleConfiguration configuration) {
        private readonly CellStyleConfiguration configuration = configuration
            ?? throw new ArgumentNullException(nameof(configuration));

        /// <summary>
        /// Gets or sets the cell style.
        /// </summary>
        /// <value>
        /// The cell style.
        /// </value>
        public CellStyle CellStyle {
            get => configuration.CellStyle;
            set => configuration.CellStyle = value;
        }

        /// <summary>
        /// Gets or sets the grid cell style.
        /// </summary>
        /// <value>
        /// The grid cell style.
        /// </value>
        public CellStyle GridCellStyle {
            get => configuration.GridCellStyle;
            set => configuration.GridCellStyle = value;
        }

        /// <summary>
        /// Gets or sets the header style.
        /// </summary>
        /// <value>
        /// The header style.
        /// </value>
        public CellStyle HeaderStyle {
            get => configuration.HeaderStyle;
            set => configuration.HeaderStyle = value;
        }

        /// <summary>
        /// Gets or sets the field style.
        /// </summary>
        /// <value>
        /// The field style.
        /// </value>
        public CellStyle FieldStyle {
            get => configuration.FieldStyle;
            set => configuration.FieldStyle = value;
        }
    }
}
