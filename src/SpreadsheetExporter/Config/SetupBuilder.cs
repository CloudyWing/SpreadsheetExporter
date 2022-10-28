using System;

namespace CloudyWing.SpreadsheetExporter.Config {
    public sealed class SetupBuilder {
        private readonly SpreadsheetConfiguration configuration;

        public SetupBuilder(SpreadsheetConfiguration configuration) => this.configuration = configuration ?? throw new ArgumentNullException(nameof(configuration));

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
        /// Gets or sets the list header style.
        /// </summary>
        /// <value>
        /// The list header style.
        /// </value>
        public CellStyle ListHeaderStyle {
            get => configuration.ListHeaderStyle;
            set => configuration.ListHeaderStyle = value;
        }

        /// <summary>
        /// Gets or sets the list text style.
        /// </summary>
        /// <value>
        /// The list text style.
        /// </value>
        public CellStyle ListTextStyle {
            get => configuration.ListTextStyle;
            set => configuration.ListTextStyle = value;
        }

        /// <summary>
        /// Gets or sets the list number style.
        /// </summary>
        /// <value>
        /// The list number style.
        /// </value>
        public CellStyle ListNumberStyle {
            get => configuration.ListNumberStyle;
            set => configuration.ListNumberStyle = value;
        }

        /// <summary>
        /// Gets or sets the list date time style.
        /// </summary>
        /// <value>
        /// The list date time style.
        /// </value>
        public CellStyle ListDateTimeStyle {
            get => configuration.ListDateTimeStyle;
            set => configuration.ListDateTimeStyle = value;
        }
    }
}
