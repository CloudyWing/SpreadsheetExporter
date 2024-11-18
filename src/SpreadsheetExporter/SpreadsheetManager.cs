using System;
using CloudyWing.SpreadsheetExporter.Config;

namespace CloudyWing.SpreadsheetExporter {
    /// <summary>
    /// The spreadsheet manager.
    /// </summary>
    public static class SpreadsheetManager {
        private static readonly object exporterFactoryLock = new();
        private static readonly object cellStyleLock = new();
        private static Func<ISpreadsheetExporter> exporterFactory;
        private static CellStyleConfiguration defaultCellStyles;

        /// <summary>
        /// Gets or sets the configuration.
        /// </summary>
        /// <value>
        /// The configuration.
        /// </value>
        public static CellStyleConfiguration DefaultCellStyles {
            get {
                if (defaultCellStyles is null) {
                    lock (cellStyleLock) {
                        defaultCellStyles ??= new CellStyleConfiguration();
                    }
                }
                return defaultCellStyles;
            }
            set {
                lock (cellStyleLock) {
                    defaultCellStyles = value;
                }
            }
        }

        /// <summary>
        /// Sets the exporter.
        /// </summary>
        /// <param name="exporterFactory">The exporter factory.</param>
        /// <exception cref="ArgumentNullException">exporterFactory</exception>
        /// <exception cref="ArgumentException">Factory return value cannot be null. - exporterFactory</exception>
        public static void SetExporter(Func<ISpreadsheetExporter> exporterFactory) {
            if (exporterFactory is null) {
                throw new ArgumentNullException(nameof(exporterFactory));
            }

            if (exporterFactory() is null) {
                throw new ArgumentException("Factory return value cannot be null.", nameof(exporterFactory));
            }

            lock (exporterFactoryLock) {
                SpreadsheetManager.exporterFactory = exporterFactory;
            }
        }

        /// <summary>
        /// Creates the exporter.
        /// </summary>
        /// <returns>The exporter.</returns>
        /// <exception cref="NullReferenceException">Exporter factory is not set.</exception>
        public static ISpreadsheetExporter CreateExporter() {
            return exporterFactory is null
                ? throw new InvalidOperationException("Exporter factory is not set.")
                : exporterFactory();
        }
    }
}
