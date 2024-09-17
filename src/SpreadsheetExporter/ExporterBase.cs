﻿using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using CloudyWing.SpreadsheetExporter.Exceptions;

namespace CloudyWing.SpreadsheetExporter {
    /// <summary>The exporter base.</summary>
    public abstract class ExporterBase {
        private readonly IList<Sheeter> sheeters = new List<Sheeter>();

        /// <summary>Occurs when [spreadsheet exporting event].</summary>
        public event EventHandler<SpreadsheetExportingEventArgs> SpreadsheetExportingEvent;

        /// <summary>Occurs when [spreadsheet exported event].</summary>
        public event EventHandler<SpreadsheetExportedEventArgs> SpreadsheetExportedEvent;

        /// <summary>Occurs when [sheet created event].</summary>
        public event EventHandler<SheetCreatedEventArgs> SheetCreatedEvent;

        /// <summary>Gets the content-type.</summary>
        /// <value>The content-type.</value>
        public abstract string ContentType { get; }

        /// <summary>Gets the file name extension.</summary>
        /// <value>The file name extension.</value>
        public abstract string FileNameExtension { get; }

        /// <summary>Gets or sets the workbook protection password.</summary>
        /// <value>The password.</value>
        public string Password { get; set; }

        /// <summary>Gets a value indicating whether this instance has password.</summary>
        /// <value>
        ///   <c>true</c> if this instance has password; otherwise, <c>false</c>.</value>
        public bool HasPassword => !string.IsNullOrEmpty(Password);

        /// <summary>
        /// Gets or sets the default font.</summary>
        /// <value>The default font.</value>
        public CellFont? DefaultFont { get; set; }

        /// <summary>Gets or sets the default basic name of the sheet.</summary>
        /// <value>The default basic name of the sheet.</value>
        public string DefaultBasicSheetName { get; set; } = "工作表";

        /// <summary>Gets the last sheeter. When there is no sheeter, create a sheeter.</summary>
        /// <value>The last sheeter.</value>
        public Sheeter LastSheeter => sheeters.LastOrDefault() ?? CreateSheeter();

        /// <summary>Creates the sheeter.</summary>
        /// <param name="sheetName">Name of the sheet.</param>
        /// <param name="defaultRowHeight">Default height of the row.</param>
        /// <returns>The sheeter.</returns>
        public Sheeter CreateSheeter(string sheetName = null, double? defaultRowHeight = null) {
            if (string.IsNullOrWhiteSpace(sheetName)) {
                sheetName = GetDefaultSheetName();
            } else if (IsSheetNameExists(sheetName)) {
                sheetName = FixSheetName(sheetName);
            }

            Sheeter sheeter = new(sheetName) {
                DefaultRowHeight = defaultRowHeight
            };
            sheeters.Add(sheeter);

            return sheeter;
        }

        /// <summary>Gets the sheeter.</summary>
        /// <param name="index">The index.</param>
        /// <returns>The sheeter.</returns>
        public Sheeter GetSheeter(int index) {
            return sheeters[index];
        }

        private bool IsSheetNameExists(string sheetName) {
            return sheeters.Select(x => x.SheetName).Contains(sheetName);
        }

        private string GetDefaultSheetName() {
            string basicSheetName = DefaultBasicSheetName;
            string defaultSheetName;
            int i = 1;
            do {
                defaultSheetName = basicSheetName + i++;
            } while (IsSheetNameExists(defaultSheetName));

            return defaultSheetName;
        }

        private string FixSheetName(string sheetName) {
            string fixedSheetName;
            int i = 1;
            do {
                fixedSheetName = $"{sheetName}({i++})";
            } while (IsSheetNameExists(fixedSheetName));

            return fixedSheetName;
        }

        /// <summary>Exports bytes of spreadsheet.</summary>
        /// <returns>The bytes of spreadsheet.</returns>
        /// <exception cref="SheeterNotFoundException">Sheeter[0] is null.</exception>
        public byte[] Export() {
            if (sheeters.Count == 0) {
                throw new SheeterNotFoundException();
            }

            IEnumerable<SheeterContext> contexts = sheeters.Select(x => new SheeterContext(x));

            OnSpreadsheetExporting(new SpreadsheetExportingEventArgs(contexts));
            byte[] bytes = ExecuteExport(contexts);
            OnSpreadsheetExported(new SpreadsheetExportedEventArgs(contexts));

            return bytes;
        }

        /// <summary>Raises the <see cref="E:SpreadsheetExporting" /> event.</summary>
        /// <param name="args">The <see cref="SpreadsheetExportingEventArgs" /> instance containing the event data.</param>
        protected virtual void OnSpreadsheetExporting(SpreadsheetExportingEventArgs args) {
            SpreadsheetExportingEvent?.Invoke(this, args);
        }

        /// <summary>Executes the export.</summary>
        /// <param name="contexts">The contexts.</param>
        /// <returns>The exported bytes of the file.</returns>
        protected abstract byte[] ExecuteExport(IEnumerable<SheeterContext> contexts);

        /// <summary>Raises the <see cref="E:SpreadsheetExported" /> event.</summary>
        /// <param name="args">The <see cref="SpreadsheetExportedEventArgs" /> instance containing the event data.</param>
        protected virtual void OnSpreadsheetExported(SpreadsheetExportedEventArgs args) {
            SpreadsheetExportedEvent?.Invoke(this, args);
        }

        /// <summary>Raises the <see cref="E:SheetCreated" /> event.</summary>
        /// <param name="args">The <see cref="SheetCreatedEventArgs" /> instance containing the event data.</param>
        protected virtual void OnSheetCreated(SheetCreatedEventArgs args) {
            SheetCreatedEvent?.Invoke(this, args);
        }

        /// <summary>Exports the spreadsheet file.</summary>
        /// <param name="path">The path.</param>
        /// <param name="fileMode">The file mode.</param>
        public void ExportFile(string path, SpreadsheetFileMode fileMode = SpreadsheetFileMode.Create) {
            byte[] bytes = Export();
            FileMode mode = fileMode == SpreadsheetFileMode.Create
                ? FileMode.Create
                : FileMode.CreateNew;

            using FileStream fileStream = new(path, mode, FileAccess.Write);
            fileStream.Write(bytes, 0, bytes.Length);
        }
    }
}
