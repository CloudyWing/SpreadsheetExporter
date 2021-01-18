namespace CloudyWing.SpreadsheetExporter.Config {
    internal class CellSettings {
        public HorizontalAlignment HorizontalAlignment { get; set; }

        public VerticalAlignment VerticalAlignment { get; set; }

        public bool HasBorder { get; set; }

        public bool WrapText { get; set; }

        public FontSettings Font { get; set; }
    }
}
