using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Linq;
using CloudyWing.SpreadsheetExporter.Extensions;
using CloudyWing.SpreadsheetExporter.Templates;

namespace CloudyWing.SpreadsheetExporter {
    /// <summary>The sheeter context.</summary>
    public class SheeterContext {
        private readonly Image watermark;

        /// <summary>Initializes a new instance of the <see cref="SheeterContext" /> class.</summary>
        /// <param name="sheeter">The sheeter.</param>
        public SheeterContext(Sheeter sheeter) {
            SheetName = sheeter.SheetName;
            TemplateContext templateContext = TemplateContext.Create(sheeter.Templates);
            Cells = templateContext.Cells;
            RowHeights = templateContext.RowHeights.AsReadOnly();
            ColumnWidths = sheeter.ColumnWidths.ToDictionary(x => x.Key, x => x.Value).AsReadOnly();
            Password = sheeter.Password;
            PageSettings = sheeter.PageSettings;
            watermark = sheeter.Watermark;
        }

        /// <summary>Gets the name of the sheet.</summary>
        /// <value>The name of the sheet.</value>
        public string SheetName { get; }

        /// <summary>Gets the cells.</summary>
        /// <value>The cells.</value>
        public IReadOnlyList<Cell> Cells { get; }

        /// <summary>Gets the width of columns.</summary>
        /// <value>The width of columns.</value>
        public IReadOnlyDictionary<int, double> ColumnWidths { get; }

        /// <summary>Gets the height of rows.</summary>
        /// <value>The height of rows.</value>
        public IReadOnlyDictionary<int, double> RowHeights { get; private set; }

        /// <summary>Gets the password.</summary>
        /// <value>The password.</value>
        public string Password { get; }

        /// <summary>Gets a value indicating whether this instance is protected.</summary>
        /// <value>
        ///   <c>true</c> if this instance is protected; otherwise, <c>false</c>.</value>
        public bool IsProtected => !string.IsNullOrEmpty(Password);

        /// <summary>Gets the page settings.</summary>
        /// <value>The page settings.</value>
        public PageSettings PageSettings { get; } = new PageSettings();

        /// <summary>Gets a value indicating whether this instance has watermark.</summary>
        /// <value>
        ///   <c>true</c> if this instance has watermark; otherwise, <c>false</c>.</value>
        public bool HasWatermark => watermark is not null;

        /// <summary>Gets or sets the watermark.</summary>
        /// <value>The watermark.</value>
        public Image Watermark {
            get {
                if (watermark is null) {
                    return null;
                }

                bool isPortrait = PageSettings.PageOrientation == PageOrientation.Portrait;

                int width = isPortrait
                    ? PageSettings.PaperSize.Width
                    : PageSettings.PaperSize.Height;
                int height = isPortrait
                    ? PageSettings.PaperSize.Height
                    : PageSettings.PaperSize.Width;

                if (watermark.Width > width || watermark.Height > height) {
                    using Image image = ZoomOutImage(width, height);
                    return ResizeImageBackgroundToFullPage(width, height, image);
                }

                return ResizeImageBackgroundToFullPage(width, height, watermark);
            }
        }

        private Image ZoomOutImage(int pageWidth, int pageHeight) {
            decimal scale = Math.Max((decimal)watermark.Width / pageWidth, (decimal)watermark.Height / pageHeight);
            return new Bitmap(watermark, (int)(watermark.Width / scale), (int)(watermark.Height / scale));
        }

        private Image ResizeImageBackgroundToFullPage(int pageWidth, int pageHeight, Image image) {
            Image bitmap = new Bitmap(pageWidth, pageHeight);
            using Graphics graphics = Graphics.FromImage(bitmap);
            graphics.Clear(Color.White);
            graphics.DrawImage(image, (pageWidth - image.Width) / 2, (pageHeight - image.Height) / 2);
            graphics.CompositingQuality = CompositingQuality.HighQuality;
            graphics.SmoothingMode = SmoothingMode.HighQuality;
            graphics.Save();

            return bitmap;
        }
    }
}
