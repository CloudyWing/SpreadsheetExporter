using System;
using System.Collections.Generic;
using System.Drawing;

namespace CloudyWing.SpreadsheetExporter {
    public struct CellStyle : IEquatable<CellStyle> {
        public CellStyle(
            HorizontalAlignment halign = HorizontalAlignment.Center, VerticalAlignment valign = VerticalAlignment.Middle,
            bool hasBorder = false, bool wrapText = false, Color? backgroundColor = null,
            CellFont? font = null, string dataFormat = null, bool isLocked = false
        ) : this() {
            backgroundColor ??= Color.Empty;
            font ??= new CellFont();
            HorizontalAlignment = halign;
            VerticalAlignment = valign;
            HasBorder = hasBorder;
            WrapText = wrapText;
            BackgroundColor = (Color)backgroundColor;
            Font = (CellFont)font;
            DataFormat = dataFormat;
            IsLocked = isLocked;
        }

        /// <summary>
        /// 水平對齊
        /// </summary>
        public HorizontalAlignment HorizontalAlignment { get; private set; }

        /// <summary>
        /// 垂直對齊
        /// </summary>
        public VerticalAlignment VerticalAlignment { get; private set; }

        /// <summary>
        /// 是否有格線
        /// </summary>
        public bool HasBorder { get; private set; }

        /// <summary>
        /// 是否自動換行
        /// </summary>
        public bool WrapText { get; private set; }

        /// <summary>
        /// 背景顏色
        /// </summary>
        public Color BackgroundColor { get; private set; }

        /// <summary>
        /// 字體格式
        /// </summary>
        public CellFont Font { get; private set; }

        /// <summary>
        /// 資料格式字串
        /// </summary>
        public string DataFormat { get; private set; }

        /// <summary>
        /// 是否鎖定儲存格
        /// </summary>
        public bool IsLocked { get; private set; }

        /// <summary>
        /// 建立一個副本，並設定副本的的水平對齊
        /// </summary>
        /// <param name="align">水平對齊</param>
        public CellStyle CloneAndSetHorizontalAlignment(HorizontalAlignment align) {
            CellStyle style = this;
            style.HorizontalAlignment = align;
            return style;
        }

        /// <summary>
        /// 建立一個副本，並設定副本的的垂直對齊
        /// </summary>
        /// <param name="valign">垂直對齊</param>
        public CellStyle CloneAndSetVerticalAlignment(VerticalAlignment valign) {
            CellStyle style = this;
            style.VerticalAlignment = valign;
            return style;
        }

        /// <summary>
        /// 建立一個副本，並設定副本是否顯示框線
        /// </summary>
        /// <param name="hasBolder">是否有框線</param>
        public CellStyle CloneAndSetBorder(bool hasBolder) {
            CellStyle style = this;
            style.HasBorder = hasBolder;
            return style;
        }

        /// <summary>
        /// 建立一個副本，並設定副本是否自動換行
        /// </summary>
        /// <param name="wrapText">是否自動換行</param>
        public CellStyle CloneAndSetWrapText(bool wrapText) {
            CellStyle style = this;
            style.WrapText = wrapText;
            return style;
        }

        /// <summary>
        /// 建立一個副本，並設定副本背景顏色
        /// </summary>
        /// <param name="backgroundColor">背景顏色</param>
        public CellStyle CloneAndSetBackgroundColor(Color backgroundColor) {
            CellStyle style = this;
            style.BackgroundColor = backgroundColor;
            return style;
        }

        /// <summary>
        /// 建立一個副本，並設定副本字體
        /// </summary>
        /// <param name="font">字體</param>
        public CellStyle CloneAndSetFont(CellFont font) {
            CellStyle style = this;
            style.Font = font;
            return style;
        }

        /// <summary>
        /// 建立一個副本，並設定副本的資料格式字串
        /// </summary>
        /// <param name="font">字體</param>
        public CellStyle CloneAndSetDataFormat(string dataForamt) {
            CellStyle style = this;
            style.DataFormat = dataForamt;
            return style;
        }

        /// <summary>
        /// 建立一個副本，並設定副本的是否鎖定
        /// </summary>
        /// <param name="font">字體</param>
        public CellStyle CloneAndSetLocked(bool isLocked) {
            CellStyle style = this;
            style.IsLocked = isLocked;
            return style;
        }

        public override bool Equals(object obj) => obj is CellStyle style && Equals(style);

        public bool Equals(CellStyle other)
            => HorizontalAlignment == other.HorizontalAlignment
                && VerticalAlignment == other.VerticalAlignment
                && HasBorder == other.HasBorder
                && WrapText == other.WrapText
                && EqualityComparer<Color>.Default.Equals(BackgroundColor, other.BackgroundColor)
                && EqualityComparer<CellFont>.Default.Equals(Font, other.Font)
                && DataFormat == other.DataFormat
                && IsLocked == other.IsLocked;

        public override int GetHashCode() {
            int hashCode = 253911835;
            hashCode = hashCode * -1521134295 + HorizontalAlignment.GetHashCode();
            hashCode = hashCode * -1521134295 + VerticalAlignment.GetHashCode();
            hashCode = hashCode * -1521134295 + HasBorder.GetHashCode();
            hashCode = hashCode * -1521134295 + WrapText.GetHashCode();
            hashCode = hashCode * -1521134295 + BackgroundColor.GetHashCode();
            hashCode = hashCode * -1521134295 + Font.GetHashCode();
            hashCode = hashCode * -1521134295 + EqualityComparer<string>.Default.GetHashCode(DataFormat);
            hashCode = hashCode * -1521134295 + IsLocked.GetHashCode();
            return hashCode;
        }

        public static bool operator ==(CellStyle left, CellStyle right) => left.Equals(right);

        public static bool operator !=(CellStyle left, CellStyle right) => !(left == right);
    }
}
