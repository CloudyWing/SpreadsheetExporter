﻿using System;
using System.Collections.Generic;
using System.Drawing;

namespace CloudyWing.SpreadsheetExporter {
    public struct CellFont : IEquatable<CellFont> {
        public static CellFont Empty = new CellFont();

        public CellFont(
            string name, short size = 0, Color? color = null, FontStyles style = FontStyles.None
        ) {
            color ??= Color.Empty;
            Name = name;
            Size = size;
            Color = (Color)color;
            Style = style;
        }

        /// <summary>
        /// 字體名稱
        /// </summary>
        public string Name { get; private set; }

        /// <summary>
        /// 字體大小
        /// </summary>
        public short Size { get; private set; }

        /// <summary>
        /// 字體顏色
        /// </summary>
        public Color Color { get; private set; }

        /// <summary>
        /// 字體樣式，如粗體、斜體等
        /// </summary>
        public FontStyles Style { get; private set; }

        /// <summary>
        /// 建立一個不同Name的副本
        /// </summary>
        /// <param name="name">字體名稱</param>
        public CellFont CloneAndSetName(string name) {
            CellFont font = this;
            font.Name = name;
            return font;
        }

        /// <summary>
        /// 建立一個副本並設定副本Size
        /// </summary>
        /// <param name="size">字體大小</param>
        public CellFont CloneAndSetSize(short size) {
            CellFont font = this;
            font.Size = size;
            return font;
        }

        /// <summary>
        /// 建立一個副本並設定副本Color
        /// </summary>
        /// <param name="color">字體顏色</param>
        public CellFont CloneAndSetColor(Color color) {
            CellFont font = this;
            font.Color = color;
            return font;
        }

        /// <summary>
        /// 建立一個副本並設定副本字體樣式
        /// </summary>
        /// <param name="style">副本字體樣式</param>
        public CellFont CloneAndSetStyle(FontStyles style) {
            CellFont font = this;
            font.Style = style;
            return font;
        }

        public override bool Equals(object obj) => obj is CellFont font && Equals(font);

        public bool Equals(CellFont other)
            => Name == other.Name
                && Size == other.Size
                && EqualityComparer<Color>.Default.Equals(Color, other.Color)
                && Style == other.Style;

        public override int GetHashCode() {
            int hashCode = 253911835;
            hashCode = hashCode * -1521134295 + EqualityComparer<string>.Default.GetHashCode(Name);
            hashCode = hashCode * -1521134295 + Size.GetHashCode();
            hashCode = hashCode * -1521134295 + Color.GetHashCode();
            hashCode = hashCode * -1521134295 + Style.GetHashCode();
            return hashCode;
        }

        public static bool operator ==(CellFont left, CellFont right) => left.Equals(right);

        public static bool operator !=(CellFont left, CellFont right) => !(left == right);
    }
}
