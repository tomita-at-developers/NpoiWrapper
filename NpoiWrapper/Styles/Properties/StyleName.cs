using Developers.NpoiWrapper.Styles.Models;
using Developers.NpoiWrapper.Utils;

namespace Developers.NpoiWrapper.Styles.Properties
{
    /// <summary>
    /// PoiCellStyle更新対象プロパティ名
    /// </summary>
    internal static class StyleName
    {
        public static readonly string HorizontaiAlignment = NameOf<PoiCellStyle>.FullName(n => n.Alignment);
        public static readonly string VerticalAlignment = NameOf<PoiCellStyle>.FullName(n => n.VerticalAlignment);
        public static readonly string Locked = NameOf<PoiCellStyle>.FullName(n => n.IsLocked);
        public static readonly string WrapText = NameOf<PoiCellStyle>.FullName(n => n.WrapText);
        public static readonly string NumberFormat = NameOf<PoiCellStyle>.FullName(n => n.DataFormatString);
        public static class Interior
        {
            public static readonly string ColorIndex = NameOf<PoiCellStyle>.FullName(n => n.FillBackgroundColor);
            public static readonly string Pattern = NameOf<PoiCellStyle>.FullName(n => n.FillPattern);
            public static readonly string PatternColorIndex = NameOf<PoiCellStyle>.FullName(n => n.FillForegroundColor);
        }
        /// <summary>
        /// 罫線関連はPoiCellStyleプロパティ名ではなく、Excel風の仮想プロパティ名を定義している
        /// </summary>
        public static class XlsBorder
        {
            public static readonly string LineStyle = NameOf<Border>.FullName(nameof(Border), n => n.LineStyle);
            public static readonly string Weight = NameOf<Border>.FullName(nameof(Border), n => n.Weight);
            public static readonly string ColorIndex = NameOf<Border>.FullName(nameof(Border), n => n.ColorIndex);
        }
        public static class PoiBorder
        {
            public static class Top
            {
                public static readonly string Style = NameOf<PoiCellStyle>.FullName(n => n.BorderTop);
                public static readonly string Color = NameOf<PoiCellStyle>.FullName(n => n.TopBorderColor);
            }
            public static class Bottom
            {
                public static readonly string Style = NameOf<PoiCellStyle>.FullName(n => n.BorderBottom);
                public static readonly string Color = NameOf<PoiCellStyle>.FullName(n => n.BottomBorderColor);
            }
            public static class Left
            {
                public static readonly string Style = NameOf<PoiCellStyle>.FullName(n => n.BorderLeft);
                public static readonly string Color = NameOf<PoiCellStyle>.FullName(n => n.LeftBorderColor);
            }
            public static class Right
            {
                public static readonly string Style = NameOf<PoiCellStyle>.FullName(n => n.BorderRight);
                public static readonly string Color = NameOf<PoiCellStyle>.FullName(n => n.RightBorderColor);
            }
            public static class Diagonal
            {
                public static readonly string Type = NameOf<PoiCellStyle>.FullName(n => n.BorderDiagonal);
                public static readonly string Style = NameOf<PoiCellStyle>.FullName(n => n.BorderDiagonalLineStyle);
                public static readonly string Color = NameOf<PoiCellStyle>.FullName(n => n.BorderDiagonalColor);
            }
        }
        public static class Font
        {
            public static readonly string Bold = NameOf<PoiCellStyle>.FullName(n => n.PoiFont.IsBold);
            public static readonly string ColorIndex = NameOf<PoiCellStyle>.FullName(n => n.PoiFont.Color);
            public static readonly string Italic = NameOf<PoiCellStyle>.FullName(n => n.PoiFont.IsItalic);
            public static readonly string Name = NameOf<PoiCellStyle>.FullName(n => n.PoiFont.FontName);
            public static readonly string Size = NameOf<PoiCellStyle>.FullName(n => n.PoiFont.FontHeightInPoints);
            public static readonly string Strikethrough = NameOf<PoiCellStyle>.FullName(n => n.PoiFont.IsStrikeout);
            public static readonly string TypeOffset = NameOf<PoiCellStyle>.FullName(n => n.PoiFont.TypeOffset);
            public static readonly string Underline = NameOf<PoiCellStyle>.FullName(n => n.PoiFont.Underline);
        }
    }
}
