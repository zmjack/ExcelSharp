using NPOI.SS.UserModel;
using NStandard;
using Richx;
using System;
using System.Linq;

namespace ExcelSharp.NPOI
{
    public class CStyleApplier : ICStyle
    {
        private CStyleApplier() { }

        internal static CStyleApplier Create(Action<CStyleApplier> init) => new CStyleApplier().Then(x => init?.Invoke(x));

        #region Alignment
        public HorizontalAlignment Alignment { get; set; } = HorizontalAlignment.General;
        public VerticalAlignment VerticalAlignment { get; set; } = VerticalAlignment.Bottom;
        #endregion

        #region Border
        public BorderStyle BorderLeft { get; set; } = BorderStyle.None;
        public IArgbColor LeftBorderColor { get; set; } = ExcelColor.Black;

        public BorderStyle BorderRight { get; set; } = BorderStyle.None;
        public IArgbColor RightBorderColor { get; set; } = ExcelColor.Black;

        public BorderStyle BorderTop { get; set; } = BorderStyle.None;
        public IArgbColor TopBorderColor { get; set; } = ExcelColor.Black;

        public BorderStyle BorderBottom { get; set; } = BorderStyle.None;
        public IArgbColor BottomBorderColor { get; set; } = ExcelColor.Black;

        public BorderStyle BorderDiagonalLineStyle { get; set; } = BorderStyle.None;
        public IArgbColor BorderDiagonalColor { get; set; } = ExcelColor.Black;
        public BorderDiagonal BorderDiagonal { get; set; } = BorderDiagonal.None;
        #endregion

        #region Fill
        public FillPattern FillPattern { get; set; } = FillPattern.NoFill;
        public IArgbColor FillBackgroundColor { get; set; } = ExcelColor.Automatic;
        public IArgbColor FillForegroundColor { get; set; } = ExcelColor.Automatic;
        #endregion

        #region Font
        public CFontApplier Font { get; set; } = CFontApplier.Create(null);
        #endregion

        #region DataFormat
        public string DataFormat { get; set; } = "General";
        #endregion

        #region Others
        public short Rotation { get; set; } = 0;
        public short Indention { get; set; } = 0;
        public bool WrapText { get; set; } = false;
        public bool IsLocked { get; set; } = true;
        public bool IsHidden { get; set; } = false;
        public bool ShrinkToFit { get; set; } = false;
        #endregion

        public void Apply(CStyle style)
        {
            //TODO: Use TypeReflectionCacheContainer to optimize it in the futrue.
            var props = typeof(ICStyle).GetProperties().Where(prop => prop.CanWrite);
            foreach (var prop in props)
                prop.SetValue(style, prop.GetValue(this));

            // Font
            style.Font = style.Book.CFont(Font);
        }

        public CStyleApplier FullBorder(BorderStyle borderStyle = BorderStyle.Thin)
        {
            BorderLeft = BorderRight = BorderTop = BorderBottom = borderStyle;
            return this;
        }
        public CStyleApplier LeftBorder(BorderStyle borderStyle = BorderStyle.Thin) { BorderLeft = borderStyle; return this; }
        public CStyleApplier RightBorder(BorderStyle borderStyle = BorderStyle.Thin) { BorderRight = borderStyle; return this; }
        public CStyleApplier TopBorder(BorderStyle borderStyle = BorderStyle.Thin) { BorderTop = borderStyle; return this; }
        public CStyleApplier BottomBorder(BorderStyle borderStyle = BorderStyle.Thin) { BorderBottom = borderStyle; return this; }

        public CStyleApplier CellColor(IArgbColor foregroundColor) => CellColor(foregroundColor, ExcelColor.Automatic, FillPattern.SolidForeground);
        public CStyleApplier CellColor(IArgbColor foregroundColor, IArgbColor backgroundColor, FillPattern pattern)
        {
            FillForegroundColor = foregroundColor;
            FillBackgroundColor = backgroundColor;

            if (ExcelColor.GetIndex(foregroundColor) == ExcelColor.AutomaticIndex
                && ExcelColor.GetIndex(backgroundColor) == ExcelColor.AutomaticIndex)
                FillPattern = FillPattern.NoFill;
            else FillPattern = pattern;

            return this;
        }

        public CStyleApplier Center() => HCenter().VCenter();

        public CStyleApplier HLeft() { Alignment = HorizontalAlignment.Left; return this; }
        public CStyleApplier HCenter() { Alignment = HorizontalAlignment.Center; return this; }
        public CStyleApplier HRight() { Alignment = HorizontalAlignment.Right; return this; }

        public CStyleApplier VTop() { VerticalAlignment = VerticalAlignment.Top; return this; }
        public CStyleApplier VCenter() { VerticalAlignment = VerticalAlignment.Center; return this; }
        public CStyleApplier VBottom() { VerticalAlignment = VerticalAlignment.Bottom; return this; }

        public CStyleApplier Bold(bool value = true) { Font.IsBold = value; return this; }
        public CStyleApplier Italic(bool value = true) { Font.IsItalic = value; return this; }
        public CStyleApplier Strikeout(bool value = true) { Font.IsStrikeout = value; return this; }
        public CStyleApplier Underline(FontUnderlineType value = FontUnderlineType.Single) { Font.Underline = value; return this; }
        public CStyleApplier TypeOffset(FontSuperScript value) { Font.TypeOffset = value; return this; }

        public CStyleApplier SetFont(string fontName, short size) => SetFont(fontName, size, ExcelColor.Automatic);
        public CStyleApplier SetFont(string fontName, short size, IArgbColor color)
        {
            Font.FontName = fontName;
            Font.FontSize = size;
            Font.FontColor = color;
            return this;
        }
        public CStyleApplier CellFormat(string dataFormat) { DataFormat = dataFormat; return this; }

        public CStyleApplier WordWrap(bool value = true) { WrapText = value; return this; }
        public CStyleApplier Shrink(bool value = true) { ShrinkToFit = value; return this; }

        public CStyleApplier SetRotation(short value) { Rotation = value; return this; }
        public CStyleApplier SetIndention(short value) { Indention = value; return this; }
    }
}
