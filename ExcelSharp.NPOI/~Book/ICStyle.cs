using NPOI.SS.UserModel;
using Richx;

namespace ExcelSharp.NPOI
{
    public interface ICStyle
    {
        #region Alignment
        HorizontalAlignment Alignment { get; set; }
        VerticalAlignment VerticalAlignment { get; set; }
        #endregion

        #region Border
        BorderStyle BorderLeft { get; set; }
        IArgbColor LeftBorderColor { get; set; }

        BorderStyle BorderRight { get; set; }
        IArgbColor RightBorderColor { get; set; }

        BorderStyle BorderTop { get; set; }
        IArgbColor TopBorderColor { get; set; }

        BorderStyle BorderBottom { get; set; }
        IArgbColor BottomBorderColor { get; set; }

        BorderStyle BorderDiagonalLineStyle { get; set; }
        IArgbColor BorderDiagonalColor { get; set; }
        BorderDiagonal BorderDiagonal { get; set; }
        #endregion

        #region Fill
        FillPattern FillPattern { get; set; }
        IArgbColor FillBackgroundColor { get; set; }
        IArgbColor FillForegroundColor { get; set; }
        #endregion

        #region Font
        #endregion

        #region DataFormat
        string DataFormat { get; set; }
        #endregion

        #region Others
        short Rotation { get; set; }
        short Indention { get; set; }
        bool WrapText { get; set; }
        bool IsLocked { get; set; }
        bool IsHidden { get; set; }
        bool ShrinkToFit { get; set; }
        #endregion

    }
}
