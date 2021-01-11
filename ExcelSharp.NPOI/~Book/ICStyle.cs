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
        RgbColor LeftBorderColor { get; set; }

        BorderStyle BorderRight { get; set; }
        RgbColor RightBorderColor { get; set; }

        BorderStyle BorderTop { get; set; }
        RgbColor TopBorderColor { get; set; }

        BorderStyle BorderBottom { get; set; }
        RgbColor BottomBorderColor { get; set; }

        BorderStyle BorderDiagonalLineStyle { get; set; }
        RgbColor BorderDiagonalColor { get; set; }
        BorderDiagonal BorderDiagonal { get; set; }
        #endregion

        #region Fill
        FillPattern FillPattern { get; set; }
        RgbColor FillBackgroundColor { get; set; }
        RgbColor FillForegroundColor { get; set; }
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
