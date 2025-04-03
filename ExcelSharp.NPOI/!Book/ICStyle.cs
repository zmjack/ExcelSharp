using NPOI.SS.UserModel;
using VerticalAlignment = NPOI.SS.UserModel.VerticalAlignment;

namespace ExcelSharp.NPOI;

public interface ICStyle
{
    #region Alignment
    HorizontalAlignment Alignment { get; set; }
    VerticalAlignment VerticalAlignment { get; set; }
    #endregion

    #region Border
    BorderStyle BorderLeft { get; set; }
    RgbaColor LeftBorderColor { get; set; }

    BorderStyle BorderRight { get; set; }
    RgbaColor RightBorderColor { get; set; }

    BorderStyle BorderTop { get; set; }
    RgbaColor TopBorderColor { get; set; }

    BorderStyle BorderBottom { get; set; }
    RgbaColor BottomBorderColor { get; set; }

    BorderStyle BorderDiagonalLineStyle { get; set; }
    RgbaColor BorderDiagonalColor { get; set; }
    BorderDiagonal BorderDiagonal { get; set; }
    #endregion

    #region Fill
    FillPattern FillPattern { get; set; }
    RgbaColor FillBackgroundColor { get; set; }
    RgbaColor FillForegroundColor { get; set; }
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
