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
    IRgbaColor LeftBorderColor { get; set; }

    BorderStyle BorderRight { get; set; }
    IRgbaColor RightBorderColor { get; set; }

    BorderStyle BorderTop { get; set; }
    IRgbaColor TopBorderColor { get; set; }

    BorderStyle BorderBottom { get; set; }
    IRgbaColor BottomBorderColor { get; set; }

    BorderStyle BorderDiagonalLineStyle { get; set; }
    IRgbaColor BorderDiagonalColor { get; set; }
    BorderDiagonal BorderDiagonal { get; set; }
    #endregion

    #region Fill
    FillPattern FillPattern { get; set; }
    IRgbaColor FillBackgroundColor { get; set; }
    IRgbaColor FillForegroundColor { get; set; }
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
