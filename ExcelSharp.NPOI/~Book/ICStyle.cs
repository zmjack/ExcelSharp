﻿using NPOI.SS.UserModel;

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
        RGBColor LeftBorderColor { get; set; }

        BorderStyle BorderRight { get; set; }
        RGBColor RightBorderColor { get; set; }

        BorderStyle BorderTop { get; set; }
        RGBColor TopBorderColor { get; set; }

        BorderStyle BorderBottom { get; set; }
        RGBColor BottomBorderColor { get; set; }

        BorderStyle BorderDiagonalLineStyle { get; set; }
        RGBColor BorderDiagonalColor { get; set; }
        BorderDiagonal BorderDiagonal { get; set; }
        #endregion

        #region Fill
        FillPattern FillPattern { get; set; }
        RGBColor FillBackgroundColor { get; set; }
        RGBColor FillForegroundColor { get; set; }
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
