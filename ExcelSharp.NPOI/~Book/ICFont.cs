﻿using NPOI.SS.UserModel;
using Richx;

namespace ExcelSharp.NPOI
{
    public interface ICFont
    {
        string FontName { get; set; }
        short FontSize { get; set; }

        bool IsItalic { get; set; }
        bool IsStrikeout { get; set; }
        bool IsBold { get; set; }

        FontUnderlineType Underline { get; set; }
        FontSuperScript TypeOffset { get; set; }

        IArgbColor FontColor { get; set; }

        //short Charset { get; set; }
        //short Index { get; }
        //short Boldweight { get; set; }
    }
}
