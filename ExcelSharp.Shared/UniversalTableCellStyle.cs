using ExcelSharp.NPOI;
using System;
using System.Collections.Generic;
using System.Text;

namespace ExcelSharp
{
    public record UniversalTableCellStyle
    {
        public UniversalTableCellStyle() { }

        public bool IsHeadCell { get; set; }

        public int? BackgroundColor { get; set; }
        public int? Color { get; set; }
        public int? FontSize { get; set; }
        public string FontFamily { get; set; }
        public bool? Bold { get; set; }
        public string TextAlign { get; set; }
        public string VerticalAlign { get; set; }

        public bool? BorderTop { get; set; }
        public bool? BorderBottom { get; set; }
        public bool? BorderLeft { get; set; }
        public bool? BorderRight { get; set; }
    }
}
