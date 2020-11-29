using NStandard;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelSharp.Abstract
{
    public class HtmlTableCell
    {
        public bool Hidden { get; set; }
        public bool IsHeadCell { get; set; }

        public int? RowSpan { get; set; }
        public int? ColSpan { get; set; }

        public int? BackgroundColor { get; set; }
        public int? Color { get; set; }
        public int? FontSize { get; set; }
        public bool? Bold { get; set; }
        public string TextAlign { get; set; }

        public string Text { get; set; }
        public bool? BorderTop { get; set; }
        public bool? BorderBottom { get; set; }
        public bool? BorderLeft { get; set; }
        public bool? BorderRight { get; set; }

        public string StyleProp
        {
            get
            {
                if (Hidden) return "";

                var dict = new Dictionary<string, string>();
                if (BackgroundColor.HasValue) dict.Add("background-color", $"#{BackgroundColor:x6}");
                if (Color.HasValue) dict.Add("color", $"#{Color:x6}");

                if (FontSize.HasValue) dict.Add("font-size", $"{FontSize}px");

                if (Bold.HasValue && Bold.Value) dict.Add("font-weight", $"bold");
                if (BorderTop.HasValue && BorderTop.Value) dict.Add("border-top", $"1px solid #000000");
                if (BorderBottom.HasValue && BorderBottom.Value) dict.Add("border-bottom", $"1px solid #000000");
                if (BorderLeft.HasValue && BorderLeft.Value) dict.Add("border-left", $"1px solid #000000");
                if (BorderRight.HasValue && BorderRight.Value) dict.Add("border-right", $"1px solid #000000");

                if (!TextAlign.IsNullOrWhiteSpace()) dict.Add("text-align", $"{TextAlign}");

                return $@"style=""{dict.Select(pair => $"{pair.Key}:{pair.Value}").Join(";")}""";
            }
        }

        internal string RowSpanProp => RowSpan.HasValue && RowSpan > 1 ? $@"rowspan=""{RowSpan}""" : "";
        internal string ColSpanProp => ColSpan.HasValue && ColSpan > 1 ? $@"colspan=""{ColSpan}""" : "";

    }
}
