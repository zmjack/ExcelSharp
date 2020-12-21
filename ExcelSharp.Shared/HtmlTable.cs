using ExcelSharp.NPOI;
using NStandard;
using NStandard.Flows;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelSharp
{
    public class HtmlTable
    {
        public RichTable Table { get; }
        public string Class { get; set; }
        public string Style { get; set; }

        public HtmlTable(RichTable table, string @class = "table", string style = "border-collapse:collapse;")
        {
            Table = table;
            Class = @class;
            Style = style;
        }

        private static string Props(params string[] props)
        {
            return props.Where(prop => !prop.IsNullOrWhiteSpace()).Select(prop => $" {prop}").Join("");
        }
        private static string RowSpanProp(RichCell cell) => cell.RowSpan.For(span => span > 1 ? $@"rowspan=""{span}""" : "");
        private static string ColSpanProp(RichCell cell) => cell.ColSpan.For(span => span > 1 ? $@"colspan=""{span}""" : "");

        private static string StyleProp(RichCell cell)
        {
            if (cell.Ignored) return "";

            var dict = new Dictionary<string, string>();
            var style = cell.Style;

            if (style.BackgroundColor.HasValue) dict.Add("background-color", $"#{style.BackgroundColor:x6}");
            if (style.Color.HasValue) dict.Add("color", $"#{style.Color:x6}");
            if (style.FontFamily is not null) dict.Add("font-family", $"{style.FontFamily}");
            if (style.FontSize.HasValue) dict.Add("font-size", $"{style.FontSize}px");
            if (style.Bold.HasValue && style.Bold.Value) dict.Add("font-weight", $"bold");
            if (style.BorderTop.HasValue && style.BorderTop.Value) dict.Add("border-top", $"1px solid #000000");
            if (style.BorderBottom.HasValue && style.BorderBottom.Value) dict.Add("border-bottom", $"1px solid #000000");
            if (style.BorderLeft.HasValue && style.BorderLeft.Value) dict.Add("border-left", $"1px solid #000000");
            if (style.BorderRight.HasValue && style.BorderRight.Value) dict.Add("border-right", $"1px solid #000000");
            if (!style.TextAlign.IsNullOrWhiteSpace()) dict.Add("text-align", $"{style.TextAlign}");
            if (!style.VerticalAlign.IsNullOrWhiteSpace()) dict.Add("vertical-align", $"{style.VerticalAlign}");

            return $@"style=""{dict.Select(pair => $"{pair.Key}:{pair.Value}").Join(";")}""";
        }

        public string ToHtml()
        {
            var sb = new StringBuilder();
            sb.AppendLine($@"<table class=""{Class}"" style=""{Style}"">");
            sb.AppendLine(@"<tbody>");

            foreach (var row in Table.Rows)
            {
                sb.AppendLine(@"<tr>");
                foreach (var cell in row.Cells.Where(x => !x.Ignored))
                {
                    sb.AppendLine($@"<td{Props(RowSpanProp(cell), ColSpanProp(cell), StyleProp(cell))}>");
                    if (!cell.Comment.IsNullOrWhiteSpace())
                    {
                        sb.AppendLine($@"
<span class=""excel-comment"">
    <div class=""excel-content"">
        {cell.Comment.Flow(StringFlow.HtmlEncode).Replace("\r\n", "<br/>")}
    </div>
    <span class=""excel-text"">{cell.Text}</span>
</span>
");
                    }
                    else sb.AppendLine(cell.Text);
                    sb.AppendLine("</td>");
                }
                sb.AppendLine(@"</tr>");
            }

            sb.AppendLine(@"</tbody>");
            sb.AppendLine(@"</table>");

            return sb.ToString();
        }

    }
}
