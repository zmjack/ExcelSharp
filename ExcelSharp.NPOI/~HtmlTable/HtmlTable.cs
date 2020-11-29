using ExcelSharp.NPOI;
using NStandard;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelSharp.Abstract
{
    public class HtmlTable
    {
        private string Props(params string[] props)
        {
            return props.Where(prop => !prop.IsNullOrWhiteSpace()).Select(prop => $" {prop}").Join("");
        }

        public string Class { get; set; }
        public string Style { get; set; }
        public List<HtmlTableRow> Rows { get; } = new List<HtmlTableRow>();

        public string ToHtml()
        {
            var sb = new StringBuilder();
            sb.AppendLine(@"<table class=""table"" style=""border-collapse:collapse"">");
            sb.AppendLine(@"<tbody>");

            foreach (var row in Rows)
            {
                sb.AppendLine(@"<tr>");
                foreach (var cell in row.Cells.Where(x => !x.Hidden))
                {
                    sb.AppendLine($@"<td{Props(cell.RowSpanProp, cell.ColSpanProp, cell.StyleProp)}>{cell.Text}</td>");
                }
                sb.AppendLine(@"</tr>");
            }

            sb.AppendLine(@"</tbody>");
            sb.AppendLine(@"</table>");

            return sb.ToString();
        }

    }
}
