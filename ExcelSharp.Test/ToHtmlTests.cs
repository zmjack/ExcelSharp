using ExcelSharp.Abstract;
using NStandard;
using System;
using System.Collections.Generic;
using System.Text;
using Xunit;

namespace ExcelSharp.Test
{
    public class ToHtmlTests
    {
        [Fact]
        public void Test1()
        {
            var table = new HtmlTable();
            table.Rows.Add(new HtmlTableRow().Then(row =>
            {
                row.Cells.Add(new HtmlTableCell { RowSpan = 2, Text = "123" });
                row.Cells.Add(new HtmlTableCell { Text = "23" });
                row.Cells.Add(new HtmlTableCell { Text = "234" });
                row.Cells.Add(new HtmlTableCell { Text = "234" });
                row.Cells.Add(new HtmlTableCell { Text = "345" });
            }));
            table.Rows.Add(new HtmlTableRow().Then(row =>
            {
                row.Cells.Add(new HtmlTableCell { Hidden = true });
                row.Cells.Add(new HtmlTableCell { Text = "23" });
                row.Cells.Add(new HtmlTableCell { Text = "234" });
                row.Cells.Add(new HtmlTableCell { Text = "234" });
                row.Cells.Add(new HtmlTableCell { Text = "345" });
            }));

            var html = table.ToHtml();
        }

    }
}
