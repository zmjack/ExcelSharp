using System.Linq;
using Xunit;

namespace ExcelSharp.Test;

public class Ver3Tests
{
    [Fact]
    public void Test1()
    {
        var sheet = new SpreadSheet("Sheet1")
        {
            new Vert(SpreadStyle.Default with
            {
                Border = true,
                Center = true,
            })
            {
                new Hori
                {
                    "00",

                    new Vert
                    {
                        from row in new[] { 0, 1, 2 }
                        select new Hori
                        {
                            from col in new[] { 1, 2, 3 }
                            select $"{row}{col}"
                        }
                    }
                }
            }
        };

        var htmlTable = new HtmlTable(sheet);
        var html = htmlTable.ToHtml();
    }
}
