using Xunit;

namespace ExcelSharp.Test;

public class Ver1Tests
{
    [Fact]
    public void Test1()
    {
        var sheet = new SpreadSheet("Sheet1");

        sheet[(0, 0)].Value = "00";

        sheet[(0, 1)].Value = "01";
        sheet[(0, 2)].Value = "02";
        sheet[(0, 3)].Value = "03";

        sheet[(1, 1)].Value = "11";
        sheet[(1, 2)].Value = "12";
        sheet[(1, 3)].Value = "13";

        sheet[(2, 1)].Value = "21";
        sheet[(2, 2)].Value = "22";
        sheet[(2, 3)].Value = "23";

        sheet[(0, 0), (2, 0)].Merge();

        foreach (var cell in sheet.GetCells())
        {
            if (cell.Value is not null)
            {
                cell.Style = SpreadStyle.Default with
                {
                    Border = true,
                    Center = true,
                };
            }
        }

        var htmlTable = new HtmlTable(sheet);
        var html = htmlTable.ToHtml();
    }
}
