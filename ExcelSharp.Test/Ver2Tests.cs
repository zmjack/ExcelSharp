using NStandard;
using Xunit;

namespace ExcelSharp.Test;

public class Ver2Tests
{
    [Fact]
    public void Test1()
    {
        var sheet = new SpreadSheet("Sheet1");

        sheet[(0, 0)].Value = "00";
        foreach (var (rowIndex, row) in new[] { 0, 1, 2 }.Pairs())
        {
            foreach (var (colIndex, col) in new[] { 1, 2, 3 }.Pairs())
            {
                sheet[(rowIndex, colIndex + 1)].Value = $"{row}{col}";
            }
        }

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
