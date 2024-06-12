using ExcelSharp.NPOI;
using NPOI.SS.Formula.Functions;
using NStandard;
using System;
using Xunit;

namespace ExcelSharp.Test;

public class ToHtmlTests
{
    [Fact]
    public void Test1()
    {
        var sheet = new SpreadSheet("Sheet1");
        sheet[(0, 0)].Pipe(cell => { cell.Value = "00"; });
        sheet[(0, 1)].Pipe(cell => { cell.Value = "01"; });
        sheet[(0, 2)].Pipe(cell => { cell.Value = "02"; });
        sheet[(0, 3)].Pipe(cell => { cell.Value = "03"; });
        sheet[(0, 4)].Pipe(cell => { cell.Value = "04"; });

        sheet[(1, 0)].Pipe(cell => { cell.Value = "10"; });
        sheet[(1, 1)].Pipe(cell => { cell.Value = "11"; });
        sheet[(1, 2)].Pipe(cell => { cell.Value = "12"; });
        sheet[(1, 3)].Pipe(cell => { cell.Value = "13"; });
        sheet[(1, 4)].Pipe(cell => { cell.Value = "14"; });

        sheet[(2, 0)].Pipe(cell => { cell.Value = "20"; });
        sheet[(2, 1)].Pipe(cell => { cell.Value = "11"; });
        sheet[(2, 2)].Pipe(cell => { cell.Value = "22"; });
        sheet[(2, 3)].Pipe(cell => { cell.Value = "23"; });
        sheet[(2, 4)].Pipe(cell => { cell.Value = "24"; });

        sheet[(0, 0), (1, 0)].Merge();

        foreach (var cell in sheet.GetCells())
        {
            if (cell.Value is not null)
            {
                cell.Style = SpreadStyle.Default with
                {
                    Border = true,
                    HoriAlign = HoriAlign.Center,
                };
            }
        }

        var htmlTable = new HtmlTable(sheet);
        var html = htmlTable.ToHtml();
    }

    [Fact]
    public void ThemeTest()
    {
        var sheet = new SpreadSheet("Sheet1");
        for (int row = 0; row < 6; row++)
        {
            for (int col = 0; col < 10; col++)
            {
                sheet[(row, col)].Grid =
                    new Hori(SpreadStyle.Default with
                    {
                        BackgroundColor = ExcelColor.Theme(row, col),
                    })
                    {
                        $"{row},{col}"
                    };
            }
        }

        sheet[(7, 0)].Value = "Standard";
        for (int col = 0; col < 10; col++)
        {
            sheet[(8, col)].Grid =
                new Hori(SpreadStyle.Default with
                {
                    BackgroundColor = ExcelColor.Standard(0, col),
                })
                {
                    $"{8},{col}"
                };
        }

        var html = new HtmlTable(sheet).ToHtml();
    }

}
