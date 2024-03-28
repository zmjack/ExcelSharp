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
                cell.Style = cell.Style with
                {
                    BorderTop = true,
                    BorderBottom = true,
                    BorderLeft = true,
                    BorderRight = true,
                    HoriAlign = HoriAlign.Center,
                };
            }
        }

        var htmlTable = new HtmlTable(sheet);
        var html = htmlTable.ToHtml();
    }

    [Fact]
    public void Test2()
    {
        var mainTitleStyle = new SpreadStyle
        {
            HoriAlign = HoriAlign.Center,
            VertAlign = VertAlign.Middle,
            BackgroundColor = new RgbaColor(0x538DE7),
            FontFamily = "Microsoft YaHei",
            FontSize = 24,
            Color = new RgbaColor(0xffffff),
            Bold = true,
            BorderTop = true,
            BorderBottom = true,
            BorderLeft = true,
            BorderRight = true,
        };
        var titleStyle = new SpreadStyle
        {
            HoriAlign = HoriAlign.Center,
            VertAlign = VertAlign.Middle,
            BackgroundColor = new RgbaColor(0x538DE7),
            FontFamily = "Microsoft YaHei",
            FontSize = 12,
            Color = new RgbaColor(0xffffff),
            Bold = true,
            BorderTop = true,
            BorderBottom = true,
            BorderLeft = true,
            BorderRight = true,
        };
        var dateStyle = new SpreadStyle
        {
            HoriAlign = HoriAlign.Right,
            VertAlign = VertAlign.Middle,
            FontFamily = "Microsoft YaHei",
            FontSize = 9,
            Bold = true,
        };

        var mungSheet = new SpreadSheet("Sheet1");
        mungSheet["B2"].Pipe(cell =>
        {
            cell.Value = $"{new DateTime(2000, 1, 1):yyyy年MM月}";
            cell.Style = mainTitleStyle with { };
        });

        mungSheet["B5"].Pipe(cell =>
        {
            cell.Value = "生产月份";
            cell.Style = titleStyle with { };
        });
        mungSheet["B5", "D5"].Merge();
        mungSheet["B6"].Pipe(cell =>
        {
            cell.Value = "货龄";
            cell.Style = titleStyle with { };
        });
        mungSheet["B6", "D6"].Merge();
        mungSheet["B7"].Pipe(cell =>
        {
            cell.Value = "合计数量";
            cell.Style = titleStyle with { };
        });
        mungSheet["B7", "D7"].Merge();
        mungSheet["B8"].Pipe(cell =>
        {
            cell.Value = "合计占比（%）";
            cell.Style = titleStyle with { };
        });
        mungSheet["B8", "D8"].Merge();

        var excel = new ExcelBook(ExcelVersion.Excel2007);
        excel.CreateSheet("sheet1");
        var sheet = excel[0];

        using (var area = sheet.BeginArea("E7"))
        {
            sheet.ExtendPrintLine(mungSheet);
        }
        excel.SaveAs("D:\\tmp\\11.xlsx");

        var html = new HtmlTable(mungSheet).ToHtml();
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

        sheet[(7, 0)].Value = "标准色";
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
