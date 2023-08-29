using ExcelSharp.NPOI;
using NPOI.SS.Formula.Functions;
using NStandard;
using Richx;
using System;
using Xunit;

namespace ExcelSharp.Test
{
    public class ToHtmlTests
    {
        [Fact]
        public void Test1()
        {
            var table = new RichTable();
            table[(0, 0)].Pipe(cell => { cell.Value = "00"; });
            table[(0, 1)].Pipe(cell => { cell.Value = "01"; });
            table[(0, 2)].Pipe(cell => { cell.Value = "02"; });
            table[(0, 3)].Pipe(cell => { cell.Value = "03"; });
            table[(0, 4)].Pipe(cell => { cell.Value = "04"; });

            table[(1, 0)].Pipe(cell => { cell.Value = "10"; });
            table[(1, 1)].Pipe(cell => { cell.Value = "11"; });
            table[(1, 2)].Pipe(cell => { cell.Value = "12"; });
            table[(1, 3)].Pipe(cell => { cell.Value = "13"; });
            table[(1, 4)].Pipe(cell => { cell.Value = "14"; });

            table[(2, 0)].Pipe(cell => { cell.Value = "20"; });
            table[(2, 1)].Pipe(cell => { cell.Value = "11"; });
            table[(2, 2)].Pipe(cell => { cell.Value = "22"; });
            table[(2, 3)].Pipe(cell => { cell.Value = "23"; });
            table[(2, 4)].Pipe(cell => { cell.Value = "24"; });

            table[(0, 0), (1, 0)].Merge();
            table[(1, 1), (2, 4)].SmartMerge(0, 1, 2, 3);

            var brush = table.BeginBrush((5, 0));
            brush.PrintLine(new object[] { 1, 2, null, "", 5, 6 });
            foreach (var cell in brush.Area.Cells)
            {
                cell.Style = cell.Style with
                {
                    BackgroundColor = new RgbColor(0x00ffff),
                };
            }

            foreach (var cell in table.Cells)
            {
                cell.Style = cell.Style with
                {
                    BorderTop = true,
                    BorderBottom = true,
                    BorderLeft = true,
                    BorderRight = true,

                    TextAlign = RichTextAlignment.Center,
                };
            }

            var htmlTable = new HtmlTable(table);
            var html = htmlTable.ToHtml();
        }

        [Fact]
        public void Test2()
        {
            var mainTitleStyle = new RichStyle
            {
                TextAlign = RichTextAlignment.Center,
                VerticalAlign = RichVerticalAlignment.Middle,
                BackgroundColor = new RgbColor(0x538DE7),
                FontFamily = "Microsoft YaHei",
                FontSize = 24,
                Color = new RgbColor(0xffffff),
                Bold = true,
                BorderTop = true,
                BorderBottom = true,
                BorderLeft = true,
                BorderRight = true,
            };
            var titleStyle = new RichStyle
            {
                TextAlign = RichTextAlignment.Center,
                VerticalAlign = RichVerticalAlignment.Middle,
                BackgroundColor = new RgbColor(0x538DE7),
                FontFamily = "Microsoft YaHei",
                FontSize = 12,
                Color = new RgbColor(0xffffff),
                Bold = true,
                BorderTop = true,
                BorderBottom = true,
                BorderLeft = true,
                BorderRight = true,
            };
            var dateStyle = new RichStyle
            {
                TextAlign = RichTextAlignment.Right,
                VerticalAlign = RichVerticalAlignment.Middle,
                FontFamily = "Microsoft YaHei",
                FontSize = 9,
                Bold = true,
            };

            var uniTable = new RichTable();
            uniTable["B2"].Pipe(cell => { cell.Value = $"{new DateTime(2000, 1, 1):yyyy年MM月各品项货龄（含锁定库存）}"; cell.Style = mainTitleStyle with { }; });

            uniTable["B5"].Pipe(cell => { cell.Value = "生产月份"; cell.Style = titleStyle with { }; });
            uniTable["B5", "D5"].Merge();
            uniTable["B6"].Pipe(cell => { cell.Value = "货龄"; cell.Style = titleStyle with { }; });
            uniTable["B6", "D6"].Merge();
            uniTable["B7"].Pipe(cell => { cell.Value = "合计数量"; cell.Style = titleStyle with { }; });
            uniTable["B7", "D7"].Merge();
            uniTable["B8"].Pipe(cell => { cell.Value = "合计占比（%）"; cell.Style = titleStyle with { }; });
            uniTable["B8", "D8"].Merge();

            var excel = new ExcelBook(ExcelVersion.Excel2007);
            excel.CreateSheet("sheet1");
            var sheet = excel[0];

            using (var area = sheet.BeginArea("E7"))
            {
                sheet.ExtendPrintLine(uniTable);
            }
            excel.SaveAs("D:\\tmp\\11.xlsx");

            var html = new HtmlTable(uniTable).ToHtml();
        }

        [Fact]
        public void ThemeTest()
        {
            var table = new RichTable();
            for (int row = 0; row < 6; row++)
            {
                for (int col = 0; col < 10; col++)
                {
                    table[(row, col)].Value = new RichValue
                    {
                        Style = RichStyle.Default with
                        {
                            BackgroundColor = ExcelColor.Theme(row, col),
                        },
                        Value = $"{row},{col}",
                    };
                }
            }

            table[(7, 0)].Value = "标准色";
            for (int col = 0; col < 10; col++)
            {
                table[(8, col)].Value = new RichValue
                {
                    Style = RichStyle.Default with
                    {
                        BackgroundColor = ExcelColor.Standard(0, col),
                    },
                    Value = $"{8},{col}",
                };
            }

            var html = new HtmlTable(table).ToHtml();
        }

    }
}
