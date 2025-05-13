using ExcelSharp.NPOI;
using System;
using Xunit;

namespace ExcelSharp.Test;

public class SpreadTests
{
    [Fact]
    public void Test()
    {
        var sheet = new SpreadSheet("Sheet1")
        {
            new Vert("main",
            [
                SpreadStyle.Table with
                {
                    BackgroundColor = ExcelColor.Theme(1, 7),
                },
                SpreadStyle.Table with
                {
                    BackgroundColor = ExcelColor.Theme(0, 2),
                },
            ])
            {
                new Hori(SpreadStyle.Table with
                {
                    BackgroundColor = ExcelColor.Theme(2, 5),
                })
                {
                    new Vert
                    {
                        new Hori { "a" },
                        new Hori
                        {
                            "b",
                            new Vert(SpreadStyle.Table with
                            {
                                BackgroundColor = ExcelColor.Theme(2, 5),
                                Round = 2,
                            })
                            {
                                1,
                                1.2345,
                            }
                        },
                    },
                    new Vert
                    {
                        new Hori(SpreadStyle.Table with
                        {
                            BackgroundColor = ExcelColor.Theme(2, 5),
                            Format = "#,##0.00"
                        })
                        {
                            120d,
                        },
                        "c",
                    },
                    new Vert
                    {
                        "d",
                    },
                    "e",
                    new Vert
                    {
                        "11",
                        "22",
                        "33",
                        "44",
                    },
                },
                new Hori
                {
                    "A",
                    new Vert
                    {
                        "f",
                        "g",
                    },
                },
                new Hori
                {
                    "B",
                    new Vert(
                    [
                        SpreadStyle.Table with
                        {
                            BackgroundColor = ExcelColor.Theme(2, 9),
                        },
                        SpreadStyle.Table with
                        {
                            BackgroundColor = ExcelColor.Theme(2, 8),
                        },
                    ])
                    {
                        "h",
                        "i",
                    },
                },
            }
        };

        var book = new SpreadBook
        {
            sheet
        };
        var ebook = book.ToExcelBook();
        var file = $"{Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)}\\excelsharp.xlsx";
        ebook.SaveAs(file);
    }

    [Fact]
    public void Test2()
    {
        var sheet = new SpreadSheet("Sheet1")
        {
            new Vert
            {
                "",
            }
        };

        var book = new SpreadBook
        {
            sheet
        };
        var ebook = book.ToExcelBook();
        var file = $"{Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)}\\excelsharp2.xlsx";
        ebook.SaveAs(file);
    }

}
