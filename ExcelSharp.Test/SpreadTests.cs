using ExcelSharp.NPOI;
using System;
using Xunit;

namespace ExcelSharp.Test;

public class SpreadTests
{
    [Fact]
    public void Test()
    {
        var book = new SpreadBook
        {
            new SpreadSheet("Sheet1")
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
                            new Hori { "b", new Vert { 1, 2 } },
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
            }
        };
        var ebook = book.ToExcelBook();
        ebook.SaveAs($"{Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)}\\excelsharp.xlsx");
    }

}
