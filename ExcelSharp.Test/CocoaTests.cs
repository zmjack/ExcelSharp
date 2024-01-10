using Cocoa;
using ExcelSharp.NPOI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Xunit;

namespace ExcelSharp.Test;

public class CocoaTests
{
    [Fact]
    public void Test()
    {
        var coBook = new CoBook
        {
            new CoSheet("Sheet1")
            {
                new CoGrid.Vert("main", [
                    CoStyle.Table with { BackgroundColor = new RgbaColor(0xcccccc) },
                    CoStyle.Table with { BackgroundColor = new RgbaColor(0x999999) },
                ])
                {
                    new CoGrid.Hori(CoStyle.Table with
                    {
                        BackgroundColor = new RgbaColor(0xcc66cc),
                    })
                    {
                        new CoGrid.Vert
                        {
                            new CoGrid.Hori { "a" },
                            new CoGrid.Hori { "b", new CoGrid.Vert { 1, 2 } },
                        },
                        new CoGrid.Vert
                        {
                            new CoGrid.Hori(CoStyle.Table with
                            {
                                BackgroundColor = new RgbaColor(0xcc66cc),
                                Format = "#,##0.00"
                            })
                            {
                                120d,
                            },
                            "c",
                        },
                        new CoGrid.Vert
                        {
                            "e",
                        },
                        "f",
                        new CoGrid.Vert
                        {
                            "11",
                            "22",
                            "33",
                            "44",
                        },
                    },
                    new CoGrid.Hori()
                    {
                        "A",
                        new CoGrid.Vert
                        {
                            "d",
                            "e",
                        },
                    },
                    new CoGrid.Hori()
                    {
                        "A",
                        new CoGrid.Vert([
                            CoStyle.Table with { BackgroundColor = new RgbaColor(0xff0000) },
                            CoStyle.Table with { BackgroundColor = new RgbaColor(0x0000ff) },
                        ])
                        {
                            "d",
                            "e",
                        },
                    },
                }
            }
        };
        var book = coBook.ToExcelBook();
    }

}
