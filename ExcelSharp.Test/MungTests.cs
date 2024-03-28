using ExcelSharp.NPOI;
using Xunit;

namespace ExcelSharp.Test;

public class SpreadTests
{
    [Fact]
    public void Test()
    {
        var mungBook = new SpreadBook
        {
            new SpreadSheet("Sheet1")
            {
                new Vert("main", [
                    SpreadStyle.Table with { BackgroundColor = new RgbaColor(0xcccccc) },
                    SpreadStyle.Table with { BackgroundColor = new RgbaColor(0x999999) },
                ])
                {
                    new Hori(SpreadStyle.Table with
                    {
                        BackgroundColor = new RgbaColor(0xcc66cc),
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
                                BackgroundColor = new RgbaColor(0xcc66cc),
                                Format = "#,##0.00"
                            })
                            {
                                120d,
                            },
                            "c",
                        },
                        new Vert
                        {
                            "e",
                        },
                        "f",
                        new Vert
                        {
                            "11",
                            "22",
                            "33",
                            "44",
                        },
                    },
                    new Hori()
                    {
                        "A",
                        new Vert
                        {
                            "d",
                            "e",
                        },
                    },
                    new Hori()
                    {
                        "A",
                        new Vert([
                            SpreadStyle.Table with { BackgroundColor = new RgbaColor(0xff0000) },
                            SpreadStyle.Table with { BackgroundColor = new RgbaColor(0x0000ff) },
                        ])
                        {
                            "d",
                            "e",
                        },
                    },
                }
            }
        };
        var book = mungBook.ToExcelBook();
    }

}
