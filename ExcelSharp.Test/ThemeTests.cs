using Xunit;

namespace ExcelSharp.Test;

public class ThemeTests
{
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
                new SpreadValue($"{8},{col}")
                {
                }
                };
        }

        var html = new HtmlTable(sheet).ToHtml();
    }
}
