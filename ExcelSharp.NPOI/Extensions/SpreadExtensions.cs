using NStandard;
using System.ComponentModel;

namespace ExcelSharp.NPOI;

[EditorBrowsable(EditorBrowsableState.Never)]
public static class SpreadExtensions
{
    public static ExcelBook ToExcelBook(this ISpreadBook @this, bool autoSize = true)
    {
        var book = new ExcelBook(ExcelVersion.Excel2007);
        foreach (var spsheet in @this.Sheets)
        {
            var sheet = book.CreateSheet(spsheet.Name);
            sheet.PrintSpreadSheet(spsheet);

            if (autoSize)
            {
                sheet.AutoSize([.. RangeEx.Create(0, spsheet.ColumnLength)]);
            }
        }
        return book;
    }

    public static ExcelBook ToExcelBook(this ISpreadBook @this, Cursor cursor, bool autoSize = true)
    {
        var book = new ExcelBook(ExcelVersion.Excel2007);
        foreach (var spsheet in @this.Sheets)
        {
            var sheet = book.CreateSheet(spsheet.Name);
            sheet.PrintSpreadSheet(spsheet, cursor);

            if (autoSize)
            {
                sheet.AutoSize([.. RangeEx.Create(0, spsheet.ColumnLength)]);
            }
        }
        return book;
    }
}
