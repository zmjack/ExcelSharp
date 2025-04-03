using System.ComponentModel;

namespace ExcelSharp.NPOI;

[EditorBrowsable(EditorBrowsableState.Never)]
public static class SpreadExtensions
{
    public static ExcelBook ToExcelBook(this ISpreadBook @this)
    {
        var book = new ExcelBook(ExcelVersion.Excel2007);
        foreach (var spsheet in @this.Sheets)
        {
            var sheet = book.CreateSheet(spsheet.Name);
            sheet.PrintSpreadSheet(spsheet);
        }
        return book;
    }

    public static ExcelBook ToExcelBook(this ISpreadBook @this, Cursor cursor)
    {
        var book = new ExcelBook(ExcelVersion.Excel2007);
        foreach (var spsheet in @this.Sheets)
        {
            var sheet = book.CreateSheet(spsheet.Name);
            sheet.PrintSpreadSheet(spsheet, cursor);
        }
        return book;
    }
}
