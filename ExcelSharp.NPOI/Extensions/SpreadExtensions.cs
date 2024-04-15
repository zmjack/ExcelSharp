using NStandard;
using System.ComponentModel;

namespace ExcelSharp.NPOI;

[EditorBrowsable(EditorBrowsableState.Never)]
public static class SpreadExtensions
{
    public static ExcelBook ToExcelBook(this ISpreadBook @this)
    {
        var book = new ExcelBook(ExcelVersion.Excel2007);
        foreach (var mungSheet in @this.Sheets)
        {
            var sheet = book.CreateSheet(mungSheet.Name);
            sheet.PrintSpreadSheet(mungSheet);
            sheet.AutoSize([.. RangeEx.Create(0, mungSheet.ColumnLength)]);
        }
        return book;
    }

    public static ExcelBook ToExcelBook(this ISpreadBook @this, Cursor cursor)
    {
        var book = new ExcelBook(ExcelVersion.Excel2007);
        foreach (var mungSheet in @this.Sheets)
        {
            var sheet = book.CreateSheet(mungSheet.Name);
            sheet.PrintSpreadSheet(mungSheet, cursor);
            sheet.AutoSize([.. RangeEx.Create(0, mungSheet.ColumnLength)]);
        }
        return book;
    }
}
