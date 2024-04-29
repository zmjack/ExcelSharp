using NStandard;
using System.ComponentModel;

namespace ExcelSharp.NPOI;

[EditorBrowsable(EditorBrowsableState.Never)]
public static class SpreadExtensions
{
    public static ExcelBook ToExcelBook(this ISpreadBook @this)
    {
        var book = new ExcelBook(ExcelVersion.Excel2007);
        foreach (var spsheet in @this.Sheets.Values)
        {
            var sheet = book.CreateSheet(spsheet.Name);
            sheet.PrintSpreadSheet(spsheet);
            sheet.AutoSize([.. RangeEx.Create(0, spsheet.ColumnLength)]);
        }
        return book;
    }

    public static ExcelBook ToExcelBook(this ISpreadBook @this, Cursor cursor)
    {
        var book = new ExcelBook(ExcelVersion.Excel2007);
        foreach (var spsheet in @this.Sheets.Values)
        {
            var sheet = book.CreateSheet(spsheet.Name);
            sheet.PrintSpreadSheet(spsheet, cursor);
            sheet.AutoSize([.. RangeEx.Create(0, spsheet.ColumnLength)]);
        }
        return book;
    }
}
