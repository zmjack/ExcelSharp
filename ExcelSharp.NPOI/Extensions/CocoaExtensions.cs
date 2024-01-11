using NStandard;
using System.ComponentModel;

namespace ExcelSharp.NPOI
{
    [EditorBrowsable(EditorBrowsableState.Never)]
    public static class MungExtensions
    {
        public static ExcelBook ToExcelBook(this IMungBook @this)
        {
            var book = new ExcelBook(ExcelVersion.Excel2007);
            foreach (var mungSheet in @this.Sheets)
            {
                var sheet = book.CreateSheet(mungSheet.Name);
                sheet.ExtendPrintLine(mungSheet);
                sheet.AutoSize([.. RangeEx.Create(0, mungSheet.ColumnLength)]);
            }
            return book;
        }
    }
}
