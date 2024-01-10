using NStandard;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelSharp.NPOI
{
    [EditorBrowsable(EditorBrowsableState.Never)]
    public static class CocoaExtensions
    {
        public static ExcelBook ToExcelBook(this ICoBook @this)
        {
            var book = new ExcelBook(ExcelVersion.Excel2007);
            foreach (var coSheet in @this.Sheets)
            {
                var sheet = book.CreateSheet(coSheet.Name);
                sheet.ExtendPrintLine(coSheet);
                sheet.AutoSize([.. RangeEx.Create(0, coSheet.ColumnLength)]);
            }
            return book;
        }
    }
}
