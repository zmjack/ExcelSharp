using System;
using System.Collections.Generic;
using System.Text;

namespace ExcelSharp.NPOI
{
    public abstract class ExcelSnippet
    {
        public static TemporaryExcelSnippet Create(Func<SheetRange> print) => new TemporaryExcelSnippet(print);

        public abstract SheetRange Print();
    }
}
