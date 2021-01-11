using System;

namespace ExcelSharp.NPOI
{
    public abstract class ExcelSnippet
    {
        public static TemporaryExcelSnippet Create(Func<SheetRange> print) => new TemporaryExcelSnippet(print);

        public abstract SheetRange Print();
    }
}
