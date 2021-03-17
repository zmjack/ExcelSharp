using System;

namespace ExcelSharp.NPOI
{
    public abstract class ExcelSnippet
    {
        public static TemporaryExcelSnippet Create(Func<SheetRange> print) => new(print);

        public abstract SheetRange Print();
    }
}
