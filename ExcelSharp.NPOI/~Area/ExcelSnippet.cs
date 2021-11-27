using System;

namespace ExcelSharp.NPOI
{
    [Obsolete("Plan to delete.")]
    public abstract class ExcelSnippet
    {
        public static TemporaryExcelSnippet Create(Func<SheetRange> print) => new(print);

        public abstract SheetRange Print();
    }
}
