using System;

namespace ExcelSharp.NPOI
{
    public class TemporaryExcelSnippet : ExcelSnippet
    {
        private Func<SheetRange> _print;

        public TemporaryExcelSnippet(Func<SheetRange> print)
        {
            _print = print;
        }

        public override SheetRange Print() => _print();
    }
}
