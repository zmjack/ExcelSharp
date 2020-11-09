using NStandard;
using System;
using System.Collections.Generic;
using System.Text;

namespace ExcelSharp.NPOI
{
    public class ExcelArea : Scope<ExcelArea>
    {
        public ExcelSheet Sheet { get; private set; }
        public Cursor Start { get; private set; }
        public Cursor End { get; private set; }

        public ExcelArea(ExcelSheet sheet)
        {
            Sheet = sheet;
            Start = sheet.Cursor;
        }

        public SheetRange GetRange() => new SheetRange(Sheet, Start, Sheet.Cursor);

        public override void Disposing()
        {
            Sheet.Cursor = new Cursor(End.Row + 1, Start.Col);
        }
    }
}
