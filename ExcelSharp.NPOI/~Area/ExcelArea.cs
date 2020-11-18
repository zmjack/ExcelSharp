using NStandard;
using System;
using System.Collections.Generic;
using System.Text;

namespace ExcelSharp.NPOI
{
    public class ExcelArea : Scope<ExcelArea>
    {
        public ExcelSheet Sheet { get; private set; }

        internal Cursor _start;
        internal Cursor _end;
        public Cursor Start { get => _start; }
        public Cursor End { get => _end; }

        internal ExcelArea(ExcelSheet sheet)
        {
            Sheet = sheet;
            _start = _end = sheet.Cursor;
        }

        public void Update(Cursor start, Cursor end)
        {
            var current = Current;
            if (current is not null)
            {
                if (current._start.Row > start.Row) current._start.Row = start.Row;
                if (current._start.Col > start.Col) current._start.Col = start.Col;

                if (current._end.Row < end.Row) current._end.Row = end.Row;
                if (current._end.Col < end.Col) current._end.Col = end.Col;
            }
        }

        public SheetRange GetRange() => new SheetRange(Sheet, _start, _end);

        public override void Disposing()
        {
            Sheet.Cursor = new Cursor(End.Row + 1, Start.Col);
        }
    }
}
