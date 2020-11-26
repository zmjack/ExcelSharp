using NStandard;
using System;
using System.Collections.Generic;
using System.Linq;
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
            if (_start.Row > start.Row) _start.Row = start.Row;
            if (_start.Col > start.Col) _start.Col = start.Col;

            if (_end.Row < end.Row) _end.Row = end.Row;
            if (_end.Col < end.Col) _end.Col = end.Col;
        }

        public SheetRange GetRange() => new SheetRange(Sheet, _start, _end);

        public override void Disposing()
        {
            Sheet.Cursor = new Cursor(End.Row + 1, Start.Col);
            var outterArea = Scopes.Skip(1).FirstOrDefault();
            if (outterArea is not null) outterArea.Update(_start, _end);
        }
    }
}
