using NStandard;
using NStandard.Flows;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace ExcelSharp
{
    public class RichTable
    {
        private Dictionary<int, UniversalTableRow> _innerRows = new Dictionary<int, UniversalTableRow>();
        internal int MaxColumnIndex = -1;

        public IEnumerable<UniversalTableRow> Rows
        {
            get
            {
                var keys = _innerRows.Keys;
                var maxKey = keys.Any() ? keys.Max() : 0;
                for (var i = 0; i <= maxKey; i++) yield return Row(i);
            }
        }
        public IEnumerable<UniversalTableCell> Cells
        {
            get
            {
                foreach (var row in Rows)
                {
                    foreach (var cell in row.Cells) yield return cell;
                }
            }
        }

        public UniversalTableCell this[Cursor cursor] => Cell(cursor);
        public UniversalTableArea this[Cursor start, Cursor end] => Area(start, end);

        public UniversalTableArea Area(Cursor start, Cursor end) => new UniversalTableArea(this, start, end);

        public UniversalTableRow Row(int index)
        {
            if (_innerRows.ContainsKey(index)) return _innerRows[index];
            else
            {
                var row = new UniversalTableRow(this);
                _innerRows.Add(index, row);
                return row;
            }
        }

        public UniversalTableCell Cell(Cursor cursor) => Row(cursor.Row).Cell(cursor.Col);

        public UniversalTableBrush GetBrush() => new UniversalTableBrush(this, (0, 0), (0, 0));
        public UniversalTableBrush GetBrush(Cursor cursor) => new UniversalTableBrush(this, cursor, cursor);
        public UniversalTableBrush GetBrush(Cursor start, Cursor end) => new UniversalTableBrush(this, start, end);

        public RichTable Merge(Cursor start, Cursor end)
        {
            var rowSpan = end.Row - start.Row + 1;
            var colSpan = end.Col - start.Col + 1;

            for (int row = start.Row; row <= end.Row; row++)
            {
                for (int col = start.Col; col <= end.Col; col++)
                {
                    var cell = Cell((row, col));
                    cell.Ignored = true;
                    cell.RowSpan = cell.ColSpan = 0;
                    cell.RowOffset = row - start.Row;
                    cell.ColOffset = col - start.Col;
                }
            }

            var startCell = Cell(start);
            startCell.Ignored = false;
            startCell.RowSpan = rowSpan;
            startCell.ColSpan = colSpan;

            return this;
        }

    }
}
