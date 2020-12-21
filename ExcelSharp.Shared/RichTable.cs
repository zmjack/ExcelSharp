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
        private Dictionary<int, RichRow> _innerRows = new Dictionary<int, RichRow>();
        public int MaxRowIndex { get; internal set; } = -1;
        public int MaxColumnIndex { get; internal set; } = -1;

        public int RowLength => MaxRowIndex + 1;
        public int ColumnLength => MaxColumnIndex + 1;

        public IEnumerable<RichRow> Rows
        {
            get
            {
                var keys = _innerRows.Keys;
                var maxKey = keys.Any() ? keys.Max() : 0;
                for (var i = 0; i <= maxKey; i++) yield return Row(i);
            }
        }
        public IEnumerable<RichCell> Cells
        {
            get
            {
                foreach (var row in Rows)
                {
                    foreach (var cell in row.Cells) yield return cell;
                }
            }
        }

        public RichCell this[Cursor cursor] => Cell(cursor);
        public RichArea this[Cursor start, Cursor end] => Area(start, end);

        public RichArea Area(Cursor start, Cursor end) => new RichArea(this, start, end);

        public RichRow Row(int index)
        {
            if (_innerRows.ContainsKey(index)) return _innerRows[index];
            else
            {
                var row = new RichRow(this);
                _innerRows.Add(index, row);
                if (MaxRowIndex < index) MaxRowIndex = index;
                return row;
            }
        }

        public RichCell Cell(Cursor cursor) => Row(cursor.Row).Cell(cursor.Col);

        public RichBrush BeginBrush() => new RichBrush(this, (0, 0), (0, 0));
        public RichBrush BeginBrush(Cursor cursor) => new RichBrush(this, cursor, cursor);
        public RichBrush BeginBrush(Cursor start, Cursor end) => new RichBrush(this, start, end);

        public RichTable Merge(Cursor start, Cursor end)
        {
            var rowSpan = end.Row - start.Row + 1;
            var colSpan = end.Col - start.Col + 1;
            var startCell = Cell(start);

            for (int row = start.Row; row <= end.Row; row++)
            {
                for (int col = start.Col; col <= end.Col; col++)
                {
                    var cell = Cell((row, col));
                    cell.Ignored = true;
                    cell.RowSpan = cell.ColSpan = 0;
                    cell.RowOffset = row - start.Row;
                    cell.ColOffset = col - start.Col;
                    cell.Style = startCell.Style;
                }
            }

            startCell.Ignored = false;
            startCell.RowSpan = rowSpan;
            startCell.ColSpan = colSpan;

            return this;
        }

        public RichTable Unmerge(Cursor cursor)
        {
            throw new NotImplementedException();
        }

    }
}
