using ExcelSharp.NPOI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelSharp
{
    public class RichRow
    {
        public RichTable Table { get; private set; }
        public int Index { get; internal set; }
        private Dictionary<int, RichCell> _innerCells = new Dictionary<int, RichCell>();

        internal RichRow(RichTable table, int index)
        {
            Table = table;
            Index = index;
        }

        public IEnumerable<RichCell> Cells
        {
            get
            {
                var maxIndex = Table.MaxColumnIndex;
                for (var i = 0; i <= maxIndex; i++) yield return Cell(i);
            }
        }

        public RichCell this[int index] => Cell(index);

        public RichCell Cell(int index)
        {
            if (_innerCells.ContainsKey(index)) return _innerCells[index];
            else
            {
                var cell = new RichCell(this, index);
                _innerCells.Add(index, cell);
                if (Table.MaxColumnIndex < index) Table.MaxColumnIndex = index;
                return cell;
            }
        }

    }
}
