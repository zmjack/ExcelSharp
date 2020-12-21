using NStandard;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace ExcelSharp
{
    public class RichArea
    {
        public RichTable Table { get; private set; }

        internal Cursor _start;
        internal Cursor _end;
        public Cursor Start { get => _start; }
        public Cursor End { get => _end; }

        internal RichArea(RichTable table, Cursor start, Cursor end)
        {
            Table = table;
            _start = start;
            _end = end;
        }

        public void Update(Cursor start, Cursor end)
        {
            if (_start.Row > start.Row) _start.Row = start.Row;
            if (_start.Col > start.Col) _start.Col = start.Col;
            if (_end.Row < end.Row) _end.Row = end.Row;
            if (_end.Col < end.Col) _end.Col = end.Col;
        }

        public void Extend(RichArea area) => Update(area._start, area._end);

        public IEnumerable<RichCell> Cells
        {
            get
            {
                for (var row = _start.Row; row <= _end.Row; row++)
                {
                    for (var col = _start.Col; col <= _end.Col; col++)
                    {
                        yield return Table.Cell((row, col));
                    }
                }
            }
        }

        public void Merge() => Table.Merge(_start, _end);

        public RichArea Column(int index)
        {
            return new RichArea(Table, (Start.Row, Start.Col + index), (End.Row, Start.Col + index));
        }

        /// <summary>
        /// Merge cells that has same value.
        ///     You can use [[ ]] to identifier cells, the result of merged cell will ignore the value which is in [[ ]].
        /// </summary>
        /// <param name="offsetCols"></param>
        public void SmartMerge(params int[] offsetCols)
        {
            var col = Start.Col + offsetCols[0];

            string take = null;
            int mergeStart = Start.Row;

            for (int takeRow = mergeStart; takeRow <= End.Row; takeRow++)
            {
                var value = Table[(takeRow, col)].Text;
                if (value != take)
                {
                    if (takeRow - mergeStart > 1) SmartMergeVertical(mergeStart, takeRow - 1, col, offsetCols);

                    mergeStart = takeRow;
                    take = value;
                }
                else continue;
            }

            if (End.Row > mergeStart) SmartMergeVertical(mergeStart, End.Row, col, offsetCols);

            //TODO: Remove identifier(type will be changed -> unchange)
            var regex_matchId = new Regex(@"^\[\[.+?\]\](.*)$");
            foreach (var colIndex in offsetCols)
            {
                foreach (var cell in Column(colIndex).Cells)
                {
                    var value = cell.Value;
                    if (value is string && !(value as string).IsNullOrWhiteSpace())
                    {
                        var match = regex_matchId.Match(value as string);
                        if (match.Success)
                        {
                            if (double.TryParse(match.Groups[1].Value, out double dvalue)) cell.Value = dvalue;
                            else cell.Value = match.Groups[1].Value;
                        }
                    }
                }
            }
        }

        public void SmartColMerge()
        {
            void InnerMerge(int row, int startCol, int endCol, string value)
            {
                value ??= "";

                if (endCol - startCol > 0) Table.Merge((row, startCol), (row, endCol));

                var regex = new Regex(@"^\[\[.+?\]\](.*)$");
                var match = regex.Match(value);

                if (match.Success)
                {
                    var cell = Table[(row, startCol)];
                    var svalue = match.Groups[1].Value;
                    if (!svalue.IsNullOrWhiteSpace())
                    {
                        if (double.TryParse(svalue, out double dvalue)) cell.Value = dvalue;
                        else cell.Value = svalue;
                    }
                    else cell.Value = svalue;
                }
            }

            for (int row = Start.Row; row <= End.Row; row++)
            {
                string take = null;
                int startCol = Start.Col;
                int takeCol = startCol;

                for (; takeCol <= End.Col; takeCol++)
                {
                    var value = Table[(row, takeCol)].Text;
                    if (value == take) continue;

                    InnerMerge(row, startCol, takeCol - 1, take);

                    startCol = takeCol;
                    take = value;
                }

                InnerMerge(row, startCol, takeCol - 1, take);
            }
        }

        private void SmartMergeVertical(int mergeStart, int mergeEnd, int col, int[] offsetCols)
        {
            Table.Merge((mergeStart, col), (mergeEnd, col));
            if (offsetCols.Length > 1)
            {
                new RichArea(Table, (mergeStart, col + offsetCols[1]), (mergeEnd, End.Col))
                    .SmartMerge(offsetCols.Slice(1).Select(_col => _col - (offsetCols[1] - offsetCols[0])).ToArray());
            }
        }

    }
}
