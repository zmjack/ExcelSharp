using NStandard;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;

namespace ExcelSharp
{
    public class RichTableBrush
    {
        public RichTable Table { get; private set; }
        public RichTableArea Area { get; private set; }
        public Cursor Cursor;

        public RichTableBrush(RichTable table, Cursor start, Cursor end)
        {
            Table = table;
            Cursor = start;
            Area = new RichTableArea(table, start, end);
        }

        public UniversalTableBrush ResetCursorColumn()
        {
            Cursor.Col = Area.Start.Col;
            return this;
        }
        private void RecalculateArea(Cursor start, Cursor end) => Area.Update(start, end);

        public UniversalTableBrush Print(IEnumerable<object> values) => Print(PrintDirection.Horizontal, values.ToArray());
        public UniversalTableBrush Print(object[] values) => Print(PrintDirection.Horizontal, values);
        public UniversalTableBrush Print(PrintDirection direction, IEnumerable<object> values) => Print(direction, values.ToArray());
        public UniversalTableBrush Print(PrintDirection direction, object[] values)
        {
            var startRow = Cursor.Row;
            var startCol = Cursor.Col;
            var length = values.Length;
            var rangeStart = new Cursor { Row = startRow, Col = startCol };
            Cursor rangeEnd;

            if (direction == PrintDirection.Horizontal)
            {
                for (int col = 0; col < values.Length; col++)
                {
                    var valueObj = values[col];
                    if (valueObj is null) continue;
                    Table[(startRow, startCol + col)].Value = valueObj;
                }
                Cursor.Col = startCol + values.Length;
                rangeEnd = new Cursor { Row = startRow, Col = startCol + length - 1 };
            }
            else if (direction == PrintDirection.Vertical)
            {
                for (int row = 0; row < values.Length; row++)
                {
                    var valueObj = values[row];
                    if (valueObj is null) continue;
                    Table[(startRow + row, startCol)].Value = valueObj;
                }
                Cursor.Col = startCol + 1;
                rangeEnd = new Cursor { Row = startRow + length - 1, Col = startCol };
            }
            else throw new NotSupportedException();

            RecalculateArea(rangeStart, rangeEnd);
            return this;
        }
        public UniversalTableBrush Print(object[,] values)
        {
            var startRow = Cursor.Row;
            var startCol = Cursor.Col;
            var rowLength = values.GetLength(0);
            var colLength = values.GetLength(1);

            for (var row = 0; row < rowLength; row++)
            {
                for (var col = 0; col < colLength; col++)
                {
                    var valueObj = values[row, col];
                    if (valueObj is null) continue;
                    Table[(startRow + row, startCol + col)].Value = valueObj;
                }
            }
            Cursor.Col = startCol + colLength;

            var rangeStart = (startRow, startCol);
            var rangeEnd = (startRow + rowLength - 1, startCol + colLength - 1);

            RecalculateArea(rangeStart, rangeEnd);
            return this;
        }
        public UniversalTableBrush Print(object[][] values)
        {
            var startRow = Cursor.Row;
            var startCol = Cursor.Col;
            var rowLength = values.Length;
            var colLength = values.Any() ? values.Max(values1 => values1.Length) : 0;

            for (int row = 0; row < rowLength; row++)
            {
                for (int col = 0; col < values[row].Length; col++)
                {
                    var valueObj = values[row][col];
                    if (valueObj is null) continue;
                    Table[(startRow + row, startCol + col)].Value = valueObj;
                }
            }
            Cursor.Col = startCol + colLength;

            var rangeStart = (startRow, startCol);
            var rangeEnd = (startRow + rowLength - 1, startCol + colLength - 1);

            RecalculateArea(rangeStart, rangeEnd);
            return this;
        }

        public UniversalTableBrush PrintLine() { Cursor.Row++; ResetCursorColumn(); return this; }
        public UniversalTableBrush PrintLine(IEnumerable<object> values) => PrintLine(PrintDirection.Horizontal, values.ToArray());
        public UniversalTableBrush PrintLine(object[] values) => PrintLine(PrintDirection.Horizontal, values);
        public UniversalTableBrush PrintLine(PrintDirection direction, IEnumerable<object> values) => PrintLine(direction, values.ToArray());
        public UniversalTableBrush PrintLine(PrintDirection direction, object[] values)
        {
            return Print(direction, values).Then(range =>
            {
                if (direction == PrintDirection.Horizontal) Cursor.Row++;
                else if (direction == PrintDirection.Vertical) Cursor.Row += values.Length;
                else throw new NotSupportedException();
                ResetCursorColumn();
            });
        }
        public UniversalTableBrush PrintLine(object[,] values)
        {
            var rowLength = values.GetLength(0);
            return Print(values).Then(range => { Cursor.Row += values.GetLength(0); ResetCursorColumn(); });
        }
        public UniversalTableBrush PrintLine(object[][] values)
        {
            var rowLength = values.GetLength(0);
            return Print(values).Then(range => { Cursor.Row += values.GetLength(0); ResetCursorColumn(); });
        }

        public UniversalTableBrush PrintDataTable(DataTable table)
        {
            var range1 = PrintLine(table.Columns.Cast<DataColumn>().Select(a => a.ColumnName).ToArray());
            var range2 = PrintLine((from DataRow row in table.Select() select row.ItemArray.ToArray()).ToArray());
            return this;
        }

    }
}
