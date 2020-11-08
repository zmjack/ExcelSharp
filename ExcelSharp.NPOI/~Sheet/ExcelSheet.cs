using ExcelSharp.NPOI.Utils;
using NPOI.SS.UserModel;
using NStandard;
using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Linq.Expressions;
using System.Reflection;
using System.Text.RegularExpressions;

namespace ExcelSharp.NPOI
{
    public partial class ExcelSheet
    {
        public const int EXCEL_WIDTH_PER_PX = 8;
        public const int AUTO_SIZE_PADDING_PX = 21;
        public const int COLUMN_BORDER_PX = 3;

        public ISheet MapedSheet { get; private set; }
        public ExcelBook Book { get; private set; }
        public Cursor Cursor;

        public ExcelSheet(ExcelBook excel, ISheet sheet)
        {
            if (sheet is null)
                throw new ArgumentException("Cannot find sheet.");

            Book = excel;
            MapedSheet = sheet;
            Cursor = (0, 0);
        }

        public void SetCursor(string cell) => Cursor = cell;

        public void CellMove() => Cursor.Col++;
        public void NextLine() => Cursor.Row++;

        public SheetCell this[Cursor cursor]
        {
            get
            {
                var irow = GetRow(cursor.Row);
                if (irow is null) irow = CreateRow(cursor.Row);

                var icol = irow.GetCell(cursor.Col);
                if (icol is null) icol = irow.CreateCell(cursor.Col);

                return new SheetCell(this, icol);
            }
        }
        public SheetRange this[Cursor start, Cursor end] => new SheetRange(this, start, end);

        public SheetRange Print(object[] values)
        {
            var startRow = Cursor.Row;
            for (int col = 0; col < values.Length; col++)
            {
                var valueObj = values[col];
                if (valueObj is null) continue;

                this[(Cursor.Row, Cursor.Col + col)].SetValue(valueObj);
            }
            return new SheetRange(this, (startRow, Cursor.Col), (startRow, Cursor.Col + values.Length - 1));
        }
        public SheetRange Print(object[,] values, bool reserveCursor = false)
        {
            var startRow = Cursor.Row;
            var rowLength = values.GetLength(0);
            var colLength = values.GetLength(1);

            for (var row = 0; row < rowLength; row++)
            {
                for (var col = 0; col < colLength; col++)
                {
                    var valueObj = values[row, col];
                    if (valueObj is null) continue;

                    if (valueObj is string && (valueObj as string).StartsWith("="))
                    {
                        //TODO: Analysis formula in same row
                        this[(Cursor.Row + row, Cursor.Col + col)].SetValue(valueObj);
                    }
                    else this[(Cursor.Row + row, Cursor.Col + col)].SetValue(valueObj);
                }
            }

            if (!reserveCursor) Cursor.Row += values.GetLength(0);

            return new SheetRange(this,
                (startRow, Cursor.Col),
                (startRow + rowLength - 1, Cursor.Col + colLength - 1));
        }
        public SheetRange Print(object[][] values)
        {
            var startRow = Cursor.Row;
            var rowLength = values.Length;
            var colLength = values.Any() ? values.Max(values1 => values1.Length) : 0;

            for (int row = 0; row < rowLength; row++)
            {
                for (int col = 0; col < values[row].Length; col++)
                {
                    var valueObj = values[row][col];
                    if (valueObj is null) continue;

                    this[(Cursor.Row + row, Cursor.Col + col)].SetValue(valueObj);
                }
            }

            Cursor.Row += values.Length;

            return new SheetRange(this,
                (startRow, Cursor.Col),
                (startRow + rowLength - 1, Cursor.Col + colLength - 1));
        }

        public SheetRange PrintLine(params object[] values)
        {
            return Print(values).Then(range => Cursor.Row++);
        }
        public SheetRange PrintDataTable(DataTable table)
        {
            var range1 = Print(table.Columns.Cast<DataColumn>().Select(a => a.ColumnName).ToArray());
            var range2 = Print((from DataRow row in table.Select()
                                select row.ItemArray.ToArray()).ToArray());
            return new SheetRange(this, range1.Start, range2.End);
        }
        public SheetRange PrintSnippet(ExcelSnippet snippet) => snippet.Print();

        public TModel[] Fetch<TModel>(Cursor startCell, Expression<Func<TModel, object>> includes = null, Predicate<int> rowSelector = null)
            where TModel : new()
        {
            (int row, int col) = (startCell.Row, startCell.Col);

            var ret = new List<TModel>();
            PropertyInfo[] props;

            var propNames = new string[0];
            if (includes != null)
            {
                switch (includes.Body)
                {
                    case MemberExpression exp:
                        propNames = new[] { exp.Member.Name };
                        break;

                    case NewExpression exp:
                        propNames = exp.Members.Select(x => x.Name).ToArray();
                        break;

                    default:
                        throw new NotSupportedException("This argument 'includes' must be MemberExpression or NewExpression.");
                }
            }

            if (propNames.Any())
            {
                props = typeof(TModel).GetProperties()
                    .Where(prop => prop.CanWrite && propNames.Contains(prop.Name))
                    .OrderBy(prop => propNames.IndexOf(prop.Name)).ToArray();
            }
            else props = typeof(TModel).GetProperties().Where(prop => prop.CanWrite).ToArray();

            for (int rowOffset = 0; row + rowOffset <= MapedSheet.LastRowNum; rowOffset++)
            {
                if (rowSelector != null)
                {
                    if (!rowSelector(rowOffset)) continue;
                }

                var @break = true;
                for (int i = 0; i < props.Length; i++)
                {
                    var cell = this[(row + rowOffset, col + i)];
                    var cellType = cell.IsMergedCell ? cell.MergedRange.Cell.CellType : cell.CellType;
                    if (cellType != CellType.Blank)
                    {
                        @break = false;
                        break;
                    }
                }
                if (@break) break;

                var item = new TModel();
                foreach (var kv in props.AsKvPairs())
                {
                    var propInfo = kv.Value;
                    var cell = this[(row + rowOffset, col + kv.Key)];

                    if (propInfo.PropertyType == typeof(DateTime) || propInfo.PropertyType == typeof(DateTime?))
                    {
                        if (cell.CellType == CellType.Blank)
                        {
                            if (propInfo.PropertyType.IsNullable()) propInfo.SetValue(item, ConvertEx.ChangeType(null, propInfo.PropertyType));
                            else propInfo.SetValue(item, ConvertEx.ChangeType(default(DateTime), propInfo.PropertyType));
                        }
                        else
                        {
                            var value = cell.IsMergedCell ? cell.MergedRange.Cell.DateTime : cell.DateTime;
                            propInfo.SetValue(item, ConvertEx.ChangeType(value, propInfo.PropertyType));
                        }
                    }
                    else
                    {
                        var value = cell.IsMergedCell ? cell.MergedRange.Cell.GetValue() : cell.GetValue();
                        propInfo.SetValue(item, ConvertUtil.Convert(value, propInfo.PropertyType));
                    }
                }
                ret.Add(item);
            }

            return ret.ToArray();
        }

        public DataTable Fetch(Cursor startCell, bool isFirstRowTitle)
        {
            (int row, int col) = (startCell.Row, startCell.Col);

            var colLength = 0;
            var @continue = true;
            for (int colOffset = 0; @continue; colOffset++)
            {
                if (this[(row, col + colOffset)].For(x => x.MapedCell.CellType != CellType.Blank && !x.String.IsNullOrWhiteSpace()))
                {
                    colLength++;
                }
                else break;
            }

            return Fetch(startCell, isFirstRowTitle, new Type[colLength].Let(i => typeof(string)).ToArray());
        }
        public DataTable Fetch(string startCell, bool isFirstRowTitle, Type[] colTypes)
        {
            return Fetch(startCell, isFirstRowTitle, colTypes);
        }
        public DataTable Fetch(Cursor startCell, bool isFirstRowTitle, Type[] colTypes)
        {
            (int row, int col) = (startCell.Row, startCell.Col);

            var ret = new DataTable();
            ret.Columns.Then(x => x.AddRange(colTypes.Select(colType => new DataColumn("", colType)).ToArray()));

            (int row, int col) dataStart;
            if (isFirstRowTitle)
            {
                for (int i = 0; i < colTypes.Length; i++)
                    ret.Columns[i].ColumnName = this[(row, col + i)].String;
                dataStart = (row + 1, col);
            }
            else dataStart = (row, col);

            var @continue = true;
            for (int rowOffset = 0; @continue; rowOffset++)
            {
                var allEmpty = true;
                for (int i = 0; i < colTypes.Length; i++)
                {
                    if (this[(row + rowOffset, col + i)].MapedCell.CellType != CellType.Blank)
                    {
                        allEmpty = false;
                        break;
                    }
                }

                if (!allEmpty)
                {
                    var rowValues = colTypes.Select((colType, i) =>
                    {
                        var cell = this[(dataStart.row + rowOffset, col + i)];

                        if (colType == typeof(DateTime))
                            return cell.DateTime;
                        else return cell.GetValue();
                    }).ToArray();
                    ret.Rows.Add(rowValues);
                }
                else @continue = false;
            }

            return ret;
        }

        public void SetWidth(string colName, double width) => SetWidth(GetCol(colName), width);
        public void SetWidth(int columnIndex, double width) => MapedSheet.SetColumnWidth(columnIndex, (int)Math.Ceiling(255.86 * width + 184.27));

        public void SetHeight(string rowName, float height) => this[(GetRow(rowName), 1)].SetHeight(height);
        public void SetHeight(int rowIndex, float height) => this[(rowIndex, 1)].SetHeight(height);

        public int GetWidth(int columnIndex) => (int)Math.Ceiling((MapedSheet.GetColumnWidth(columnIndex) - 184.27) / 255.86);

        public void AutoSize(params string[] columns) => AutoSize(columns.Select(x => LetterSequence.GetNumber(x)).ToArray());
        public void AutoSize(params int[] columns)
        {
            foreach (var col in columns)
            {
                int defaultWidth = GetWidth(col);
                var exceptRows = EnumerableEx.Concat(
                    MergedRanges
                        .Where(x => x.Start.Col < col && col <= x.End.Col)
                        .SelectMany(x => new int[x.End.Row - x.Start.Row + 1].Let(i => x.Start.Row + i))
                        .ToArray(),
                    MergedRanges
                        .Where(x => x.Start.Col == col)
                        .SelectMany(x => new int[x.End.Row - x.Start.Row].Let(i => x.Start.Row + i + 1))
                        .ToArray())
                    .ToArray();

                var rowNumbers = new int[LastRowNum + 1].Let(i => i).Where(row => !exceptRows.Contains(row));
                var widths = rowNumbers.Select(row =>
                {
                    var cell = this[(row, col)];
                    if (cell.MergedRange?.For(_ => !_.IsSingleColumn || !_.IsDefinitionCell(cell)) ?? false) return 0;

                    var cstyle = cell.GetCStyle();
                    var value = cell.GetValue();

                    string valueString;
                    if (value is double)
                        valueString = ((double)value).ToString(cstyle.DataFormat);
                    else if (value is DateTime)
                        valueString = ((DateTime)value).ToString(cstyle.DataFormat);
                    else valueString = value?.ToString() ?? "";

                    using (var bitmap = new Bitmap(1, 1))
                    using (var graphics = Graphics.FromImage(new Bitmap(1, 1)))
                    {
                        var fontSize = graphics.MeasureString(valueString, new Font(cstyle.Font.FontName, cstyle.Font.FontSize));
                        var width = fontSize.Width > 0 ? (int)((COLUMN_BORDER_PX + AUTO_SIZE_PADDING_PX + fontSize.Width) / EXCEL_WIDTH_PER_PX) : 0;
                        return width;
                    }
                });

                SetWidth(col, new[] { defaultWidth }.Concat(widths).Max());
            }
        }

        public IEnumerable<SheetRange> MergedRanges
        {
            get
            {

                return new SheetRange[NumMergedRegions].Let(i => GetMergedRegion(i)
                    .For(x => new SheetRange(this, (x.FirstRow, x.FirstColumn), (x.LastRow, x.LastColumn))));
            }
        }

    }
}