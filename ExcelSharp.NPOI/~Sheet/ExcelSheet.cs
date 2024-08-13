using ExcelSharp.NPOI.Utils;
using NPOI.SS.Formula.Functions;
using NPOI.SS.UserModel;
using NStandard;
using SkiaSharp;
using System;
using System.Collections.Generic;
using System.Collections.Immutable;
using System.Data;
using System.Linq;
using System.Linq.Expressions;
using System.Reflection;
using ISheet = NPOI.SS.UserModel.ISheet;

namespace ExcelSharp.NPOI;

public partial class ExcelSheet
{
    public const int EXCEL_WIDTH_PER_PX = 8;
    public const int AUTO_SIZE_PADDING_PX = 21;
    public const int COLUMN_BORDER_PX = 3;

    public ISheet MapedSheet { get; private set; }
    public ExcelBook Book { get; private set; }
    public Cursor Cursor;

    internal ExcelSheet(ExcelBook excel, ISheet sheet)
    {
        if (sheet is null) throw new ArgumentException("Cannot find sheet.");
        Book = excel;
        MapedSheet = sheet;
        Cursor = (0, 0);
    }

    public ExcelArea BeginArea() => new(this);
    public ExcelArea BeginArea(Cursor cursor)
    {
        Cursor = cursor;
        return new ExcelArea(this);
    }
    public ExcelArea BeginDefaultArea()
    {
        int lastCol = 0;
        for (int row = 0; row < MapedSheet.LastRowNum; row++)
        {
            var rowLastCol = MapedSheet.GetRow(row)?.LastCellNum ?? 0;
            if (lastCol < rowLastCol) lastCol = rowLastCol;
        }
        return new ExcelArea(this, (0, 0), (MapedSheet.LastRowNum, lastCol));
    }

    public void SetCursor(string cell) => Cursor = cell;

    public void CellMove() => Cursor.Col++;
    public void NextLine() => Cursor.Row++;

    public SheetCell this[Cursor cursor]
    {
        get
        {
            var irow = GetRow(cursor.Row) ?? CreateRow(cursor.Row);
            var icol = irow.GetCell(cursor.Col) ?? irow.CreateCell(cursor.Col);
            return new SheetCell(this, icol);
        }
    }
    public SheetRange this[Cursor start, Cursor end] => new(this, start, end);

    public void ResetCursorColumn()
    {
        var recent = ExcelArea.Scopes.FirstOrDefault(x => x.Sheet == this);
        if (recent is not null) Cursor.Col = ExcelArea.Current.Start.Col;
        else Cursor.Col = 0;
    }
    private void RecalculateArea(Cursor start, Cursor end)
    {
        var recent = ExcelArea.Scopes.FirstOrDefault(x => x.Sheet == this);
        recent?.Update(start, end);
    }

    public void PrintSpreadSheet(ISpreadSheet sheet, Cursor cursor)
    {
        Cursor = cursor;
        PrintSpreadSheet(sheet);
    }

    public void PrintSpreadSheet(ISpreadSheet sheet)
    {
        var widthDict = new Dictionary<int, float>();

        var cellsGroups = sheet.Rows.SelectMany(x => x.Cells).GroupBy(x => new { x.Style, x.Style?.Format });
        foreach (var cells in cellsGroups)
        {
            var style = cells.Key.Style;
            var fontSize = (short)Math.Round((style?.FontSize ?? 14) * 0.75, 0, MidpointRounding.AwayFromZero);

            var fontName = style?.FontFamily;
            using SkiaUtil? skia = fontName is not null ? new(fontName, fontSize) : null;

            CStyle? cstyle = null;
            if (style is not null)
            {
                cstyle = Book.CStyle(x =>
                {
                    if (style.BackgroundColor is not null) x.CellColor(style.BackgroundColor.Value);
                    if (style.Color is not null || style.FontSize.HasValue || style.FontFamily is not null)
                    {
                        var fontFamily = style.FontFamily;
                        if (style.Color is not null)
                        {
                            var color = style.Color;
                            x.SetFont(fontFamily, fontSize, color.Value);
                        }
                        else x.SetFont(fontFamily, fontSize);
                    }
                    if (style.Bold.HasValue && style.Bold.Value) x.Bold();
                    if (style.BorderTop.HasValue && style.BorderTop.Value) x.BorderTop = BorderStyle.Thin;
                    if (style.BorderBottom.HasValue && style.BorderBottom.Value) x.BorderBottom = BorderStyle.Thin;
                    if (style.BorderLeft.HasValue && style.BorderLeft.Value) x.BorderLeft = BorderStyle.Thin;
                    if (style.BorderRight.HasValue && style.BorderRight.Value) x.BorderRight = BorderStyle.Thin;
                    if (style.HoriAlign != HoriAlign.Preserve)
                    {
                        switch (style.HoriAlign)
                        {
                            case HoriAlign.Left: x.HLeft(); break;
                            case HoriAlign.Center: x.HCenter(); break;
                            case HoriAlign.Right: x.HRight(); break;
                        }
                    }
                    if (style.VertAlign != VertAlign.Preserve)
                    {
                        switch (style.VertAlign)
                        {
                            case VertAlign.Top: x.VTop(); break;
                            case VertAlign.Middle: x.VCenter(); break;
                            case VertAlign.Bottom: x.VBottom(); break;
                        }
                    }

                    x.DataFormat = cells.Key.Format ?? "General";
                    x.WordWrap();
                });
            }

            foreach (var cell in cells.Where(x => !x.Ignored))
            {
                var ecell = this[(Cursor.Row + cell.RowIndex, Cursor.Col + cell.ColIndex)];

                if (!cell.Ignored)
                {
                    ecell.SetValue(cell.Value);

                    if (skia is not null)
                    {
                        var width = skia.GetPaintWidth(cell?.Value?.ToString());
                        if (width > widthDict.GetValueOrDefault(cell!.ColIndex))
                        {
                            widthDict[cell!.ColIndex] = width;
                        }
                    }

                    if (cell.RowSpan > 1 || cell.ColSpan > 1)
                    {
                        var endRow = ecell.RowIndex + cell.RowSpan - 1;
                        var endCol = ecell.ColumnIndex + cell.ColSpan - 1;
                        var range = new SheetRange(this, (ecell.RowIndex, ecell.ColumnIndex), (endRow, endCol));

                        if (cstyle is not null)
                        {
                            for (int row = ecell.RowIndex; row <= endRow; row++)
                            {
                                for (int col = ecell.ColumnIndex; col <= endCol; col++)
                                {
                                    var _ecell = this[(row, col)];
                                    _ecell.SetCStyle(cstyle);
                                }
                            }
                        }

                        range.Merge();
                    }
                    else
                    {
                        if (cstyle is not null)
                        {
                            ecell.SetCStyle(cstyle);
                        }
                    }
                }
            }
        }

        foreach (var pair in widthDict)
        {
            var col = pair.Key;
            var width = pair.Value;
            var targetWidth = (int)((COLUMN_BORDER_PX + AUTO_SIZE_PADDING_PX + width) / EXCEL_WIDTH_PER_PX);
            SetWidth(col, targetWidth);
        }

        Cursor.Row += sheet.RowLength;
    }

    public SheetRange Print(IEnumerable<object> values) => Print(PrintDirection.Horizontal, values.ToArray());
    public SheetRange Print(object[] values) => Print(PrintDirection.Horizontal, values);
    public SheetRange Print(PrintDirection direction, IEnumerable<object> values) => Print(direction, values.ToArray());
    public SheetRange Print(PrintDirection direction, object[] values)
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
                this[(startRow, startCol + col)].SetValue(valueObj);
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
                this[(startRow + row, startCol)].SetValue(valueObj);
            }
            Cursor.Col = startCol + 1;
            rangeEnd = new Cursor { Row = startRow + length - 1, Col = startCol };
        }
        else throw new NotSupportedException();

        RecalculateArea(rangeStart, rangeEnd);

        return new SheetRange(this, rangeStart, rangeEnd);
    }
    public SheetRange Print(object[,] values)
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

                if ((valueObj as string)?.StartsWith("=") ?? false)
                {
                    //TODO: Analysis formula in same row
                    this[(startRow + row, startCol + col)].SetValue(valueObj);
                }
                else this[(startRow + row, startCol + col)].SetValue(valueObj);
            }
        }
        Cursor.Col = startCol + colLength;

        var rangeStart = (startRow, startCol);
        var rangeEnd = (startRow + rowLength - 1, startCol + colLength - 1);
        RecalculateArea(rangeStart, rangeEnd);

        return new SheetRange(this, rangeStart, rangeEnd);
    }
    public SheetRange Print(object[][] values)
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
                this[(startRow + row, startCol + col)].SetValue(valueObj);
            }
        }
        Cursor.Col = startCol + colLength;

        var rangeStart = (startRow, startCol);
        var rangeEnd = (startRow + rowLength - 1, startCol + colLength - 1);
        RecalculateArea(rangeStart, rangeEnd);

        return new SheetRange(this, rangeStart, rangeEnd);
    }

    public void PrintLine() { Cursor.Row++; ResetCursorColumn(); }
    public SheetRange PrintLine(IEnumerable<object> values) => PrintLine(PrintDirection.Horizontal, values.ToArray());
    public SheetRange PrintLine(object[] values) => PrintLine(PrintDirection.Horizontal, values);
    public SheetRange PrintLine(PrintDirection direction, IEnumerable<object> values) => PrintLine(direction, values.ToArray());
    public SheetRange PrintLine(PrintDirection direction, object[] values)
    {
        var range = Print(direction, values);

        if (direction == PrintDirection.Horizontal) Cursor.Row++;
        else if (direction == PrintDirection.Vertical) Cursor.Row += values.Length;
        else throw new NotSupportedException();
        ResetCursorColumn();

        return range;
    }
    public SheetRange PrintLine(object[,] values)
    {
        var range = Print(values);
        Cursor.Row += values.GetLength(0);
        ResetCursorColumn();
        return range;
    }
    public SheetRange PrintLine(object[][] values)
    {
        var range = Print(values);
        Cursor.Row += values.GetLength(0);
        ResetCursorColumn();
        return range;
    }

    public SheetRange PrintDataTable(DataTable table)
    {
        var range1 = PrintLine(table.Columns.Cast<DataColumn>().Select(a => a.ColumnName).ToArray());
        var range2 = PrintLine((from DataRow row in table.Select() select row.ItemArray.ToArray()).ToArray());
        return new SheetRange(this, range1.Start, range2.End);
    }

    private static object? GetCellValue(SheetCell cell, Type type)
    {
        if (cell.IsMergedCell && cell.MergedRange!.Cell.CellName != cell.CellName) return GetCellValue(cell.MergedRange.Cell, type);

        try
        {
            if (cell.CellType == CellType.Blank)
            {
                if (type.IsNullable()) return ConvertEx.ChangeType(null, type);
                else return type.CreateDefault();
            }
            else if (type == typeof(DateTime) || type == typeof(DateTime?))
            {
                var value = cell.DateTime;
                return ConvertEx.ChangeType(value, type);
            }
#if NET6_0_OR_GREATER
            else if (type == typeof(DateOnly) || type == typeof(DateOnly?))
            {
                var value = cell.DateTime;
                return DateOnly.FromDateTime((DateTime)ConvertEx.ChangeType(value, typeof(DateTime)));
            }
#endif
            else
            {
                var value = cell.GetValue();
                return ConvertUtil.Convert(value, type);
            }
        }
        catch
        {
            return type.CreateDefault();
        }
    }
    public TModel[] Fetch<TModel, TImplement>(Cursor startCell) where TModel : TImplement, new()
    {
        (int row, int col) = (startCell.Row, startCell.Col);
        var props = (
             from prop in typeof(TImplement).GetProperties()
             let names = prop.GetCustomAttribute<DataColumnAttribute>()?.Names ?? []
             where prop.CanWrite
             select KeyValuePair.Create(prop.Name, names)
        ).ToArray();

        var propNames = new List<string>();
        for (int i = 0; i < 200; i++)
        {
            var cell = this[(row, col + i)];
            var value = GetCellValue(cell, typeof(string));
            var svalue = value?.ToString().Unique() ?? string.Empty;
            if (!svalue.IsNullOrWhiteSpace())
            {
                var prop = props.FirstOrDefault(x => x.Value.Contains(svalue));
                if (prop.Key is not null)
                {
                    propNames.Add(prop.Key);
                }
            }
        }

        return Fetch<TModel>((startCell.Row + 1, startCell.Col), [.. propNames]);
    }

    public TModel[] Fetch<TModel>(Cursor startCell) where TModel : new()
    {
        (int row, int col) = (startCell.Row, startCell.Col);
        var props = (
             from prop in typeof(TModel).GetProperties()
             let names = prop.GetCustomAttribute<DataColumnAttribute>()?.Names ?? []
             where prop.CanWrite
             select KeyValuePair.Create(prop.Name, names)
        ).ToArray();

        var propNames = new List<string>();
        for (int i = 0; i < 200; i++)
        {
            var cell = this[(row, col + i)];
            var value = GetCellValue(cell, typeof(string));
            var svalue = value?.ToString().Unique() ?? string.Empty;
            if (!svalue.IsNullOrWhiteSpace())
            {
                var prop = props.FirstOrDefault(x => x.Value.Contains(svalue));
                if (prop.Key is not null)
                {
                    propNames.Add(prop.Key);
                }
            }
        }

        return Fetch<TModel>((startCell.Row + 1, startCell.Col), [.. propNames]);
    }

    public IReadOnlyCollection<TModel> Fetch2D<TModel, TValue>(Cursor archer)
        where TModel : IFetchModel<TValue>, new()
    {
        var cell = this[archer];
        return Fetch2D<TModel, TValue>(archer, (archer.Row + cell.RowSpan - 1, archer.Col + cell.ColSpan - 1));
    }

    public IReadOnlyCollection<TModel> Fetch2D<TModel, TValue>(Cursor archerStart, Cursor archerEnd)
        where TModel : IFetchModel<TValue>, new()
    {
        var archerRowLength = archerEnd.Row - archerStart.Row + 1;
        var archerColLength = archerEnd.Col - archerStart.Col + 1;

        var colList = new List<string?[]>();
        for (int col = archerColLength; col < 200; col++)
        {
            var values = new List<string?>();
            var valid = false;

            for (int i = 0; i < archerRowLength; i++)
            {
                var cell = this[(archerStart.Row + i, archerStart.Col + col)];
                var value = GetCellValue(cell, typeof(string));
                var svalue = value?.ToString().Unique() ?? string.Empty;
                if (!svalue.IsNullOrWhiteSpace())
                {
                    valid = true;
                    values.Add(svalue);
                }
                else values.Add(null);
            }

            if (valid)
            {
                colList.Add([.. values]);
            }
            else break;
        }

        var rowList = new List<string?[]>();
        for (int row = archerRowLength; row < 20000; row++)
        {
            var values = new List<string?>();
            var valid = false;

            for (int i = 0; i < archerColLength; i++)
            {
                var cell = this[(archerStart.Row + row, archerStart.Col + i)];
                var value = GetCellValue(cell, typeof(string));
                var svalue = value?.ToString().Unique() ?? string.Empty;

                if (!svalue.IsNullOrWhiteSpace())
                {
                    valid = true;
                    values.Add(svalue);
                }
                else values.Add(null);
            }

            if (valid)
            {
                rowList.Add([.. values]);
            }
            else break;
        }

        var models = new List<TModel>();
        var start = new Cursor(archerStart.Row + archerRowLength, archerStart.Col + archerColLength);
        var type = typeof(TValue);

        var provider = new HeaderProvider<TModel, TValue>();

        foreach (var (rowIndex, rowNames) in rowList.Pairs())
        {
            foreach (var (colIndex, colNames) in colList.Pairs())
            {
                var model = provider.GetModel(rowNames, colNames)!;

                var cell = this[(start.Row + rowIndex, start.Col + colIndex)];
                var cellType = cell.IsMergedCell ? cell.MergedRange!.Cell.CellType : cell.CellType;

                if (cellType != CellType.Blank)
                {
                    var value = (TValue?)GetCellValue(cell, type);
                    model.Value = value;
                }

                models.Add(model);
            }
        }

        return models;
    }

    public TModel[] Fetch<TModel>(Cursor startCell, Expression<Func<TModel, object>> includes, Predicate<int>? rowSelector = null)
        where TModel : new()
    {
        var propNames = includes is not null ? includes.Body switch
        {
            MemberExpression exp => [exp.Member.Name],
            NewExpression exp => exp.Members!.Select(x => x.Name).ToArray(),
            _ => throw new ArgumentException("Any element must be MemberExpression or NewExpression.", nameof(includes)),
        } : [];

        return Fetch<TModel>(startCell, propNames, rowSelector);
    }

    public TModel[] Fetch<TModel, TImplement>(Cursor startCell, Expression<Func<TImplement, object>> includes, Predicate<int>? rowSelector = null)
        where TModel : TImplement, new()
    {
        var propNames = includes is not null ? includes.Body switch
        {
            MemberExpression exp => [exp.Member.Name],
            NewExpression exp => exp.Members!.Select(x => x.Name).ToArray(),
            _ => throw new ArgumentException("Any element must be MemberExpression or NewExpression.", nameof(includes)),
        } : [];

        return Fetch<TModel, TImplement>(startCell, propNames, rowSelector);
    }

    public TModel[] Fetch<TModel>(Cursor startCell, string[] propNames, Predicate<int>? rowSelector = null)
        where TModel : new()
    {
        var props = propNames.Any() ?
        [
            .. from prop in typeof(TModel).GetProperties()
               where prop.CanWrite && propNames.Contains(prop.Name)
               orderby propNames.IndexOf(prop.Name)
               select prop
        ] : typeof(TModel).GetProperties().Where(prop => prop.CanWrite).ToArray();

        return Fetch<TModel, TModel>(startCell, props, rowSelector);
    }

    public TModel[] Fetch<TModel, TImplement>(Cursor startCell, string[] propNames, Predicate<int>? rowSelector = null)
        where TModel : TImplement, new()
    {
        var props = propNames.Any() ?
        [
            .. from prop in typeof(TImplement).GetProperties()
               where prop.CanWrite && propNames.Contains(prop.Name)
               orderby propNames.IndexOf(prop.Name)
               select prop
        ] : typeof(TModel).GetProperties().Where(prop => prop.CanWrite).ToArray();

        return Fetch<TModel, TImplement>(startCell, props, rowSelector);
    }

    private TModel[] Fetch<TModel, TImplement>(Cursor startCell, PropertyInfo[] props, Predicate<int>? rowSelector = null)
        where TModel : new()
    {
        (int row, int col) = (startCell.Row, startCell.Col);

        var ret = new List<TModel>();
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
                var cellType = cell.IsMergedCell ? cell.MergedRange!.Cell.CellType : cell.CellType;
                if (cellType != CellType.Blank)
                {
                    @break = false;
                    break;
                }
            }
            if (@break) break;

            var model = new TModel();
            var colOffset = 0;
            foreach (var property in props)
            {
                var propertyType = property.PropertyType;

                if (propertyType.IsArray)
                {
                    var fetchAttribute = property.GetCustomAttribute<FetchAttribute>();
                    if (fetchAttribute is null) continue;

                    var fetchLength = fetchAttribute.Length;
                    var elementType = propertyType.GetElementType()!;
                    var array = Array.CreateInstance(elementType, fetchLength);

                    for (var i = 0; i < fetchLength; i++)
                    {
                        var cell = this[(row + rowOffset, col + colOffset)];
                        var cellValue = GetCellValue(cell, elementType);

                        var value = property.GetValue(model);
                        array.SetValue(cellValue, i);

                        colOffset++;
                    }
                    property.SetValue(model, array);
                }
                else
                {
                    var cell = this[(row + rowOffset, col + colOffset)];
                    var cellValue = GetCellValue(cell, property.PropertyType);
                    property.SetValue(model, cellValue);
                    colOffset++;
                }
            }
            ret.Add(model);
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
            if (this[(row, col + colOffset)].Pipe(x => x.MapedCell.CellType != CellType.Blank && !x.String.IsNullOrWhiteSpace()))
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
        ret.Columns.AddRange(colTypes.Select(colType => new DataColumn("", colType)).ToArray());

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

                    return colType switch
                    {
                        Type type when type == typeof(DateTime) => cell.DateTime,
#if NET6_0_OR_GREATER
                        Type type when type == typeof(DateOnly) => cell.DateTime,
#endif
                        _ => cell.GetValue(),
                    };
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

    public void AutoSize(string[] columns, int maxWidth = 64) => AutoSize(columns.Select(ColumnIndex.GetIndex).ToArray(), maxWidth);
    public void AutoSize(int[] columns, int maxWidth = 64)
    {
        foreach (var col in columns)
        {
            int excelWidth = GetWidth(col);
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
            var stringSet = new HashSet<string>();

            foreach (var row in rowNumbers)
            {
                var cell = this[(row, col)];
                var range = cell.MergedRange;
                if (range is not null && (!range.IsSingleColumn || !range.IsDefinitionCell(cell))) continue;

                var cstyle = cell.GetCStyle();
                var value = cell.GetValue();

                string valueString;
                if (value is double d) valueString = d.ToString(cstyle.DataFormat);
                else if (value is DateTime dt) valueString = dt.ToString(cstyle.DataFormat);
                else valueString = value?.ToString() ?? "";

                if (stringSet.Contains(valueString)) continue;

                stringSet.Add(valueString);

                using var skia = new SkiaUtil(cstyle.Font.FontName!, cstyle.Font.FontSize);

                float width = skia.GetPaintWidth(valueString);
                var targetWidth = width > 0 ? (int)((COLUMN_BORDER_PX + AUTO_SIZE_PADDING_PX + width) / EXCEL_WIDTH_PER_PX) : 0;

                if (targetWidth >= maxWidth)
                {
                    excelWidth = maxWidth;
                    break;
                }

                if (targetWidth > excelWidth) excelWidth = targetWidth;
            }

            SetWidth(col, excelWidth);
        }
    }

    public IEnumerable<SheetRange> MergedRanges
    {
        get
        {
            return new SheetRange[NumMergedRegions].Let(i => GetMergedRegion(i)
                .Pipe(x => new SheetRange(this, (x.FirstRow, x.FirstColumn), (x.LastRow, x.LastColumn))));
        }
    }

}
