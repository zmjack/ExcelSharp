﻿using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NStandard;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace ExcelSharp.NPOI;

public partial class SheetRange : IEnumerable<SheetCell>, IStylizable
{
    public ExcelSheet Sheet { get; private set; }
    public Cursor Start { get; private set; }
    public Cursor End { get; private set; }
    public string Name => $"{Start}:{End}";
    public int RowLength => End.Row - Start.Row + 1;
    public int ColumnLengh => End.Col - Start.Col + 1;
    public SheetCell Cell => Sheet[Start];

    public bool IsSingleRow => RowLength == 1;
    public bool IsSingleColumn => ColumnLengh == 1;

    internal SheetRange(ExcelSheet sheet, Cursor start, Cursor end)
    {
        Sheet = sheet;
        Start = start;
        End = end;
    }

    public bool IsInRange(SheetCell cell)
    {
        return Start.Row <= cell.RowIndex && cell.RowIndex <= End.Row
            && Start.Col <= cell.ColumnIndex && cell.ColumnIndex <= End.Col;
    }
    public bool IsDefinitionCell(SheetCell cell) => cell.RowIndex == Start.Row && cell.ColumnIndex == Start.Col;

    public void SetCellStyle(ICellStyle style)
    {
        for (int row = Start.Row; row <= End.Row; row++)
        {
            for (int col = Start.Col; col <= End.Col; col++)
            {
                Sheet[(row, col)].SetCellStyle(style);
            }
        }
    }
    public void SetCStyle(CStyle style) => SetCellStyle(style.CellStyle);
    public void SetCStyle(Action<CStyleApplier> initApplier) => SetCellStyle(Sheet.Book.CStyle(initApplier).CellStyle);
    public void UpdateCStyle(Action<CStyleApplier> initApplier)
    {
        for (int row = Start.Row; row <= End.Row; row++)
        {
            for (int col = Start.Col; col <= End.Col; col++)
            {
                var applier = Sheet[(row, col)].GetCStyle().GetApplier().Pipe(x => initApplier(x));
                Sheet[(row, col)].SetCellStyle(Sheet.Book.CStyle(applier).CellStyle);
            }
        }
    }

    public SheetRangeColumnSelector Columns => new(this);
    public SheetRangeGroup Column(params int[] indexes)
    {
        IEnumerable<SheetRange> select()
        {
            foreach (var index in indexes)
                yield return new SheetRange(Sheet, (Start.Row, Start.Col + index), (End.Row, Start.Col + index));
        }
        return new SheetRangeGroup(select());
    }

    public SheetRangeRowSelector Rows => new(this);
    public SheetRangeGroup Row(params int[] indexes)
    {
        IEnumerable<SheetRange> select()
        {
            foreach (var index in indexes)
                yield return new SheetRange(Sheet, (Start.Row + index, Start.Col), (Start.Row + index, End.Col));
        }
        return new SheetRangeGroup(select());
    }

    public void Merge()
    {
        for (int row = Start.Row; row <= End.Row; row++)
        {
            for (int col = Start.Col; col <= End.Col; col++)
            {
                if (!(row == Start.Row && col == Start.Col))
                {
                    Sheet[(row, col)].SetValue("");
                }
            }
        }

        var region = new CellRangeAddress(Start.Row, End.Row, Start.Col, End.Col);
        Sheet.AddMergedRegion(region);
    }

    [GeneratedRegex("^\\[\\[.+?\\]\\](.*)$")]
    private static partial Regex SmartMergeRegex();

    /// <summary>
    /// Merge cells that has same value.
    ///     You can use [[ ]] to identifier cells, the result of merged cell will ignore the value which is in [[ ]].
    /// </summary>
    /// <param name="offsetCols"></param>
    public void SmartMerge(params int[] offsetCols)
    {
        var col = Start.Col + offsetCols[0];

        string? take = null;
        int mergeStart = Start.Row;

        for (int takeRow = mergeStart; takeRow <= End.Row; takeRow++)
        {
            var value = Sheet[(takeRow, col)].GetValue()?.ToString();
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
        var regex_matchId = SmartMergeRegex();
        foreach (var colIndex in offsetCols)
        {
            foreach (var cell in Column(colIndex))
            {
                var value = cell.GetValue();
                if (value is string && !(value as string).IsNullOrWhiteSpace())
                {
                    var match = regex_matchId.Match((value as string)!);
                    if (match.Success)
                    {
                        if (double.TryParse(match.Groups[1].Value, out double dvalue))
                            cell.SetValue(dvalue);
                        else cell.SetValue(match.Groups[1].Value);
                    }
                }
            }
        }
    }

    public void SmartColMerge()
    {
        void InnerMerge(int row, int startCol, int endCol, string? value)
        {
            value ??= "";

            if (endCol - startCol > 0)
                new SheetRange(Sheet, (row, startCol), (row, endCol)).Merge();

            var regex = SmartMergeRegex();
            var match = regex.Match(value);

            if (match.Success)
            {
                var cell = Sheet[(row, startCol)];
                var svalue = match.Groups[1].Value;
                if (!svalue.IsNullOrWhiteSpace())
                {
                    if (double.TryParse(svalue, out double dvalue))
                        cell.SetValue(dvalue);
                    else cell.SetValue(svalue);
                }
                else cell.SetValue(svalue);
            }
        }

        for (int row = Start.Row; row <= End.Row; row++)
        {
            string? take = null;
            int startCol = Start.Col;
            int takeCol = startCol;

            for (; takeCol <= End.Col; takeCol++)
            {
                var value = Sheet[(row, takeCol)].GetValue()?.ToString();
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
        new SheetRange(Sheet, (mergeStart, col), (mergeEnd, col)).Merge();
        if (offsetCols.Length > 1)
        {
            new SheetRange(Sheet, (mergeStart, col + offsetCols[1]), (mergeEnd, End.Col))
                .SmartMerge(offsetCols.Slice(1).Select(_col => _col - (offsetCols[1] - offsetCols[0])).ToArray());
        }
    }

    public SheetCell this[(int offsetRow, int offsetCol) pos] => Sheet[(Start.Row + pos.offsetRow, Start.Col + pos.offsetCol)];

    public IEnumerable<SheetRange> GetRows()
    {
        for (int row = Start.Row; row <= End.Row; row++)
            yield return new SheetRange(Sheet, (row, Start.Col), (row, End.Col));
    }

    public IEnumerator<SheetCell> GetEnumerator()
    {
        for (int row = Start.Row; row <= End.Row; row++)
            for (int col = Start.Col; col <= End.Col; col++)
                yield return new SheetCell(Sheet, Sheet[(row, col)]);
    }
    IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();

}
