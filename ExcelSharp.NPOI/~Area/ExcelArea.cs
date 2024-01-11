using NPOI.SS.UserModel;
using NStandard;
using System.Linq;
using VerticalAlignment = NPOI.SS.UserModel.VerticalAlignment;

namespace ExcelSharp.NPOI;

public class ExcelArea : Scope<ExcelArea>
{
    public ExcelSheet Sheet { get; private set; }

    internal Cursor _start;
    internal Cursor _end;
    public Cursor Start { get => _start; }
    public Cursor End { get => _end; }

    public ExcelArea(ExcelSheet sheet)
    {
        Sheet = sheet;
        _start = _end = sheet.Cursor;
    }
    public ExcelArea(ExcelSheet sheet, Cursor start, Cursor end)
    {
        Sheet = sheet;
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

    public SheetRange GetRange() => new(Sheet, _start, _end);

    public HtmlTable ToHtmlTable()
    {
        //TODO: sheet name
        var uniTable = new MungSheet("Sheet1");
        for (int row = _start.Row; row <= _end.Row; row++)
        {
            var rowOffset = row - _start.Row;
            var uniRow = uniTable.Rows[rowOffset];

            for (int col = _start.Col; col <= _end.Col; col++)
            {
                var colOffset = col - _start.Col;
                var uniCell = uniRow[colOffset];

                var cell = Sheet[(row, col)];
                var cstyle = cell.GetCStyle();

                if (!cell.IsMergedCell || cell.IsMergedDefinitionCell)
                {
                    var value = cell.GetValue();

                    uniCell.Value = cell.GetValue()?.Pipe(value =>
                    {
                        var format = cstyle.DataFormat;
                        if (cell.CellType == CellType.Numeric && format != "General") return ((double)value).ToString(format);
                        else return value?.ToString();
                    });

                    uniCell.RowSpan = cell.RowSpan;
                    uniCell.ColSpan = cell.ColSpan;

                    if (uniCell.Value is not null)
                    {
                        uniCell.Comment = cell.CellComment?.String.String;
                        uniCell.Style = new MungStyle
                        {
                            BackgroundColor = cstyle.FillPattern != FillPattern.NoFill ? cstyle.FillForegroundColor : null,
                            Color = cstyle.Font.FontColor,
                            Bold = cstyle.Font.IsBold,
                            FontSize = cstyle.Font.FontSize,
                            BorderTop = cstyle.BorderTop != BorderStyle.None,
                            BorderBottom = cstyle.BorderBottom != BorderStyle.None,
                            BorderLeft = cstyle.BorderLeft != BorderStyle.None,
                            BorderRight = cstyle.BorderRight != BorderStyle.None,
                            HoriAlign = cstyle.Alignment switch
                            {
                                HorizontalAlignment.Left => HoriAlign.Left,
                                HorizontalAlignment.Center => HoriAlign.Center,
                                HorizontalAlignment.Right => HoriAlign.Right,
                                _ => HoriAlign.Preserve,
                            },
                            VertAlign = cstyle.VerticalAlignment switch
                            {
                                VerticalAlignment.Top => VertAlign.Top,
                                VerticalAlignment.Center => VertAlign.Middle,
                                VerticalAlignment.Bottom => VertAlign.Bottom,
                                _ => VertAlign.Preserve,
                            },
                        };
                    }
                }
                else uniCell.Ignored = true;
            }
        }

        return new HtmlTable(uniTable);
    }

    public override void Disposing()
    {
        Sheet.Cursor = new Cursor(End.Row + 1, Start.Col);
        var outterArea = Scopes.Skip(1).FirstOrDefault();
        outterArea?.Update(_start, _end);
    }
}
