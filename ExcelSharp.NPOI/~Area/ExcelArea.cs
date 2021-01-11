using NPOI.SS.UserModel;
using NStandard;
using Richx;
using System.Linq;

namespace ExcelSharp.NPOI
{
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

        public SheetRange GetRange() => new SheetRange(Sheet, _start, _end);

        public HtmlTable ToHtmlTable()
        {
            var uniTable = new RichTable();
            for (int row = _start.Row; row <= _end.Row; row++)
            {
                var rowOffset = row - _start.Row;
                var uniRow = uniTable.Row(rowOffset);

                for (int col = _start.Col; col <= _end.Col; col++)
                {
                    var colOffset = col - _start.Col;
                    var uniCell = uniRow.Cell(colOffset);

                    var cell = Sheet[(row, col)];
                    var cstyle = cell.GetCStyle();

                    if (!cell.IsMergedCell || cell.IsMergedDefinitionCell)
                    {
                        uniCell.Value = cell.GetValue()?.For(value =>
                        {
                            var format = cstyle.DataFormat;
                            if (cell.CellType == CellType.Numeric && format != "General") return ((double)value).ToString(format);
                            else return value?.ToString();
                        });
                        uniCell.Comment = cell.CellComment?.String.String;
                        uniCell.RowSpan = cell.RowSpan;
                        uniCell.ColSpan = cell.ColSpan;

                        uniCell.Style = new RichStyle
                        {
                            BackgroundColor = cstyle.FillPattern != FillPattern.NoFill ? cstyle.FillForegroundColor.Value : null,
                            Color = cstyle.Font.FontColor.Value,
                            Bold = cstyle.Font.IsBold,
                            FontSize = cstyle.Font.FontSize,
                            BorderTop = cstyle.BorderTop != BorderStyle.None,
                            BorderBottom = cstyle.BorderBottom != BorderStyle.None,
                            BorderLeft = cstyle.BorderLeft != BorderStyle.None,
                            BorderRight = cstyle.BorderRight != BorderStyle.None,
                            TextAlign = cstyle.Alignment switch
                            {
                                HorizontalAlignment.Left => RichTextAlignment.Left,
                                HorizontalAlignment.Center => RichTextAlignment.Center,
                                HorizontalAlignment.Right => RichTextAlignment.Right,
                                _ => RichTextAlignment.Preserve,
                            },
                            VerticalAlign = cstyle.VerticalAlignment switch
                            {
                                VerticalAlignment.Top => RichVerticalAlignment.Top,
                                VerticalAlignment.Center => RichVerticalAlignment.Middle,
                                VerticalAlignment.Bottom => RichVerticalAlignment.Bottom,
                                _ => RichVerticalAlignment.Preserve,
                            },
                        };
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
            if (outterArea is not null) outterArea.Update(_start, _end);
        }
    }
}
