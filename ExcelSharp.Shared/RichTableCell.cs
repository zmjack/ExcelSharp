using ExcelSharp.NPOI;
using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Text;

namespace ExcelSharp
{
    public class RichTableCell
    {
        public RichTableRow Row { get; private set; }
        public RichTableCellStyle Style { get; set; } = new RichTableCellStyle();

        internal RichTableCell(RichTableRow row)
        {
            Row = row;
        }

        public bool Ignored { get; internal set; }
        public int RowSpan { get; internal set; } = 1;
        public int RowOffset { get; internal set; }
        public int ColSpan { get; internal set; } = 1;
        public int ColOffset { get; internal set; }

        public string Comment { get; set; }
        public object Value { get; set; }
        public string Format { get; set; }
        public string Text
        {
            get
            {
                object value;
                if (Value is RichValue cvalue)
                {
                    value = cvalue.Value;
                    var style = cvalue.Style;

                    Style = new RichTableCellStyle
                    {
                        BackgroundColor = style.FillForegroundColor.Value,
                        Color = style.Font.FontColor.Value,
                        FontSize = style.Font.FontSize,
                        Bold = style.Font.IsBold,
                        TextAlign = style.Alignment == HorizontalAlignment.Center ? "center" : null,
                        VerticalAlign = style.VerticalAlignment == VerticalAlignment.Center ? "center" : null,

                        BorderTop = style.BorderTop != BorderStyle.None,
                        BorderBottom = style.BorderBottom != BorderStyle.None,
                        BorderLeft = style.BorderLeft != BorderStyle.None,
                        BorderRight = style.BorderRight != BorderStyle.None,
                    };
                }
                else value = Value;

                var str = value switch
                {
                    short v => v.ToString(Format),
                    int v => v.ToString(Format),
                    long v => v.ToString(Format),
                    ushort v => v.ToString(Format),
                    uint v => v.ToString(Format),
                    ulong v => v.ToString(Format),
                    float v => v.ToString(Format),
                    double v => v.ToString(Format),
                    DateTime v => v.ToString(Format),
                    decimal v => v.ToString(Format),
                    _ => value?.ToString(),
                };
                return str;
            }
        }

        public RichTableCell FullBorder()
        {
            Style.BorderTop = Style.BorderBottom = Style.BorderLeft = Style.BorderRight = true;
            return this;
        }

    }
}
