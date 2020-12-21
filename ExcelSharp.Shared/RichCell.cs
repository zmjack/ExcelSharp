using ExcelSharp.NPOI;
using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Text;

namespace ExcelSharp
{
    public class RichCell
    {
        public RichRow Row { get; private set; }
        public RichStyle Style { get; set; } = new RichStyle();
        public object _innerValue;

        internal RichCell(RichRow row)
        {
            Row = row;
        }

        public bool Ignored { get; internal set; }
        public int RowSpan { get; internal set; } = 1;
        public int RowOffset { get; internal set; }
        public int ColSpan { get; internal set; } = 1;
        public int ColOffset { get; internal set; }

        public string Comment { get; set; }
        public object Value
        {
            get => _innerValue;
            set
            {
                if (value is RichValue richValue)
                {
                    _innerValue = richValue.Value;
                    Style = richValue.Style;
                }
                else _innerValue = value;
            }
        }
        public string Format { get; set; }
        public string Text
        {
            get
            {
                var str = Value switch
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
                    _ => Value?.ToString(),
                };
                return str;
            }
        }

    }
}
