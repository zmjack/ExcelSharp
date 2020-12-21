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
        public int RowIndex => Row.Index;
        public int Index { get; internal set; }
        public RichStyle Style { get; set; } = new RichStyle();
        public object _innerValue;

        internal RichCell(RichRow row, int index)
        {
            Row = row;
            Index = index;
        }

        public bool Merged
        {
            get
            {
                if (Ignored) return true;
                else if (RowSpan > 1 || ColSpan > 1) return true;
                else return false;
            }
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
                    Comment = richValue.Comment;
                    Format = richValue.Format;
                    if (richValue.Style is not null) Style = richValue.Style;
                }
                else _innerValue = value;
            }
        }
        public string Format { get; set; }
        public string Text
        {
            get
            {
                if (Format is not null)
                {
                    return Value switch
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
                }
                else return Value?.ToString();
            }
        }

    }
}
