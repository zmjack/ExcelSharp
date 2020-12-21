using System;
using System.Collections.Generic;
using System.Text;

namespace ExcelSharp
{
    public class RichValue
    {
        public object Value { get; set; }
        public string Format { get; set; }
        public RichStyle Style { get; set; }
    }
}
