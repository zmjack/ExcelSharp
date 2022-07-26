using System;
using System.Collections.Generic;
using System.Text;

namespace ExcelSharp.NPOI
{
    public class FetchAttribute : Attribute
    {
        public int Length { get; set; }

        public FetchAttribute(int length)
        {
            if (length < 0) throw new ArgumentException(nameof(length), "Non-negative number required.");

            Length = length;
        }
    }
}
