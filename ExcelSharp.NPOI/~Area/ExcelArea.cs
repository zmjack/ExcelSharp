using NStandard;
using System;
using System.Collections.Generic;
using System.Text;

namespace ExcelSharp.NPOI
{
    public class ExcelArea : IDisposable
    {
        public (int row, int col) Start { get; private set; }
        public (int row, int col) End { get; private set; }

        public ExcelArea(ExcelSheet sheet)
        {

        }

        public void Dispose() { }
    }
}
