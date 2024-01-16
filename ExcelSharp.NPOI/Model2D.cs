using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelSharp.NPOI;

[DebuggerDisplay(@"{string.Join("","", RowNames)} & {string.Join("","", ColNames)}: {Value}")]
public class Model2D<T>
{
    public string[] RowNames { get; set; }
    public string[] ColNames { get; set; }
    public T Value { get; set; }
}
