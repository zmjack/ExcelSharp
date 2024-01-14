using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelSharp.NPOI;

[DebuggerDisplay("{RowName} & {ColName}: {Value}")]
public class Model2D<T>
{
    public string RowName { get; set; }
    public string ColName { get; set; }
    public T Value { get; set; }
}
