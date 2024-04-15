using System.Diagnostics;

namespace ExcelSharp.NPOI;

[DebuggerDisplay(@"{string.Join("","", RowNames)} & {string.Join("","", ColNames)}: {Value}")]
public class Model2D<T>
{
    public string?[] RowNames { get; set; } = [];
    public string?[] ColNames { get; set; } = [];
    public T? Value { get; set; }
}
