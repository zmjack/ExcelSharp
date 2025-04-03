using System;

namespace ExcelSharp.NPOI;

public partial class ExcelSheet : ICloneable
{
    object ICloneable.Clone() => Book.CloneSheet(SheetName);
    public ExcelSheet Clone() => (this as ICloneable)?.Clone() as ExcelSheet ?? throw new InvalidCastException("Not a cloneable object.");
    public ExcelSheet Clone(string newSheetName)
    {
        var clone = Clone();
        Book.SetSheetName(Book.GetSheetIndex(clone), newSheetName);
        return clone;
    }
}