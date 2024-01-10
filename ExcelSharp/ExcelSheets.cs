using System.Collections;
using System.Collections.Generic;
using System.Linq;

namespace ExcelSharp;

public class ExcelSheets : ICollection<ExcelSheet>
{
    private readonly List<ExcelSheet> sheets = new List<ExcelSheet>();

    public int Count => ((ICollection<ExcelSheet>)sheets).Count;

    public bool IsReadOnly => ((ICollection<ExcelSheet>)sheets).IsReadOnly;

    public void Add(ExcelSheet item)
    {
        ((ICollection<ExcelSheet>)sheets).Add(item);
    }

    public void Clear()
    {
        ((ICollection<ExcelSheet>)sheets).Clear();
    }

    public bool Contains(ExcelSheet item)
    {
        return ((ICollection<ExcelSheet>)sheets).Contains(item);
    }

    public void CopyTo(ExcelSheet[] array, int arrayIndex)
    {
        ((ICollection<ExcelSheet>)sheets).CopyTo(array, arrayIndex);
    }

    public IEnumerator<ExcelSheet> GetEnumerator()
    {
        return ((IEnumerable<ExcelSheet>)sheets).GetEnumerator();
    }

    public bool Remove(ExcelSheet item)
    {
        return ((ICollection<ExcelSheet>)sheets).Remove(item);
    }

    IEnumerator IEnumerable.GetEnumerator()
    {
        return ((IEnumerable)sheets).GetEnumerator();
    }

    public ExcelSheet this[int index] => sheets.Skip(index).FirstOrDefault();

}
