using System;

namespace ExcelSharp.NPOI;

[AttributeUsage(AttributeTargets.Property, AllowMultiple = false)]
public class FetchAttribute : Attribute
{
    public int Length { get; set; }

    public FetchAttribute(int length)
    {
        if (length < 0) throw new ArgumentException("Non-negative number required.", nameof(length));
        Length = length;
    }
}

[AttributeUsage(AttributeTargets.Property, AllowMultiple = false)]
public class DataColumnAttribute(params string[] names) : Attribute
{
    public string[] Names { get; set; } = names;
}
