using NStandard;
using System;

namespace ExcelSharp.NPOI.Utils;

public static class ConvertUtil
{
    public static object? Convert(object? obj, Type convertionType)
    {
        return obj switch
        {
            double d => convertionType switch
            {
                Type when convertionType == typeof(char) => (char)(int)d,
                Type when convertionType == typeof(char?) => (char?)(int)d,
                _ => ConvertEx.ChangeType(obj, convertionType),
            },
            string s => convertionType switch
            {
                Type when convertionType == typeof(Guid) => Guid.Parse(s),
                Type when convertionType == typeof(Guid?) => Guid.TryParse(s, out var guid) ? guid : null,
                _ => ConvertEx.ChangeType(obj, convertionType),
            },
            _ => ConvertEx.ChangeType(obj, convertionType),
        };
    }

    public static T? Convert<T>(object? obj) => (T?)Convert(obj, typeof(T));

}
