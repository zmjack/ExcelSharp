using NStandard;
using System;
using System.Collections.Generic;
using System.Text;

namespace ExcelSharp.NPOI.Utils
{
    public static class ConvertUtil
    {
        public static object Convert(object obj, Type convertionType)
        {
            if (obj is string str && convertionType == typeof(Guid)) return Guid.Parse(str);
            else return ConvertEx.ChangeType(obj, convertionType);
        }

        public static T Convert<T>(object obj) => (T)Convert(obj, typeof(T));

    }
}
