using NPOI.XSSF.UserModel;
using Richx;

namespace ExcelSharp.NPOI.Utils
{
    public static class XSSFColorUtil
    {
        public static IArgbColor GetRgbColor(XSSFColor color)
        {
            var bytes = color.RGB;
            if (bytes is null) return ExcelColor.Automatic;
            else return RgbColor.Create(bytes[0], bytes[1], bytes[2]);
        }

        public static XSSFColor GetXSSFColor(IArgbColor color)
        {
            if (color.Alpha == 0) return null;
            return new XSSFColor(new[] { color.Red, color.Green, color.Blue });
        }

    }
}
