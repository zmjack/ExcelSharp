using NPOI.XSSF.UserModel;
using Richx;

namespace ExcelSharp.NPOI.Utils
{
    public static class XSSFColorUtil
    {
        public static RgbColor GetRgbColor(XSSFColor color)
        {
            var bytes = color.RGB;
            if (bytes is null) return ExcelColor.Automatic;
            else return new RgbColor { Red = bytes[0], Green = bytes[1], Blue = bytes[2] };
        }

        public static XSSFColor GetXSSFColor(RgbColor color)
        {
            return new XSSFColor(new[] { color.Red, color.Green, color.Blue });
        }

    }
}
