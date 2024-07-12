using NPOI.XSSF.UserModel;

namespace ExcelSharp.NPOI.Utils;

public static class XSSFColorUtil
{
    public static RgbaColor GetRgbColor(XSSFColor color)
    {
        var bytes = color.RGB;
        if (bytes is null) return ExcelColor.Automatic;
        else return new RgbaColor(bytes[0], bytes[1], bytes[2]);
    }

    public static XSSFColor? GetXSSFColor(RgbaColor color)
    {
        if (color.A == 0) return null;
        return new XSSFColor([color.R, color.G, color.B]);
    }

}
