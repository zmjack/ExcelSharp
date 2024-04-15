using NPOI.XSSF.UserModel;

namespace ExcelSharp.NPOI.Utils;

public static class XSSFColorUtil
{
    public static IRgbaColor GetRgbColor(XSSFColor color)
    {
        var bytes = color.RGB;
        if (bytes is null) return ExcelColor.Automatic;
        else return new RgbaColor(bytes[0], bytes[1], bytes[2]);
    }

    public static XSSFColor? GetXSSFColor(IRgbaColor color)
    {
        if (color.Alpha == 0) return null;
        return new XSSFColor([color.Red, color.Green, color.Blue]);
    }

}
