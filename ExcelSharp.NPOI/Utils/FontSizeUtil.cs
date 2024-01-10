namespace ExcelSharp.NPOI.Utils;

public static class FontSizeUtil
{
    public const float RATE = 1.39325842696629f;
    public static float GetSKSize(float emSize) => emSize * RATE;
}
