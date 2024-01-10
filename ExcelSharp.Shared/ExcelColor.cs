using NStandard;
using System.Collections.Generic;
using System.Linq;

namespace ExcelSharp;

public static class ExcelColor
{
    public static readonly short AutomaticIndex = 64;

    public static readonly IRgbaColor Black = new RgbaColor(0xff000000);
    public static readonly IRgbaColor White = new RgbaColor(0xffffffff);
    public static readonly IRgbaColor Red = new RgbaColor(0xffff0000);
    public static readonly IRgbaColor BrightGreen = new RgbaColor(0xff00ff00);
    public static readonly IRgbaColor Blue = new RgbaColor(0xff0000ff);
    public static readonly IRgbaColor Yellow = new RgbaColor(0xffffff00);
    public static readonly IRgbaColor Pink = new RgbaColor(0xffff00ff);
    public static readonly IRgbaColor Turquoise = new RgbaColor(0xff00ffff);
    public static readonly IRgbaColor DarkRed = new RgbaColor(0xff800000);
    public static readonly IRgbaColor Green = new RgbaColor(0xff008000);
    public static readonly IRgbaColor DarkBlue = new RgbaColor(0xff000080);
    public static readonly IRgbaColor DarkYellow = new RgbaColor(0xff808000);
    public static readonly IRgbaColor Violet = new RgbaColor(0xff800080);
    public static readonly IRgbaColor Teal = new RgbaColor(0xff008080);
    public static readonly IRgbaColor Grey25Percent = new RgbaColor(0xffc0c0c0);
    public static readonly IRgbaColor Grey50Percent = new RgbaColor(0xff808080);
    public static readonly IRgbaColor CornflowerBlue = new RgbaColor(0xff9999ff);
    public static readonly IRgbaColor Maroon = new RgbaColor(0xff7f0000);
    public static readonly IRgbaColor LemonChiffon = new RgbaColor(0xffffffcc);
    public static readonly IRgbaColor Orchid = new RgbaColor(0xff660066);
    public static readonly IRgbaColor Coral = new RgbaColor(0xffff8080);
    public static readonly IRgbaColor RoyalBlue = new RgbaColor(0xff0066cc);
    public static readonly IRgbaColor LightCornflowerBlue = new RgbaColor(0xffccccff);
    public static readonly IRgbaColor SkyBlue = new RgbaColor(0xff00ccff);
    public static readonly IRgbaColor LightTurquoise = new RgbaColor(0xffccffff);
    public static readonly IRgbaColor LightGreen = new RgbaColor(0xffccffcc);
    public static readonly IRgbaColor LightYellow = new RgbaColor(0xffffff99);
    public static readonly IRgbaColor PaleBlue = new RgbaColor(0xff99ccff);
    public static readonly IRgbaColor Rose = new RgbaColor(0xffff99cc);
    public static readonly IRgbaColor Lavender = new RgbaColor(0xffcc99ff);
    public static readonly IRgbaColor Tan = new RgbaColor(0xffffcc99);
    public static readonly IRgbaColor LightBlue = new RgbaColor(0xff3366ff);
    public static readonly IRgbaColor Aqua = new RgbaColor(0xff33cccc);
    public static readonly IRgbaColor Lime = new RgbaColor(0xff99cc00);
    public static readonly IRgbaColor Gold = new RgbaColor(0xffffcc00);
    public static readonly IRgbaColor LightOrange = new RgbaColor(0xffff9900);
    public static readonly IRgbaColor Orange = new RgbaColor(0xffff6600);
    public static readonly IRgbaColor BlueGrey = new RgbaColor(0xff666699);
    public static readonly IRgbaColor Grey40Percent = new RgbaColor(0xff969696);
    public static readonly IRgbaColor DarkTeal = new RgbaColor(0xff003366);
    public static readonly IRgbaColor SeaGreen = new RgbaColor(0xff339966);
    public static readonly IRgbaColor DarkGreen = new RgbaColor(0xff003300);
    public static readonly IRgbaColor OliveGreen = new RgbaColor(0xff333300);
    public static readonly IRgbaColor Brown = new RgbaColor(0xff993300);
    public static readonly IRgbaColor Plum = new RgbaColor(0xff993366);
    public static readonly IRgbaColor Indigo = new RgbaColor(0xff333399);
    public static readonly IRgbaColor Grey80Percent = new RgbaColor(0xff333333);
    public static readonly IRgbaColor Automatic = new RgbaColor(0x00000000);

    private static readonly Dictionary<short, IRgbaColor> _dict = new()
    {
        [8] = Black,
        [9] = White,
        [10] = Red,
        [11] = BrightGreen,
        [12] = Blue,
        [13] = Yellow,
        [14] = Pink,
        [15] = Turquoise,
        [16] = DarkRed,
        [17] = Green,
        [18] = DarkBlue,
        [19] = DarkYellow,
        [20] = Violet,
        [21] = Teal,
        [22] = Grey25Percent,
        [23] = Grey50Percent,
        [24] = CornflowerBlue,
        [25] = Maroon,
        [26] = LemonChiffon,
        [28] = Orchid,
        [29] = Coral,
        [30] = RoyalBlue,
        [31] = LightCornflowerBlue,
        [40] = SkyBlue,
        [41] = LightTurquoise,
        [42] = LightGreen,
        [43] = LightYellow,
        [44] = PaleBlue,
        [45] = Rose,
        [46] = Lavender,
        [47] = Tan,
        [48] = LightBlue,
        [49] = Aqua,
        [50] = Lime,
        [51] = Gold,
        [52] = LightOrange,
        [53] = Orange,
        [54] = BlueGrey,
        [55] = Grey40Percent,
        [56] = DarkTeal,
        [57] = SeaGreen,
        [58] = DarkGreen,
        [59] = OliveGreen,
        [60] = Brown,
        [61] = Plum,
        [62] = Indigo,
        [63] = Grey80Percent,
        [64] = Automatic,
    };

    public static IRgbaColor GetColor(short index) => _dict[index];
    public static short GetIndex(IRgbaColor color) => _dict.FirstOrDefault(x => x.Value == color).Key;

    private static readonly RgbaColor[,] _themeColors = new RgbaColor[,]
    {
        {
            new RgbaColor(0xFFFFFF), new RgbaColor(0x000000), new RgbaColor(0xE7E6E6), new RgbaColor(0x44546A), new RgbaColor(0x5B9BD5),
            new RgbaColor(0xED7D31), new RgbaColor(0xA5A5A5), new RgbaColor(0xFFC000), new RgbaColor(0x4472C4), new RgbaColor(0x70AD47),
        },
        {
            new RgbaColor(0xF2F2F2), new RgbaColor(0x7F7F7F), new RgbaColor(0xD0CECE), new RgbaColor(0xD6DCE4), new RgbaColor(0xDEEBF6),
            new RgbaColor(0xFBE5D5), new RgbaColor(0xEDEDED), new RgbaColor(0xFFF2CC), new RgbaColor(0xD9E2F3), new RgbaColor(0xE2EFD9),
        },
        {
            new RgbaColor(0xD8D8D8), new RgbaColor(0x595959), new RgbaColor(0xAEABAB), new RgbaColor(0xADB9CA), new RgbaColor(0xBDD7EE),
            new RgbaColor(0xF7CBAC), new RgbaColor(0xDBDBDB), new RgbaColor(0xFEE599), new RgbaColor(0xB4C6E7), new RgbaColor(0xC5E0B3),
        },
        {
            new RgbaColor(0xBFBFBF), new RgbaColor(0x3F3F3F), new RgbaColor(0x757070), new RgbaColor(0x8496B0), new RgbaColor(0x9CC3E5),
            new RgbaColor(0xF4B183), new RgbaColor(0xC9C9C9), new RgbaColor(0xFFD965), new RgbaColor(0x8EAADB), new RgbaColor(0xA8D08D),
        },
        {
            new RgbaColor(0xA5A5A5), new RgbaColor(0x262626), new RgbaColor(0x3A3838), new RgbaColor(0x323F4F), new RgbaColor(0x2E75B5),
            new RgbaColor(0xC55A11), new RgbaColor(0x7B7B7B), new RgbaColor(0xBF9000), new RgbaColor(0x2F5496), new RgbaColor(0x538135),
        },
        {
            new RgbaColor(0x7F7F7F), new RgbaColor(0x0C0C0C), new RgbaColor(0x171616), new RgbaColor(0x222A35), new RgbaColor(0x1E4E79),
            new RgbaColor(0x833C0B), new RgbaColor(0x525252), new RgbaColor(0x7F6000), new RgbaColor(0x1F3864), new RgbaColor(0x375623),
        },
    };

    public static RgbaColor Theme(int row, int col)
    {
        return _themeColors[row, col];
    }

    private static readonly RgbaColor[,] _standardColors = new RgbaColor[,]
    {
        {
            new RgbaColor(0xBF0000), new RgbaColor(0xFE0000), new RgbaColor(0xFEBF00), new RgbaColor(0xFEFE00), new RgbaColor(0x91CE4F),
            new RgbaColor(0x00B050), new RgbaColor(0x00AFEE), new RgbaColor(0x006FBE), new RgbaColor(0x002060), new RgbaColor(0x6F309F),
        },
    };

    public static RgbaColor Standard(int row, int col)
    {
        return _standardColors[row, col];
    }
}
