using System;
using System.Collections.Generic;
using System.Linq;

namespace ExcelSharp;

public static class ExcelColor
{
    public static readonly short AutomaticIndex = 64;

    public static readonly RgbaColor Black = new(0xff000000);
    public static readonly RgbaColor White = new(0xffffffff);
    public static readonly RgbaColor Red = new(0xffff0000);
    public static readonly RgbaColor BrightGreen = new(0xff00ff00);
    public static readonly RgbaColor Blue = new(0xff0000ff);
    public static readonly RgbaColor Yellow = new(0xffffff00);
    public static readonly RgbaColor Pink = new(0xffff00ff);
    public static readonly RgbaColor Turquoise = new(0xff00ffff);
    public static readonly RgbaColor DarkRed = new(0xff800000);
    public static readonly RgbaColor Green = new(0xff008000);
    public static readonly RgbaColor DarkBlue = new(0xff000080);
    public static readonly RgbaColor DarkYellow = new(0xff808000);
    public static readonly RgbaColor Violet = new(0xff800080);
    public static readonly RgbaColor Teal = new(0xff008080);
    public static readonly RgbaColor Grey25Percent = new(0xffc0c0c0);
    public static readonly RgbaColor Grey50Percent = new(0xff808080);
    public static readonly RgbaColor CornflowerBlue = new(0xff9999ff);
    public static readonly RgbaColor Maroon = new(0xff7f0000);
    public static readonly RgbaColor LemonChiffon = new(0xffffffcc);
    public static readonly RgbaColor Orchid = new(0xff660066);
    public static readonly RgbaColor Coral = new(0xffff8080);
    public static readonly RgbaColor RoyalBlue = new(0xff0066cc);
    public static readonly RgbaColor LightCornflowerBlue = new(0xffccccff);
    public static readonly RgbaColor SkyBlue = new(0xff00ccff);
    public static readonly RgbaColor LightTurquoise = new(0xffccffff);
    public static readonly RgbaColor LightGreen = new(0xffccffcc);
    public static readonly RgbaColor LightYellow = new(0xffffff99);
    public static readonly RgbaColor PaleBlue = new(0xff99ccff);
    public static readonly RgbaColor Rose = new(0xffff99cc);
    public static readonly RgbaColor Lavender = new(0xffcc99ff);
    public static readonly RgbaColor Tan = new(0xffffcc99);
    public static readonly RgbaColor LightBlue = new(0xff3366ff);
    public static readonly RgbaColor Aqua = new(0xff33cccc);
    public static readonly RgbaColor Lime = new(0xff99cc00);
    public static readonly RgbaColor Gold = new(0xffffcc00);
    public static readonly RgbaColor LightOrange = new(0xffff9900);
    public static readonly RgbaColor Orange = new(0xffff6600);
    public static readonly RgbaColor BlueGrey = new(0xff666699);
    public static readonly RgbaColor Grey40Percent = new(0xff969696);
    public static readonly RgbaColor DarkTeal = new(0xff003366);
    public static readonly RgbaColor SeaGreen = new(0xff339966);
    public static readonly RgbaColor DarkGreen = new(0xff003300);
    public static readonly RgbaColor OliveGreen = new(0xff333300);
    public static readonly RgbaColor Brown = new(0xff993300);
    public static readonly RgbaColor Plum = new(0xff993366);
    public static readonly RgbaColor Indigo = new(0xff333399);
    public static readonly RgbaColor Grey80Percent = new(0xff333333);
    public static readonly RgbaColor Automatic = new(0x00000000, 0);

    private static readonly Dictionary<short, RgbaColor> _dict = new()
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

    public static RgbaColor GetColor(short index) => _dict[index];
    public static short GetIndex(RgbaColor color) => _dict.FirstOrDefault(x => x.Value == color).Key;

    private static readonly RgbaColor[,] _themeColors = new RgbaColor[,]
    {
        {
            new(0xFFFFFF), new(0x000000), new(0xE7E6E6), new(0x44546A), new(0x5B9BD5),
            new(0xED7D31), new(0xA5A5A5), new(0xFFC000), new(0x4472C4), new(0x70AD47),
        },
        {
            new(0xF2F2F2), new(0x7F7F7F), new(0xD0CECE), new(0xD6DCE4), new(0xDEEBF6),
            new(0xFBE5D5), new(0xEDEDED), new(0xFFF2CC), new(0xD9E2F3), new(0xE2EFD9),
        },
        {
            new(0xD8D8D8), new(0x595959), new(0xAEABAB), new(0xADB9CA), new(0xBDD7EE),
            new(0xF7CBAC), new(0xDBDBDB), new(0xFEE599), new(0xB4C6E7), new(0xC5E0B3),
        },
        {
            new(0xBFBFBF), new(0x3F3F3F), new(0x757070), new(0x8496B0), new(0x9CC3E5),
            new(0xF4B183), new(0xC9C9C9), new(0xFFD965), new(0x8EAADB), new(0xA8D08D),
        },
        {
            new(0xA5A5A5), new(0x262626), new(0x3A3838), new(0x323F4F), new(0x2E75B5),
            new(0xC55A11), new(0x7B7B7B), new(0xBF9000), new(0x2F5496), new(0x538135),
        },
        {
            new(0x7F7F7F), new(0x0C0C0C), new(0x171616), new(0x222A35), new(0x1E4E79),
            new(0x833C0B), new(0x525252), new(0x7F6000), new(0x1F3864), new(0x375623),
        },
    };

    public static RgbaColor Theme(int row, int col)
    {
        if (row < 0) throw new ArgumentOutOfRangeException(nameof(row), "Row index can not be less than 0.");
        if (row > 6) throw new ArgumentOutOfRangeException(nameof(row), "Row index can not be greater than 6.");

        if (col < 0) throw new ArgumentOutOfRangeException(nameof(row), "Column index can not be less than 0.");
        if (col > 9) throw new ArgumentOutOfRangeException(nameof(row), "Column index can not be greater than 9.");

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
        if (row < 0) throw new ArgumentOutOfRangeException(nameof(row), "Row index can not be less than 0.");
        if (row > 0) throw new ArgumentOutOfRangeException(nameof(row), "Row index can not be greater than 0.");

        if (col < 0) throw new ArgumentOutOfRangeException(nameof(row), "Column index can not be less than 0.");
        if (col > 9) throw new ArgumentOutOfRangeException(nameof(row), "Column index can not be greater than 9.");

        return _standardColors[row, col];
    }

}
