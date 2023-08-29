using Richx;
using System.Collections.Generic;
using System.Linq;

namespace ExcelSharp
{
    public static class ExcelColor
    {
        public static readonly short AutomaticIndex = 64;

        public static readonly ArgbColor Black = new(0xff000000);
        public static readonly ArgbColor White = new(0xffffffff);
        public static readonly ArgbColor Red = new(0xffff0000);
        public static readonly ArgbColor BrightGreen = new(0xff00ff00);
        public static readonly ArgbColor Blue = new(0xff0000ff);
        public static readonly ArgbColor Yellow = new(0xffffff00);
        public static readonly ArgbColor Pink = new(0xffff00ff);
        public static readonly ArgbColor Turquoise = new(0xff00ffff);
        public static readonly ArgbColor DarkRed = new(0xff800000);
        public static readonly ArgbColor Green = new(0xff008000);
        public static readonly ArgbColor DarkBlue = new(0xff000080);
        public static readonly ArgbColor DarkYellow = new(0xff808000);
        public static readonly ArgbColor Violet = new(0xff800080);
        public static readonly ArgbColor Teal = new(0xff008080);
        public static readonly ArgbColor Grey25Percent = new(0xffc0c0c0);
        public static readonly ArgbColor Grey50Percent = new(0xff808080);
        public static readonly ArgbColor CornflowerBlue = new(0xff9999ff);
        public static readonly ArgbColor Maroon = new(0xff7f0000);
        public static readonly ArgbColor LemonChiffon = new(0xffffffcc);
        public static readonly ArgbColor Orchid = new(0xff660066);
        public static readonly ArgbColor Coral = new(0xffff8080);
        public static readonly ArgbColor RoyalBlue = new(0xff0066cc);
        public static readonly ArgbColor LightCornflowerBlue = new(0xffccccff);
        public static readonly ArgbColor SkyBlue = new(0xff00ccff);
        public static readonly ArgbColor LightTurquoise = new(0xffccffff);
        public static readonly ArgbColor LightGreen = new(0xffccffcc);
        public static readonly ArgbColor LightYellow = new(0xffffff99);
        public static readonly ArgbColor PaleBlue = new(0xff99ccff);
        public static readonly ArgbColor Rose = new(0xffff99cc);
        public static readonly ArgbColor Lavender = new(0xffcc99ff);
        public static readonly ArgbColor Tan = new(0xffffcc99);
        public static readonly ArgbColor LightBlue = new(0xff3366ff);
        public static readonly ArgbColor Aqua = new(0xff33cccc);
        public static readonly ArgbColor Lime = new(0xff99cc00);
        public static readonly ArgbColor Gold = new(0xffffcc00);
        public static readonly ArgbColor LightOrange = new(0xffff9900);
        public static readonly ArgbColor Orange = new(0xffff6600);
        public static readonly ArgbColor BlueGrey = new(0xff666699);
        public static readonly ArgbColor Grey40Percent = new(0xff969696);
        public static readonly ArgbColor DarkTeal = new(0xff003366);
        public static readonly ArgbColor SeaGreen = new(0xff339966);
        public static readonly ArgbColor DarkGreen = new(0xff003300);
        public static readonly ArgbColor OliveGreen = new(0xff333300);
        public static readonly ArgbColor Brown = new(0xff993300);
        public static readonly ArgbColor Plum = new(0xff993366);
        public static readonly ArgbColor Indigo = new(0xff333399);
        public static readonly ArgbColor Grey80Percent = new(0xff333333);
        public static readonly ArgbColor Automatic = new(0x00000000);

        private static readonly Dictionary<short, IArgbColor> _dict = new()
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

        public static IArgbColor GetColor(short index) => _dict[index];
        public static short GetIndex(IArgbColor color) => _dict.FirstOrDefault(x => x.Value == color).Key;


        private static readonly RgbColor[,] _themeColors = new RgbColor[,]
        {
            { 0xFFFFFF, 0x000000, 0xE7E6E6, 0x44546A, 0x5B9BD5, 0xED7D31, 0xA5A5A5, 0xFFC000, 0x4472C4, 0x70AD47 },
            { 0xF2F2F2, 0x7F7F7F, 0xD0CECE, 0xD6DCE4, 0x222A35, 0xFBE5D5, 0xEDEDED, 0x525252, 0xD9E2F3, 0xE2EFD9 },
            { 0xD8D8D8, 0x595959, 0xAEABAB, 0xADB9CA, 0xBDD7EE, 0xF7CBAC, 0xDBDBDB, 0xFEE599, 0xB4C6E7, 0xC5E0B3 },
            { 0xBFBFBF, 0x3F3F3F, 0x757070, 0x8496B0, 0x9CC3E5, 0xF4B183, 0xC9C9C9, 0xFFD965, 0xB4C6E7, 0xA8D08D },
            { 0xA5A5A5, 0x262626, 0x3A3838, 0x323F4F, 0x2E75B5, 0xC55A11, 0x7B7B7B, 0xBF9000, 0x2F5496, 0xA8D08D },
            { 0x7F7F7F, 0x0C0C0C, 0x171616, 0x222A35, 0x1E4E79, 0x833C0B, 0x525252, 0x7F6000, 0x2F5496, 0x375623 },
        };
        public static RgbColor Theme(int row, int col)
        {
            return _themeColors[row, col];
        }

        private static readonly RgbColor[,] _standardColors = new RgbColor[,]
        {
            { 0xBF0000, 0xFE0000, 0xFEBF00, 0xFEFE00, 0x91CE4F, 0x00B050, 0x00AFEE, 0x006FBE, 0x002060, 0x6F309F },
        };
        public static RgbColor Standard(int row, int col)
        {
            return _standardColors[row, col];
        }
    }
}
