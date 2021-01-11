using Richx;
using System.Collections.Generic;
using System.Linq;

namespace ExcelSharp
{
    public static class ExcelColor
    {
        public static short AutomaticIndex = 64;

        public static RgbColor Black = new RgbColor { Value = 0x000000 };
        public static RgbColor White = new RgbColor { Value = 0xffffff };
        public static RgbColor Red = new RgbColor { Value = 0xff0000 };
        public static RgbColor BrightGreen = new RgbColor { Value = 0x00ff00 };
        public static RgbColor Blue = new RgbColor { Value = 0x0000ff };
        public static RgbColor Yellow = new RgbColor { Value = 0xffff00 };
        public static RgbColor Pink = new RgbColor { Value = 0xff00ff };
        public static RgbColor Turquoise = new RgbColor { Value = 0x00ffff };
        public static RgbColor DarkRed = new RgbColor { Value = 0x800000 };
        public static RgbColor Green = new RgbColor { Value = 0x008000 };
        public static RgbColor DarkBlue = new RgbColor { Value = 0x000080 };
        public static RgbColor DarkYellow = new RgbColor { Value = 0x808000 };
        public static RgbColor Violet = new RgbColor { Value = 0x800080 };
        public static RgbColor Teal = new RgbColor { Value = 0x008080 };
        public static RgbColor Grey25Percent = new RgbColor { Value = 0xc0c0c0 };
        public static RgbColor Grey50Percent = new RgbColor { Value = 0x808080 };
        public static RgbColor CornflowerBlue = new RgbColor { Value = 0x9999ff };
        public static RgbColor Maroon = new RgbColor { Value = 0x7f0000 };
        public static RgbColor LemonChiffon = new RgbColor { Value = 0xffffcc };
        public static RgbColor Orchid = new RgbColor { Value = 0x660066 };
        public static RgbColor Coral = new RgbColor { Value = 0xff8080 };
        public static RgbColor RoyalBlue = new RgbColor { Value = 0x0066cc };
        public static RgbColor LightCornflowerBlue = new RgbColor { Value = 0xccccff };
        public static RgbColor SkyBlue = new RgbColor { Value = 0x00ccff };
        public static RgbColor LightTurquoise = new RgbColor { Value = 0xccffff };
        public static RgbColor LightGreen = new RgbColor { Value = 0xccffcc };
        public static RgbColor LightYellow = new RgbColor { Value = 0xffff99 };
        public static RgbColor PaleBlue = new RgbColor { Value = 0x99ccff };
        public static RgbColor Rose = new RgbColor { Value = 0xff99cc };
        public static RgbColor Lavender = new RgbColor { Value = 0xcc99ff };
        public static RgbColor Tan = new RgbColor { Value = 0xffcc99 };
        public static RgbColor LightBlue = new RgbColor { Value = 0x3366ff };
        public static RgbColor Aqua = new RgbColor { Value = 0x33cccc };
        public static RgbColor Lime = new RgbColor { Value = 0x99cc00 };
        public static RgbColor Gold = new RgbColor { Value = 0xffcc00 };
        public static RgbColor LightOrange = new RgbColor { Value = 0xff9900 };
        public static RgbColor Orange = new RgbColor { Value = 0xff6600 };
        public static RgbColor BlueGrey = new RgbColor { Value = 0x666699 };
        public static RgbColor Grey40Percent = new RgbColor { Value = 0x969696 };
        public static RgbColor DarkTeal = new RgbColor { Value = 0x003366 };
        public static RgbColor SeaGreen = new RgbColor { Value = 0x339966 };
        public static RgbColor DarkGreen = new RgbColor { Value = 0x003300 };
        public static RgbColor OliveGreen = new RgbColor { Value = 0x333300 };
        public static RgbColor Brown = new RgbColor { Value = 0x993300 };
        public static RgbColor Plum = new RgbColor { Value = 0x993366 };
        public static RgbColor Indigo = new RgbColor { Value = 0x333399 };
        public static RgbColor Grey80Percent = new RgbColor { Value = 0x333333 };
        public static RgbColor Automatic = new RgbColor { Value = 0x000000 };

        private static readonly Dictionary<short, RgbColor> _dict = new Dictionary<short, RgbColor>
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

        public static RgbColor GetColor(short index) => _dict[index];
        public static short GetIndex(RgbColor color) => _dict.FirstOrDefault(x => x.Value == color).Key;

    }
}
