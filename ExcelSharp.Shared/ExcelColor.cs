using Richx;
using System.Collections.Generic;
using System.Linq;

namespace ExcelSharp
{
    public static class ExcelColor
    {
        public static short AutomaticIndex = 64;

        public static ArgbColor Black = ArgbColor.Create(0xff000000);
        public static ArgbColor White = ArgbColor.Create(0xffffffff);
        public static ArgbColor Red = ArgbColor.Create(0xffff0000);
        public static ArgbColor BrightGreen = ArgbColor.Create(0xff00ff00);
        public static ArgbColor Blue = ArgbColor.Create(0xff0000ff);
        public static ArgbColor Yellow = ArgbColor.Create(0xffffff00);
        public static ArgbColor Pink = ArgbColor.Create(0xffff00ff);
        public static ArgbColor Turquoise = ArgbColor.Create(0xff00ffff);
        public static ArgbColor DarkRed = ArgbColor.Create(0xff800000);
        public static ArgbColor Green = ArgbColor.Create(0xff008000);
        public static ArgbColor DarkBlue = ArgbColor.Create(0xff000080);
        public static ArgbColor DarkYellow = ArgbColor.Create(0xff808000);
        public static ArgbColor Violet = ArgbColor.Create(0xff800080);
        public static ArgbColor Teal = ArgbColor.Create(0xff008080);
        public static ArgbColor Grey25Percent = ArgbColor.Create(0xffc0c0c0);
        public static ArgbColor Grey50Percent = ArgbColor.Create(0xff808080);
        public static ArgbColor CornflowerBlue = ArgbColor.Create(0xff9999ff);
        public static ArgbColor Maroon = ArgbColor.Create(0xff7f0000);
        public static ArgbColor LemonChiffon = ArgbColor.Create(0xffffffcc);
        public static ArgbColor Orchid = ArgbColor.Create(0xff660066);
        public static ArgbColor Coral = ArgbColor.Create(0xffff8080);
        public static ArgbColor RoyalBlue = ArgbColor.Create(0xff0066cc);
        public static ArgbColor LightCornflowerBlue = ArgbColor.Create(0xffccccff);
        public static ArgbColor SkyBlue = ArgbColor.Create(0xff00ccff);
        public static ArgbColor LightTurquoise = ArgbColor.Create(0xffccffff);
        public static ArgbColor LightGreen = ArgbColor.Create(0xffccffcc);
        public static ArgbColor LightYellow = ArgbColor.Create(0xffffff99);
        public static ArgbColor PaleBlue = ArgbColor.Create(0xff99ccff);
        public static ArgbColor Rose = ArgbColor.Create(0xffff99cc);
        public static ArgbColor Lavender = ArgbColor.Create(0xffcc99ff);
        public static ArgbColor Tan = ArgbColor.Create(0xffffcc99);
        public static ArgbColor LightBlue = ArgbColor.Create(0xff3366ff);
        public static ArgbColor Aqua = ArgbColor.Create(0xff33cccc);
        public static ArgbColor Lime = ArgbColor.Create(0xff99cc00);
        public static ArgbColor Gold = ArgbColor.Create(0xffffcc00);
        public static ArgbColor LightOrange = ArgbColor.Create(0xffff9900);
        public static ArgbColor Orange = ArgbColor.Create(0xffff6600);
        public static ArgbColor BlueGrey = ArgbColor.Create(0xff666699);
        public static ArgbColor Grey40Percent = ArgbColor.Create(0xff969696);
        public static ArgbColor DarkTeal = ArgbColor.Create(0xff003366);
        public static ArgbColor SeaGreen = ArgbColor.Create(0xff339966);
        public static ArgbColor DarkGreen = ArgbColor.Create(0xff003300);
        public static ArgbColor OliveGreen = ArgbColor.Create(0xff333300);
        public static ArgbColor Brown = ArgbColor.Create(0xff993300);
        public static ArgbColor Plum = ArgbColor.Create(0xff993366);
        public static ArgbColor Indigo = ArgbColor.Create(0xff333399);
        public static ArgbColor Grey80Percent = ArgbColor.Create(0xff333333);
        public static ArgbColor Automatic = ArgbColor.Create(0x00000000);

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

    }
}
