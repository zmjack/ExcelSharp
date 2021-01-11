using Richx;
using System;
using System.Collections.Generic;
using System.Text;

namespace ExcelSharp
{
    public static class IndexColor
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

        public static RgbColor GetColor(int index)
        {
            switch (index)
            {
                case 8: return Black;
                case 9: return White;
                case 10: return Red;
                case 11: return BrightGreen;
                case 12: return Blue;
                case 13: return Yellow;
                case 14: return Pink;
                case 15: return Turquoise;
                case 16: return DarkRed;
                case 17: return Green;
                case 18: return DarkBlue;
                case 19: return DarkYellow;
                case 20: return Violet;
                case 21: return Teal;
                case 22: return Grey25Percent;
                case 23: return Grey50Percent;
                case 24: return CornflowerBlue;
                case 25: return Maroon;
                case 26: return LemonChiffon;
                case 28: return Orchid;
                case 29: return Coral;
                case 30: return RoyalBlue;
                case 31: return LightCornflowerBlue;
                case 40: return SkyBlue;
                case 41: return LightTurquoise;
                case 42: return LightGreen;
                case 43: return LightYellow;
                case 44: return PaleBlue;
                case 45: return Rose;
                case 46: return Lavender;
                case 47: return Tan;
                case 48: return LightBlue;
                case 49: return Aqua;
                case 50: return Lime;
                case 51: return Gold;
                case 52: return LightOrange;
                case 53: return Orange;
                case 54: return BlueGrey;
                case 55: return Grey40Percent;
                case 56: return DarkTeal;
                case 57: return SeaGreen;
                case 58: return DarkGreen;
                case 59: return OliveGreen;
                case 60: return Brown;
                case 61: return Plum;
                case 62: return Indigo;
                case 63: return Grey80Percent;
                default:
                case 64: return Automatic;
            }
        }
    }
}
