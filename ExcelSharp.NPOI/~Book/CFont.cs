using ExcelSharp.NPOI.Utils;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using NStandard;
using System.Linq;

namespace ExcelSharp.NPOI;

public class CFont : ICFont
{
    public ExcelBook Book { get; }
    public IFont Font { get; }

    internal CFont(ExcelBook book)
        : this(book, book.MapedWorkbook.CreateFont())
    { }

    internal CFont(ExcelBook book, IFont cellStyle)
    {
        Book = book;
        Font = cellStyle;
    }

    public short Index => Font.Index;

    public string? FontName
    {
        get => Font.FontName;
        set => Font.FontName = value;
    }
    public short FontSize
    {
        get => Font.FontHeightInPoints;
        set => Font.FontHeightInPoints = value;
    }

    public bool IsBold
    {
        get => Font.IsBold;
        set => Font.IsBold = value;
    }
    public bool IsItalic
    {
        get => Font.IsItalic;
        set => Font.IsItalic = value;
    }
    public bool IsStrikeout
    {
        get => Font.IsStrikeout;
        set => Font.IsStrikeout = value;
    }

    public FontUnderlineType Underline
    {
        get => Font.Underline;
        set => Font.Underline = value;
    }
    public FontSuperScript TypeOffset
    {
        get => Font.TypeOffset;
        set => Font.TypeOffset = value;
    }

    public IRgbaColor FontColor
    {
        get => Font.Color > 0 ? ExcelColor.GetColor(Font.Color)
            : (Font as XSSFFont)?.GetXSSFColor()?.Pipe(XSSFColorUtil.GetRgbColor) ?? ExcelColor.Automatic;
        set
        {
            var index = ExcelColor.GetIndex(value);
            Font.Color = index;
            if (Font is XSSFFont)
            {
                (Font as XSSFFont)!.SetColor(index == ExcelColor.AutomaticIndex ? null : XSSFColorUtil.GetXSSFColor(value));
            }
        }
    }

    internal bool InterfaceValuesEqual(CFontApplier obj)
    {
        if (obj is not ICFont instance) return false;

        //TODO: Use TypeReflectionCacheContainer to optimize it in the futrue.
        var props = typeof(ICFont).GetProperties().Where(prop => prop.CanWrite);
        return props.All(prop => Equals(prop.GetValue(this), prop.GetValue(instance)));
    }

}
