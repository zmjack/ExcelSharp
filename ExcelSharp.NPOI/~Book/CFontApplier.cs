using NPOI.SS.UserModel;
using NStandard;
using System;
using System.Linq;

namespace ExcelSharp.NPOI;

public class CFontApplier : ICFont
{
    private CFontApplier() { }

    internal static CFontApplier Create(Action<CFontApplier>? init) => new CFontApplier().Pipe(x => init?.Invoke(x));

    public string FontName { get; set; } = "Calibri";
    public short FontSize { get; set; } = 11;

    public bool IsBold { get; set; } = false;
    public bool IsItalic { get; set; } = false;
    public bool IsStrikeout { get; set; } = false;

    public FontUnderlineType Underline { get; set; } = FontUnderlineType.None;
    public FontSuperScript TypeOffset { get; set; } = FontSuperScript.None;

    public IRgbaColor FontColor { get; set; } = ExcelColor.Black;

    public void Apply(CFont style)
    {
        //TODO: Use TypeReflectionCacheContainer to optimize it in the futrue.
        var props = typeof(ICFont).GetProperties().Where(prop => prop.CanWrite);
        foreach (var prop in props)
        {
            prop.SetValue(style, prop.GetValue(this));
        }
    }

}
