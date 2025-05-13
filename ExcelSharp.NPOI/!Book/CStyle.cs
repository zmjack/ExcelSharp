using ExcelSharp.NPOI.Utils;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using NStandard;
using System.Linq;
using VerticalAlignment = NPOI.SS.UserModel.VerticalAlignment;

namespace ExcelSharp.NPOI;

public class CStyle : ICStyle
{
    public ExcelBook Book { get; }
    public ICellStyle CellStyle { get; }

    internal CStyle(ExcelBook book)
        : this(book, book.MapedWorkbook.CreateCellStyle())
    { }

    internal CStyle(ExcelBook book, ICellStyle cellStyle)
    {
        Book = book;
        CellStyle = cellStyle;
    }

    public short Index => CellStyle.Index;

    #region Alignment
    public HorizontalAlignment Alignment
    {
        get => CellStyle.Alignment;
        set => CellStyle.Alignment = value;
    }
    public VerticalAlignment VerticalAlignment
    {
        get => CellStyle.VerticalAlignment;
        set => CellStyle.VerticalAlignment = value;
    }
    #endregion

    #region Border
    public BorderStyle BorderLeft
    {
        get => CellStyle.BorderLeft;
        set => CellStyle.BorderLeft = value;
    }
    public RgbaColor LeftBorderColor
    {
        get => CellStyle.LeftBorderColor > 0 ? ExcelColor.GetColor(CellStyle.LeftBorderColor)
            : (CellStyle as XSSFCellStyle)?.LeftBorderXSSFColor?.Pipe(XSSFColorUtil.GetRgbColor) ?? ExcelColor.Automatic;
        set
        {
            var index = ExcelColor.GetIndex(value);
            CellStyle.LeftBorderColor = index;
            if (CellStyle is XSSFCellStyle)
            {
                (CellStyle as XSSFCellStyle)!.SetLeftBorderColor(index == ExcelColor.AutomaticIndex ? null : XSSFColorUtil.GetXSSFColor(value));
            }
        }
    }

    public BorderStyle BorderRight
    {
        get => CellStyle.BorderRight;
        set => CellStyle.BorderRight = value;
    }
    public RgbaColor RightBorderColor
    {
        get => CellStyle.RightBorderColor > 0 ? ExcelColor.GetColor(CellStyle.RightBorderColor)
            : (CellStyle as XSSFCellStyle)?.RightBorderXSSFColor?.Pipe(XSSFColorUtil.GetRgbColor) ?? ExcelColor.Automatic;
        set
        {
            var index = ExcelColor.GetIndex(value);
            CellStyle.RightBorderColor = index;
            if (CellStyle is XSSFCellStyle)
            {
                (CellStyle as XSSFCellStyle)!.SetRightBorderColor(index == ExcelColor.AutomaticIndex ? null : XSSFColorUtil.GetXSSFColor(value));
            }
        }
    }

    public BorderStyle BorderTop
    {
        get => CellStyle.BorderTop;
        set => CellStyle.BorderTop = value;
    }
    public RgbaColor TopBorderColor
    {
        get => CellStyle.TopBorderColor > 0 ? ExcelColor.GetColor(CellStyle.TopBorderColor)
            : (CellStyle as XSSFCellStyle)?.TopBorderXSSFColor?.Pipe(XSSFColorUtil.GetRgbColor) ?? ExcelColor.Automatic;
        set
        {
            var index = ExcelColor.GetIndex(value);
            CellStyle.TopBorderColor = index;
            if (CellStyle is XSSFCellStyle)
            {
                (CellStyle as XSSFCellStyle)!.SetTopBorderColor(index == ExcelColor.AutomaticIndex ? null : XSSFColorUtil.GetXSSFColor(value));
            }
        }
    }

    public BorderStyle BorderBottom
    {
        get => CellStyle.BorderBottom;
        set => CellStyle.BorderBottom = value;
    }
    public RgbaColor BottomBorderColor
    {
        get => CellStyle.BottomBorderColor > 0 ? ExcelColor.GetColor(CellStyle.BottomBorderColor)
            : (CellStyle as XSSFCellStyle)?.BottomBorderXSSFColor?.Pipe(XSSFColorUtil.GetRgbColor) ?? ExcelColor.Automatic;
        set
        {
            var index = ExcelColor.GetIndex(value);
            CellStyle.BottomBorderColor = index;
            if (CellStyle is XSSFCellStyle)
            {
                (CellStyle as XSSFCellStyle)!.SetBottomBorderColor(index == ExcelColor.AutomaticIndex ? null : XSSFColorUtil.GetXSSFColor(value));
            }
        }
    }

    public BorderStyle BorderDiagonalLineStyle
    {
        get => CellStyle.BorderDiagonalLineStyle;
        set => CellStyle.BorderDiagonalLineStyle = value;
    }
    public RgbaColor BorderDiagonalColor
    {
        get => CellStyle.BorderDiagonalColor > 0 ? ExcelColor.GetColor(CellStyle.BorderDiagonalColor)
            : (CellStyle as XSSFCellStyle)?.DiagonalBorderXSSFColor?.Pipe(XSSFColorUtil.GetRgbColor) ?? ExcelColor.Automatic;
        set
        {
            var index = ExcelColor.GetIndex(value);
            CellStyle.BorderDiagonalColor = index;
            if (CellStyle is XSSFCellStyle)
            {
                (CellStyle as XSSFCellStyle)!.SetDiagonalBorderColor(index == ExcelColor.AutomaticIndex ? null : XSSFColorUtil.GetXSSFColor(value));
            }
        }
    }
    public BorderDiagonal BorderDiagonal
    {
        get => CellStyle.BorderDiagonal;
        set => CellStyle.BorderDiagonal = value;
    }
    #endregion

    #region Fill
    public FillPattern FillPattern
    {
        get => CellStyle.FillPattern;
        set => CellStyle.FillPattern = value;
    }

    public RgbaColor FillBackgroundColor
    {
        get
        {
            if (CellStyle.FillBackgroundColor > 0)
            {
                return ExcelColor.GetColor(CellStyle.FillBackgroundColor);
            }
            else
            {
                var xss = (CellStyle as XSSFCellStyle)?.FillBackgroundXSSFColor;
                return xss?.Pipe(XSSFColorUtil.GetRgbColor) ?? ExcelColor.Automatic;
            }
        }
        set
        {
            var index = ExcelColor.GetIndex(value);
            CellStyle.FillBackgroundColor = index;
            if (CellStyle is XSSFCellStyle)
            {
                (CellStyle as XSSFCellStyle)!.FillBackgroundXSSFColor = index == ExcelColor.AutomaticIndex ? null : XSSFColorUtil.GetXSSFColor(value);
            }
        }
    }

    public RgbaColor FillForegroundColor
    {
        get
        {
            if (CellStyle.FillForegroundColor > 0)
            {
                return ExcelColor.GetColor(CellStyle.FillForegroundColor);
            }
            else
            {
                var xss = (CellStyle as XSSFCellStyle)?.FillForegroundXSSFColor;
                return xss?.Pipe(XSSFColorUtil.GetRgbColor) ?? ExcelColor.Automatic;
            }
        }
        set
        {
            var index = (short)ExcelColor.GetIndex(value);
            CellStyle.FillForegroundColor = index;
            if (CellStyle is XSSFCellStyle)
            {
                (CellStyle as XSSFCellStyle)!.FillForegroundXSSFColor = index == ExcelColor.AutomaticIndex ? null : XSSFColorUtil.GetXSSFColor(value);
            }
        }
    }
    #endregion

    #region Font
    public CFont Font
    {
        get => Book.CFontAt(CellStyle.FontIndex);
        set => CellStyle.SetFont(value.Font);
    }
    #endregion

    #region DataFormat
    public string DataFormat
    {
        get => CellStyle.GetDataFormatString();
        set => CellStyle.DataFormat = value is null ? (short)0 : Book.GetDataFormat(value);
    }
    #endregion

    #region Others
    public short Rotation
    {
        get => CellStyle.Rotation;
        set => CellStyle.Rotation = value;
    }
    public short Indention
    {
        get => CellStyle.Indention;
        set => CellStyle.Indention = value;
    }
    public bool WrapText
    {
        get => CellStyle.WrapText;
        set => CellStyle.WrapText = value;
    }
    public bool IsLocked
    {
        get => CellStyle.IsLocked;
        set => CellStyle.IsLocked = value;
    }
    public bool IsHidden
    {
        get => CellStyle.IsHidden;
        set => CellStyle.IsHidden = value;
    }
    public bool ShrinkToFit
    {
        get => CellStyle.ShrinkToFit;
        set => CellStyle.ShrinkToFit = value;
    }
    #endregion

    public CStyleApplier GetApplier()
    {
        return CStyleApplier.Create(x =>
        {
            // Alignment
            x.Alignment = Alignment;
            x.VerticalAlignment = VerticalAlignment;

            // Border
            x.BorderLeft = BorderLeft;
            x.LeftBorderColor = LeftBorderColor;
            x.BorderRight = BorderRight;
            x.RightBorderColor = RightBorderColor;
            x.BorderTop = BorderTop;
            x.TopBorderColor = TopBorderColor;
            x.BorderBottom = BorderBottom;
            x.BottomBorderColor = BottomBorderColor;
            x.BorderDiagonalLineStyle = BorderDiagonalLineStyle;
            x.BorderDiagonalColor = BorderDiagonalColor;
            x.BorderDiagonal = BorderDiagonal;

            // Fill
            x.FillPattern = FillPattern;
            x.FillBackgroundColor = FillBackgroundColor;
            x.FillForegroundColor = FillForegroundColor;

            // Font
            x.Font = CFontApplier.Create(f =>
            {
                f.FontName = Font.FontName;
                f.FontSize = Font.FontSize;
                f.IsBold = Font.IsBold;
                f.IsItalic = Font.IsItalic;
                f.IsStrikeout = Font.IsStrikeout;
                f.Underline = Font.Underline;
                f.TypeOffset = Font.TypeOffset;
                f.FontColor = Font.FontColor;
            });

            // DataFormat
            x.DataFormat = DataFormat;

            // Others
            x.Rotation = Rotation;
            x.Indention = Indention;
            x.WrapText = WrapText;
            x.IsLocked = IsLocked;
            x.IsHidden = IsHidden;
            x.ShrinkToFit = ShrinkToFit;
        });
    }

    internal bool InterfaceValuesEqual(CStyleApplier obj)
    {
        var instance = obj as ICStyle;
        if (instance is null) return false;

        //TODO: Use TypeReflectionCacheContainer to optimize it in the futrue.
        var props = typeof(ICStyle).GetProperties().Where(prop => prop.CanWrite);
        return props.All(prop => Equals(prop.GetValue(this), prop.GetValue(instance)))
            && Font.InterfaceValuesEqual(obj.Font);
    }

}
