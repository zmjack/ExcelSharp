using ExcelSharp.NPOI.Utils;
using NPOI.SS.Formula.Functions;
using NStandard;
using SkiaSharp;
using System;

namespace ExcelSharp.NPOI;

internal class SkiaUtil : IDisposable
{
    private bool disposedValue;

    private readonly SKTypeface _face;
    private readonly SKPaint _paint;

    public SkiaUtil(string fontName, short fontSize)
    {
        _face = SKTypeface.FromFamilyName(fontName);
        _paint = new SKPaint(new SKFont(_face, FontSizeUtil.GetSKSize(fontSize)));
    }

    public float GetPaintWidth(string? value)
    {
        if (value.IsNullOrEmpty()) return 0;

        var rect = new SKRect();
        float width = 0;
        foreach (var valueRow in value!.NormalizeNewLine().Split(Environment.NewLine, StringSplitOptions.RemoveEmptyEntries))
        {
            _paint.MeasureText(valueRow, ref rect);
            if (rect.Width > width) width = rect.Width;
        }
        return width;
    }

    protected virtual void Dispose(bool disposing)
    {
        if (!disposedValue)
        {
            if (disposing)
            {
                // TODO: dispose managed state (managed objects)
                _paint.Dispose();
                _face.Dispose();
            }

            // TODO: free unmanaged resources (unmanaged objects) and override finalizer
            // TODO: set large fields to null
            disposedValue = true;
        }
    }

    // // TODO: override finalizer only if 'Dispose(bool disposing)' has code to free unmanaged resources
    // ~SkiaUtil()
    // {
    //     // Do not change this code. Put cleanup code in 'Dispose(bool disposing)' method
    //     Dispose(disposing: false);
    // }

    public void Dispose()
    {
        // Do not change this code. Put cleanup code in 'Dispose(bool disposing)' method
        Dispose(disposing: true);
        GC.SuppressFinalize(this);
    }
}
