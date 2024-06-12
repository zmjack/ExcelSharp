using ExcelSharp.NPOI;
using ExcelSharp.Test.Models;
using NStandard;
using System;
using System.Linq;
using Xunit;

namespace ExcelSharp.Test;

public class FetchTests
{
    [Fact]
    public void DateFetchTest_SheetColumnAttribute()
    {
        var book = new ExcelBook("excel-sharp.xlsx");
        {
            var sheet = book.GetSheet("Fetch");
            var models = sheet.Fetch<DateFetchModel>("A5");
            var date = new DateTime(2000, 1, 1);

            var valueModel = models[0];
            Assert.Equal(date, valueModel.Date);
            Assert.Equal(date, valueModel.StringDate);
            Assert.Equal(date, valueModel.NumberDate);
            Assert.Equal(date, valueModel.FormulaDate);

            var blankModel = models[1];
            Assert.Equal(DateTime.MinValue, blankModel.Date);
            Assert.Equal(DateTime.MinValue, blankModel.StringDate);
            Assert.Equal(DateTime.MinValue, blankModel.NumberDate);
            Assert.Equal(DateTime.MinValue, blankModel.FormulaDate);
        }

        {
            var sheet = book.GetSheet("2D");
            var models = sheet.Fetch2D<double?>("A1");
            Assert.Equal(21, models.Sum(x => x.Value));
        }
        {
            var sheet = book.GetSheet("2D");
            var models = sheet.Fetch2D<double?>("A6");
            Assert.Equal(21, models.Sum(x => x.Value));
        }
        {
            var sheet = book.GetSheet("2D");
            var models = sheet.Fetch2D<double?>("A12", "B12");
            Assert.Equal(21, models.Sum(x => x.Value));
        }
    }

    [Fact]
    public void DateFetchTest()
    {
        var book = new ExcelBook("excel-sharp.xlsx");
        var sheet = book.GetSheet("Fetch");

        var models = sheet.Fetch<DateFetchModel>("A6", x => new
        {
            x.Name,
            x.Date,
            x.StringDate,
            x.NumberDate,
            x.FormulaDate,
        });
        var date = new DateTime(2000, 1, 1);

        var valueModel = models[0];
        Assert.Equal(date, valueModel.Date);
        Assert.Equal(date, valueModel.StringDate);
        Assert.Equal(date, valueModel.NumberDate);
        Assert.Equal(date, valueModel.FormulaDate);

        var blankModel = models[1];
        Assert.Equal(DateTime.MinValue, blankModel.Date);
        Assert.Equal(DateTime.MinValue, blankModel.StringDate);
        Assert.Equal(DateTime.MinValue, blankModel.NumberDate);
        Assert.Equal(DateTime.MinValue, blankModel.FormulaDate);
    }

    [Fact]
    public void NullableDateFetchTest()
    {
        var book = new ExcelBook("excel-sharp.xlsx");
        var sheet = book.GetSheet("Fetch");
        var models = sheet.Fetch<NullableDateFetchModel>("A11", x => new
        {
            x.Name,
            x.Date,
            x.StringDate,
            x.NumberDate,
            x.FormulaDate,
        });
        var date = new DateTime(2000, 1, 1);

        var valueModel = models[0];
        Assert.Equal(date, valueModel.Date);
        Assert.Equal(date, valueModel.StringDate);
        Assert.Equal(date, valueModel.NumberDate);
        Assert.Equal(date, valueModel.FormulaDate);

        var blankModel = models[1];
        Assert.Null(blankModel.Date);
        Assert.Null(blankModel.StringDate);
        Assert.Null(blankModel.NumberDate);
        Assert.Null(blankModel.FormulaDate);
    }

    [Fact]
    public void NumberFetchTest()
    {
        var book = new ExcelBook("excel-sharp.xlsx");
        var sheet = book.GetSheet("Fetch");
        var models = sheet.Fetch<NumberFetchModel>("A16", x => new
        {
            x.Name,
            x.Byte,
            x.Char,
            x.Int16,
            x.UInt16,
            x.Int32,
            x.UInt32,
            x.Int64,
            x.UInt64,
            x.Single,
            x.Double,
            x.Decimal,
            x.Formula,
        });

        var valueModel = models[0];
        Assert.Equal(100, valueModel.Byte);
        Assert.Equal(100, valueModel.Char);
        Assert.Equal(100, valueModel.Int16);
        Assert.Equal(100, valueModel.UInt16);
        Assert.Equal(100, valueModel.Int32);
        Assert.Equal(100u, valueModel.UInt32);
        Assert.Equal(100, valueModel.Int64);
        Assert.Equal(100uL, valueModel.UInt64);
        Assert.Equal(100, valueModel.Single);
        Assert.Equal(100, valueModel.Double);
        Assert.Equal(100m, valueModel.Decimal);
        Assert.Equal(100, valueModel.Formula);

        var blankModel = models[1];
        Assert.Equal(0, blankModel.Byte);
        Assert.Equal(0, blankModel.Char);
        Assert.Equal(0, blankModel.Int16);
        Assert.Equal(0, blankModel.UInt16);
        Assert.Equal(0, blankModel.Int32);
        Assert.Equal(0u, blankModel.UInt32);
        Assert.Equal(0, blankModel.Int64);
        Assert.Equal(0uL, blankModel.UInt64);
        Assert.Equal(0, blankModel.Single);
        Assert.Equal(0, blankModel.Double);
        Assert.Equal(0, blankModel.Decimal);
        Assert.Equal(0, blankModel.Formula);
    }

    [Fact]
    public void NullableNumberFetchTest()
    {
        var book = new ExcelBook("excel-sharp.xlsx");
        var sheet = book.GetSheet("Fetch");
        var models = sheet.Fetch<NullableNumberFetchModel>("A21", x => new
        {
            x.Name,
            x.Byte,
            x.Char,
            x.Int16,
            x.UInt16,
            x.Int32,
            x.UInt32,
            x.Int64,
            x.UInt64,
            x.Single,
            x.Double,
            x.Decimal,
            x.Formula,
        });

        var value = models[0];
        Assert.Equal((byte)100, value.Byte);
        Assert.Equal((char)100, value.Char);
        Assert.Equal((short)100, value.Int16);
        Assert.Equal((ushort)100, value.UInt16);
        Assert.Equal(100, value.Int32);
        Assert.Equal(100u, value.UInt32);
        Assert.Equal(100, value.Int64);
        Assert.Equal(100uL, value.UInt64);
        Assert.Equal(100, value.Single);
        Assert.Equal(100, value.Double);
        Assert.Equal(100m, value.Decimal);
        Assert.Equal(100, value.Formula);

        var blank = models[1];
        Assert.Null(blank.Byte);
        Assert.Null(blank.Char);
        Assert.Null(blank.Int16);
        Assert.Null(blank.UInt16);
        Assert.Null(blank.Int32);
        Assert.Null(blank.UInt32);
        Assert.Null(blank.Int64);
        Assert.Null(blank.UInt64);
        Assert.Null(blank.Single);
        Assert.Null(blank.Double);
        Assert.Null(blank.Decimal);
        Assert.Null(blank.Formula);
    }

    [Fact]
    public void ArrayFetchTest()
    {
        var book = new ExcelBook("excel-sharp.xlsx");
        var sheet = book.GetSheet("Fetch");

        var models = sheet.Fetch<ArrayFetchModel>("A26", x => new
        {
            x.Name,
            x.Numbers,
            x.Total,
        });

        var valueModel = models[0];
        Assert.Equal([10, 20, 30, 40, 50, 60, 70, 80, 90, 100], valueModel.Numbers);
        Assert.Equal(550, valueModel.Total);

        var blankModel = models[1];
        Assert.Equal(new int?[10].Let(i => null), blankModel.Numbers);
        Assert.Null(blankModel.Total);
    }

    [Fact]
    public void MergedFetchTest()
    {
        var book = new ExcelBook("excel-sharp.xlsx");
        var sheet = book.GetSheet("Fetch");

        var models = sheet.Fetch<MergedFetchModel>("A31", x => new
        {
            x.Name,
            x.Number1,
            x.Number2,
        });

        var model0 = models[0];
        Assert.Equal(10, model0.Number1);
        Assert.Equal(20, model0.Number2);

        var model1 = models[1];
        Assert.Equal(10, model1.Number1);
        Assert.Equal(30, model1.Number2);
    }

}
