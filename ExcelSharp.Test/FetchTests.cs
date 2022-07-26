using ExcelSharp.NPOI;
using ExcelSharp.Test.Models;
using NStandard;
using System;
using Xunit;

namespace ExcelSharp.Test
{
    public class FetchTests
    {
        [Fact]
        public void NoRowsTest()
        {
            var book = new ExcelBook("excel-sharp.xlsx");
            var sheet = book.GetSheet("Fetch");

            var models = sheet.Fetch<DateFetchModel>("A2", x => new { x.Name, x.Date, x.StringDate, x.NumberDate, x.FormulaDate });
            Assert.Empty(models);
        }

        [Fact]
        public void DateFetchTest()
        {
            var book = new ExcelBook("excel-sharp.xlsx");
            var sheet = book.GetSheet("Fetch");

            var models = sheet.Fetch<DateFetchModel>("A6", x => new { x.Name, x.Date, x.StringDate, x.NumberDate, x.FormulaDate });
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
            var models = sheet.Fetch<NullableDateFetchModel>("A11", x => new { x.Name, x.Date, x.StringDate, x.NumberDate, x.FormulaDate });
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
            var models = sheet.Fetch<NumberFetchModel>("A16", x => new { x.Name, x.Byte, x.Char, x.Int16, x.UInt16, x.Int32, x.UInt32, x.Int64, x.UInt64, x.Single, x.Double, x.Decimal, x.Formula });

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
            var models = sheet.Fetch<NullableNumberFetchModel>("A21", x => new { x.Name, x.Byte, x.Char, x.Int16, x.UInt16, x.Int32, x.UInt32, x.Int64, x.UInt64, x.Single, x.Double, x.Decimal, x.Formula });

            var valueModel = models[0];
            Assert.Equal((byte)100, valueModel.Byte);
            Assert.Equal((char)100, valueModel.Char);
            Assert.Equal((short)100, valueModel.Int16);
            Assert.Equal((ushort)100, valueModel.UInt16);
            Assert.Equal(100, valueModel.Int32);
            Assert.Equal(100u, valueModel.UInt32);
            Assert.Equal(100, valueModel.Int64);
            Assert.Equal(100uL, valueModel.UInt64);
            Assert.Equal(100, valueModel.Single);
            Assert.Equal(100, valueModel.Double);
            Assert.Equal(100m, valueModel.Decimal);
            Assert.Equal(100, valueModel.Formula);

            var blankModel = models[1];
            Assert.Null(blankModel.Byte);
            Assert.Null(blankModel.Char);
            Assert.Null(blankModel.Int16);
            Assert.Null(blankModel.UInt16);
            Assert.Null(blankModel.Int32);
            Assert.Null(blankModel.UInt32);
            Assert.Null(blankModel.Int64);
            Assert.Null(blankModel.UInt64);
            Assert.Null(blankModel.Single);
            Assert.Null(blankModel.Double);
            Assert.Null(blankModel.Decimal);
            Assert.Null(blankModel.Formula);
        }

        [Fact]
        public void ArrayFetchTest()
        {
            var book = new ExcelBook("excel-sharp.xlsx");
            var sheet = book.GetSheet("Fetch");

            var models = sheet.Fetch<ArrayFetchModel>("A26", x => new { x.Name, x.Numbers, x.Total });

            var valueModel = models[0];
            Assert.Equal(new int?[] { 10, 20, 30, 40, 50, 60, 70, 80, 90, 100 }, valueModel.Numbers);
            Assert.Equal(550, valueModel.Total);

            var blankModel = models[1];
            Assert.Equal(new int?[10].Let(() => null), blankModel.Numbers);
            Assert.Null(blankModel.Total);
        }

    }
}
