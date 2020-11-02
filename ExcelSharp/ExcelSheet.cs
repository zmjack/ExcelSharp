using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Text;

namespace ExcelSharp
{
    public class ExcelSheet
    {
        protected ExcelBook Book;
        protected Sheet Sheet;

        public ExcelSheet(ExcelBook book, Sheet sheet)
        {
            Book = book;
            Sheet = sheet;
        }
    }
}
