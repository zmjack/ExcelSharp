using DocumentFormat.OpenXml.Spreadsheet;

namespace ExcelSharp;

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
