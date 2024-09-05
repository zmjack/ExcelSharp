using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using NStandard;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace ExcelSharp;

public class ExcelBook : IDisposable
{
    private readonly SpreadsheetDocument Document;
    private bool disposedValue;

    public ExcelBook(string path)
    {
        if (File.Exists(path)) Document = SpreadsheetDocument.Open(path, true);
        else Document = SpreadsheetDocument.Create(path, SpreadsheetDocumentType.Workbook);
    }

    private WorkbookPart WorkbookPart
    {
        get
        {
            var part = Document.WorkbookPart ?? Document.AddWorkbookPart();
            part.Workbook = new Workbook();
            return part;
        }
    }
    private SharedStringTablePart SharedStringTablePart
    {
        get
        {
            return Document.WorkbookPart.GetPartsOfType<SharedStringTablePart>().FirstOrDefault()
                ?? Document.WorkbookPart.AddNewPart<SharedStringTablePart>();
        }
    }
    private Sheets InnerSheets => WorkbookPart.Workbook.GetFirstChild<Sheets>() ?? Document.WorkbookPart.Workbook.AppendChild(new Sheets());
    private uint NewSheetId()
    {
        if (!InnerSheets.Elements<Sheet>().Any()) return 1;
        else return InnerSheets.Elements<Sheet>().Select(s => s.SheetId.Value).Max() + 1;
    }

    public IEnumerable<string> SharedStrings => SharedStringTablePart.SharedStringTable.Elements<SharedStringItem>().Select(x => x.InnerText);

    public int GetIndexOfSharedString(string text)
    {
        SharedStringTablePart.SharedStringTable ??= new SharedStringTable();

        int i = 0;
        foreach (var item in SharedStringTablePart.SharedStringTable.Elements<SharedStringItem>())
        {
            if (item.InnerText == text) return i;
            i++;
        }
        SharedStringTablePart.SharedStringTable.AppendChild(new SharedStringItem(new Text(text)));
        SharedStringTablePart.SharedStringTable.Save();

        return i;
    }

    public WorksheetPart CreateSheet(string sheetName)
    {
        var part = WorkbookPart.AddNewPart<WorksheetPart>();
        part.Worksheet = new Worksheet(new SheetData());
        part.Worksheet.Save();

        var relationshipId = WorkbookPart.GetIdOfPart(part);
        var sheet = new Sheet
        {
            Id = relationshipId,
            SheetId = NewSheetId(),
            Name = sheetName
        };
        InnerSheets.Append(sheet);
        WorkbookPart.Workbook.Save();

        return part;
    }

    public ExcelSheet GetSheet(string name)
    {
        return new ExcelSheet(this, InnerSheets.ChildElements.OfType<Sheet>().FirstOrDefault(x => x.Name == name));
    }

    protected virtual void Dispose(bool disposing)
    {
        if (!disposedValue)
        {
            if (disposing)
            {
                // TODO: dispose managed state (managed objects)
                Document.Dispose();
            }

            // TODO: free unmanaged resources (unmanaged objects) and override finalizer
            // TODO: set large fields to null
            disposedValue = true;
        }
    }

    // // TODO: override finalizer only if 'Dispose(bool disposing)' has code to free unmanaged resources
    // ~ExcelBook()
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
