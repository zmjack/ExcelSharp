using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using NStandard;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace ExcelSharp
{
    public class ExcelBook : IDisposable
    {
        private readonly SpreadsheetDocument Document;

        public ExcelBook(string path)
        {
            if (File.Exists(path)) Document = SpreadsheetDocument.Open(path, true);
            else Document = SpreadsheetDocument.Create(path, SpreadsheetDocumentType.Workbook);
        }

        private WorkbookPart WorkbookPart => Document.WorkbookPart ?? Document.AddWorkbookPart().Then(x => x.Workbook = new Workbook());
        private SharedStringTablePart SharedStringTablePart => Document.WorkbookPart.GetPartsOfType<SharedStringTablePart>().FirstOrDefault() ?? Document.WorkbookPart.AddNewPart<SharedStringTablePart>();
        private Sheets InnerSheets => WorkbookPart.Workbook.GetFirstChild<Sheets>() ?? Document.WorkbookPart.Workbook.AppendChild(new Sheets());
        private uint NewSheetId()
        {
            if (InnerSheets.Elements<Sheet>().Count() == 0) return 1;
            else return InnerSheets.Elements<Sheet>().Select(s => s.SheetId.Value).Max() + 1;
        }

        public IEnumerable<string> SharedStrings => SharedStringTablePart.SharedStringTable.Elements<SharedStringItem>().Select(x => x.InnerText);

        public int GetIndexOfSharedString(string text)
        {
            SharedStringTablePart.SharedStringTable ??= new SharedStringTable();

            int i = 0;
            foreach (SharedStringItem item in SharedStringTablePart.SharedStringTable.Elements<SharedStringItem>())
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
            var sheet = new Sheet { Id = relationshipId, SheetId = NewSheetId(), Name = sheetName };
            InnerSheets.Append(sheet);
            WorkbookPart.Workbook.Save();

            return part;
        }

        public ExcelSheet Sheets(string name)
        {
            return new ExcelSheet(this, InnerSheets.ChildElements.OfType<Sheet>().FirstOrDefault(x => x.Name == name));
        }

        public void Dispose()
        {
            Document.Dispose();
        }
    }
}
