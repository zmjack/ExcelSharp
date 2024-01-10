using ExcelSharp.NPOI;
using System;

namespace ExcelSharp.Test.Models;

public class NullableDateFetchModel
{
    [SheetColumn("Name")]
    public string Name { get; set; }

    [SheetColumn("DateTime")]
    public DateTime? Date { get; set; }

    [SheetColumn("DateTime (String)")]
    public DateTime? StringDate { get; set; }

    [SheetColumn("DateTime (Number)")]
    public DateTime? NumberDate { get; set; }

    [SheetColumn("DateTime (Formula)")]
    public DateTime? FormulaDate { get; set; }
}
