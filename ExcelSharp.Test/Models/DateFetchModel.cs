using ExcelSharp.NPOI;
using System;

namespace ExcelSharp.Test.Models;

public class DateFetchModel
{
    [DataColumn("Name")]
    public string Name { get; set; }

    [DataColumn("DateTime")]
    public DateTime Date { get; set; }

    [DataColumn("DateTime (String)")]
    public DateTime StringDate { get; set; }

    [DataColumn("DateTime (Number)")]
    public DateTime NumberDate { get; set; }

    [DataColumn("DateTime (Formula)")]
    public DateTime FormulaDate { get; set; }
}
