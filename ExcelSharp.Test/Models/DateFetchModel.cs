using System;

namespace ExcelSharp.Test.Models
{
    public class DateFetchModel
    {
        public string Name { get; set; }
        public DateTime Date { get; set; }
        public DateTime StringDate { get; set; }
        public DateTime NumberDate { get; set; }
        public DateTime FormulaDate { get; set; }
    }
}
