using System;
using System.Collections.Generic;
using System.Text;
using System.Text.RegularExpressions;

namespace ExcelSharp
{
    public struct Cursor
    {
        public int Row { get; set; }
        public int Col { get; set; }

        public Cursor(string cellName)
        {
            var regex = new Regex(@"^([A-Z]+)(\d+)$");
            var match = regex.Match(cellName);
            if (match.Success)
            {
                Row = int.Parse(match.Groups[2].Value) - 1;
                Col = LetterSequence.GetNumber(match.Groups[1].Value);
            }
            else throw new FormatException($"Illegal cell format：{cellName}。");
        }

        public Cursor(int row, int col)
        {
            Row = row;
            Col = col;
        }

        public Cursor((int row, int col) cell)
        {
            Row = cell.row;
            Col = cell.col;
        }

        public override string ToString() => $"{LetterSequence.GetLetter(Col)}{Row + 1}";

        public static implicit operator Cursor(string cellName) => new Cursor(cellName);
        public static implicit operator Cursor((int row, int col) cell) => new Cursor(cell);
    }
}
