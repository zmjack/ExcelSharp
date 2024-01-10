namespace ExcelSharp.NPOI;

public partial class ExcelSheet
{
    public static int GetCol(string name) => ColumnIndex.GetIndex(name);
    public static string GetColName(int index) => ColumnIndex.GetName(index);

    private static int GetRow(string name) => int.Parse(name) - 1;
    private static string GetRowName(int row) => (row + 1).ToString();

}