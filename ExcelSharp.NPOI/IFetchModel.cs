namespace ExcelSharp.NPOI;

public interface IFetchModel<TValue>
{
    public TValue? Value { get; set; }
}
