using NPOI.SS.UserModel;
using System;

namespace ExcelSharp.NPOI;

public interface IStylizable
{
    void SetCellStyle(ICellStyle style);
    void SetCStyle(CStyle style);
    void SetCStyle(Action<CStyleApplier> initApplier);
}
