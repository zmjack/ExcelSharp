using NStandard;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq.Expressions;
using System.Reflection;

namespace ExcelSharp.NPOI;

[DebuggerDisplay(@"{Header}: {Value}")]
public class Model2D<THeader, TValue> where THeader : new()
{
    public required THeader Header { get; set; }
    public required TValue? Value { get; set; }
}

internal class Model2DScope<THeader> : Scope<Model2DScope<THeader>> where THeader : new()
{
    public delegate THeader GetHeaderDelagate(string[] rowNames, string[] colNames);
    public GetHeaderDelagate GetHeader { get; set; }

    public Model2DScope()
    {
        var rowNames = Expression.Parameter(typeof(string[]));
        var colNames = Expression.Parameter(typeof(string[]));

        var codes = new List<Expression>();
        var header = Expression.Variable(typeof(THeader));

        codes.Add(Expression.Assign(header, Expression.New(typeof(THeader).GetConstructor([])!)));

        var props = typeof(THeader).GetProperties();
        foreach (var prop in props)
        {
            var row = prop.GetCustomAttribute<RowHeaderAttribute>();
            if (row is not null)
            {
                var setMethod = prop.GetSetMethod();
                if (setMethod is null) continue;

                var rowIndex = row.Index;
                codes.Add(
                    Expression.Call(
                        header,
                        setMethod,
                        Expression.ArrayAccess(rowNames, Expression.Constant(rowIndex, typeof(int)))
                    )
                );

                continue;
            }

            var col = prop.GetCustomAttribute<ColumnHeaderAttribute>();
            if (col is not null)
            {
                var setMethod = prop.GetSetMethod();
                if (setMethod is null) continue;

                var colIndex = col.Index;
                codes.Add(
                    Expression.Call(
                        header,
                        setMethod,
                        Expression.ArrayAccess(colNames, Expression.Constant(colIndex, typeof(int)))
                    )
                );

                continue;
            }
        }
        codes.Add(header);

        var body = Expression.Block([header], codes);
        var lambda = Expression.Lambda<GetHeaderDelagate>(body, [rowNames, colNames]);
        GetHeader = lambda.Compile();
    }
}
